Attribute VB_Name = "Module1"

Public Declare Sub InitCommonControls Lib "comctl32.dll" () 'XP效果

Public stRfpc95qtytxt As String 'ccs add 20161207 95fpc手动录入的开工单数量


Public Type WOTRANSTMP

strWorkOrderID As String
strLotID As String
    
End Type

Public Type CODE37

strCus As String
strDev As String
strBline As String
strCode  As String
strStatus As String
        
End Type

Public Type WorkOrderData
WO_TYPE As String            '工单类型
WO_FIELD As String           '工单用途
WO_FLAG As String            '工单状态
WO_ID   As String            '工单号
WO_DEPT As String            '工单部门
WO_CREATOR As String         '工单创建人员
WO_CREATE_TIME As String     '工单建立时间
WO_NPI_OWNER  As String      '工单物料负责人
WO_START_DATE As String      '工单开始日期
WO_END_DATE  As String       '工单结束日期
WO_GROSS_QTY As Long         '工单总DIE数量
WO_CUST_CODE As String       '工单客户代码
WO_CUST_PN  As String        '工单客户机种
WO_HT_PN    As String        '工单厂内机种
WO_HT_PRODUCTNAME As String  '工单成品料号
WO_LOT_ID  As String         '工单客户批号
WAFER_NO  As String          '客户WAFER序号
wafer_id  As String          '客户LOTWAFERID
WAFER_GROSS_QTY As Long      'WAFER总DIE数量
WAFER_GOOD_QTY As Long       'WAFER良品DIE数量
WAFER_NG_QTY  As Long        'WAFER不良品DIE数量
WAFER_BONDED As String       'WAFER保税/非保税
WAFER_MARKING_CODE As String 'WAFER打标码
WAFER_SEC_CODE As String     'WAFER二级代码
WAFER_JOB_ID As String       'WAFER_37JOBID
End Type

Public Type DN_REAL_INFO
dn        As String
MPN       As String
CPN       As String
PO        As String
QTY       As Long
REELS     As Long
BOXS      As Long
CARTONS   As Long
LABELTYPE As String
ADDRESS1    As String
ADDRESS2    As String
FAB_SITE As String
ASSEMBLY_SITE As String
TEST_SITE As String
End Type

Public Type DN_SCAN_INFO
dn As String
JOBID As String
REELID As String
REELQTY As Long
LAST_OP_NO As Integer
LAST_IP_NO As Integer
SCAN_FLAG As String
CHECK_FLAG As String
QBOXNO As String
RID As String
bid As String
CID As String
KID As String
DATECODE As String
SEQ As Long
        
End Type

Public Type SHIP_WO
WO_ID As String         ' 工单号
WO_STATUS As String     ' 工单状态

End Type




Public Type SHIP_PLAN

PLAN_ID As String
CUSTOMER As String
CUST_PART As String
PRODUCT_NAME As String
PRODUCT_ID As String
GROSS_DIE As Long
BAD_FLAG As String
SHIP_AD As String
SHIP_TYPE As String
SHIP_CUST As String
PLAN_DATE As String
ALARM_TIME As String
APPEROVER As String
APPROV_TIME As String
LAST_UPDATE_TIME As String
LAST_UPDATE_BY As String
SHIP_FLAG As String
SHIP_ORDER As String
REMARK1 As String
REMARK2 As String
REMARK3 As String
REMARK4 As String
REMARK5 As String
wafer_id As String
Lot_id As String
PO_NUM As String
SHOP_ORDER As String
TOTAL_DIES As Long
GOOD_DIES As Long
QBOXNO As String
PLAN_ITEM As Integer




        

End Type

Public Type CUSSHIPADDRESS
CUSTOMER As String
SHIP_TO As String
SHIPPER As String
SOLD_TO As String
BILL_TO As String
SHIP_TO_AD As String
SOLD_BY As String
PAYMENT_TERMS As String
CURRENCY As String
BANKINFO As String
TK As String
PO As String
SHIPPER_PACK As String
        
End Type

Public Type ty37PO
CUSTCODE As String
HTPN As String
HTPRODUCT As String
PCES As Integer
WAFERLIST As String
DIEQTY As Long
PO As String
CUSTPNNAME As String
CUSTPN  As String
CUSJOBID As String
PRODUCTIONORDER As String

End Type

Public Type SH103_REEL_INFO
SH103_DN As String
SH103_REEL_CODE As String
SH103_LOT_NO As String
SH103_H As String
PACKING_DATE_10 As String
IN_GOOD_DIE As String
Customer_Device As String
SH103_QBOXID As String
SH103_QBOXSEQ As String
SH103_QBOXWEIGHT As String
HT_PN As String
End Type

Public Type SH103_QBOX_INFO
SH103_QBOXID As String
SH103_QBOXWEIGHT As String
SH103_QBOXSEQ As String
End Type

Public Type tyWO
CUSTOMER_CODE As String
Customer_Device As String
Fab_Device As String
HT_DEVICE As String
ITEM As String
po_no As String
SUPPLIER As String
SHIP_TO As String
WAFER_VERSION As String
MARKING_CODE As String
WO_DATE As String
Lot_id As String
wafer_id As String
lot_wafer_id As String
GOOD_DIES_PCS As Long
GROSS_DIES_PCS As Long
NG_DIES_PCS As Long
WO_NO As String
TRADE_TYPE As String
TAX_TYPE As String
PO_ITEM As String
UPLOADER_NAME As String
UPLOADER_ID As String
UPLOAD_DATE As String
DATA1 As String
DATA2 As String
DATA3 As String
DATA4 As String
DATA5 As String
REMARK As String
MARKING_CODE2 As String
MARKING_CODE3 As String

End Type

Public Type B2B
EVENT_DATE As String
EVENT As String
OVT_COMPANY As String
OVT_ORG As String
SUB_NAME As String
Stage As String
WAFER_LOT As String
OVT_JOB As String
SUB_LOT As String
Fab_Device As String
SOURCE_DEVICE As String
TARGET_DEVICE As String
WAFER_QTY As String
WAFER_DIE As String
wafer_id As String
PO As String
PO_RELEASE As String
OPERATION_CODE As String
OPERATION_DESCRIPTION As String
RECEIVE_QTY As String
JOB_ISSUE_QTY As String
START_QTY As String
COMPLETED_QTY As String
GRADE_RECORD As String
HOLD_CODE As String
HOLD_QTY As String
SCRAP_CODE As String
SCRAP_QTY As String
SCRAP_WAFER_ID As String
Priority As String
DATECODE As String
ENG_NO As String
NEXT_SUB_LOCATION As String
LOT_TYPE As String
E_SOD As String
TEST_PROGRAM As String
RMA_NO As String
BILLING_INVOICE As String
SHIPPING_INVOICE As String
SHIPPING_DESTINATION As String
JOB_FLAG As String
REMARK As String
SO As String
SO_LINE As String
GROSS_WEIGHT As String
NET_WEIGHT As String
FORWARDER As String
AIR_WAYBILL As String
CARTON_QTY As String
        
End Type

Public Type US026INVOICE
SHIPPING_INVOICE As String
SHIPPING_DESTINATION As String
GROSS_WEIGHT As String
NET_WEIGHT As String
FORWAROER As String
AIR_WAYBILL As String
CARTON_QTY As String
BILLING_INVOICE As String
FLAG As String
    
End Type

Public Type OIRecord
   id As Long
   PoNum As String
   PoItem As String
   LOTID As String
   MPN As String
   MPNDec As String
   
   WaferQTY As Integer
   DIEQTY As Long
   DESIGNID As String
   CountryFab As String
   ImageRev As String
   
   FFacility As String
   MarkId As String
   LotPriority As String
   FilmApld As String
   Ship260 As String
   
   ShipLevel As String
   MicMaterial As String
   ShipSite As String
   LotStatus As String
   customerName As String
   
   FLAG As String
   
   CreateBy As String
   CreateDate As Date
   
End Type

Public Type ForeCastRecord
  id As Long
  TYPENAME As String
  DemandType   As String
  StartPartId  As String
  OutPartId  As String
  OutDate As Date
  outQty As Long
  WorkWeek As String
  SiteId As String
  StageId As String
  Ctg As String
  Pti2 As String
  Comments As String
  CreateDate As Date
  Site  As String
  Stage As String
  PartId As String
  ProdPartId  As String
  ConsItem As String
  StartingWeek  As Date
  NextSite As String
  MfgAreaCd  As String
  OracleLocCd As String
  PTI As String
  PkgCd As String
  PkgGrpCd As String
  SchComments As String
  QTY  As Long
  IT  As String
  OnHand As String
   FLAG As String
   QtechCreateBy As String
   QtechCreateDate As Date
   
End Type



Public Type MapRecord
   SUBSTRATEID As String
   substratetype As String
   LOTID As String
   PRODUCTID As String
   CreateDate As String
   MicronLotId As String
   PASSBINCOUNT As Long
   FailBinCount As Long
   TotalQty As Long
   filename As String
End Type

Public Type GCHeader
   po_no As String
   SUPPLIER As String
   ShipTo As String
   Fab_Device As String
   FAB_Device2 As String
   taxTemp As String
   Customer_Device As String
   GC_Version As String
   GC_Date As String
   Lot_id As String
   WO_NO As String
   Created_By As String
   remarkTemp As String
   DATE_CODE As String
   Marking_Lot_ID1 As String
   Marking_Lot_ID2 As String
   Ship_Out As String
   TradeType As String
   TAXTYPE As String
   Attri01 As String
   Attri02 As String
   Attri03 As String
   Attri04 As String
   Attri05 As String
   Veqdatecode As String  'ccs add 2016-07-05
   Coo  As String
   Level As String  'tw add 20171017
   SecondFlag As String 'tw add 20180115
   LotProperty As String
   LotOwner As String
   Telephone As String
   Flow As String   'Merry add 20200318
   htdevice As String   'Merry add 20200318
End Type

Public Type WoWafer
    chech As String
    CustName As String
    lot As String
    WAFER As String
    id As String
    PO As String
    device As String
    gooddie As Long
    ngdie As Long
    idd As Long
  
    
End Type

Public Type EQISHeader
   Created_Datetime As String
   Vendor As String
   Process As String
   ORDERTYPE As String
   ESR_No As String
   
   AssemblyDateCode As String
   po_no As String
   WO_NO As String
   WorkOrder_PartNo As String
   device As String
   
   WaferQTY As Double
   AssyQty As Long
   PACKAGE As String
   FabLotNo As String
   TSM_A As String
   
   TSM_B As String
   TSM_C As String
   TSM_D As String
   BondingDiagram As String
   CompleteLotno As String
   
   Remarks As String
   MarketingPartNumber As String
   SPA As String
   DATECODE As String
   DieID As String
   
   LabelFormat As String
   waferid As String
   SPADESC As String
   Attention As String
   CompanyName As String
   
   Created_By As String
   
   SUBCONPO As String
   ITEM As String
   Quantity As String
   devicetemp As String
   SPATemp As String
   CSD As String
   lot As String
   DATECODE1 As String
   DELIVERYNAME As String
   DELIVERYADDRESS As String
   WAREHOUSE As String
   LOCATION As String
   MODEOFDELIVERY As String
   dateCodeTemp As String
   SO As String
   CARRIERNOTES As String
   LINE As String
   SCHEDULELINE As String
   CUSTPN As String
   COUNTRYANDNAMEOFDISTRIBUTOR As String
   CUSTOMER As String
   CUSTOMERPO As String
End Type



Public Type GCDetail
   ITEM As String
   Marking_Lot_ID As String
   Lot_id As String
   wafer_id As String
   Good_Die_Qty As Long
   NG_Die_Qty As Long
   REMARK As String
   
   strWaferID As String
   strMarkingCode As String
   strLotID As String
   strWaferNo As String
   lGoodDies As Long
   lGrossDies As Long
   strRemark As String
   TAX_TYPE As String
   
End Type




Public Type SemtechPOHeader

PurchaseOrderNo As String
ProductionOrderNo As String
Version As String
DATE As String
CURRENCY As String

ShippingAddress As String
TermsPayment As String
TermsDelivery As String
JOBNO As String
FreightCarrier As String
ITEM As Integer
MaterialDes As String
LotNO As String
Quantity As Long
UM As String
DelDate As String
UnitPrice As String
NetAmount As Long
YourMaterialNumber As String
TypeService As String
MfgPlant As String
ReceivingPlant As String
PartNumber As String
WAFERLOT As String
WaferFAB As String
WaferREV As String
KeyStr As String
waferIDList As String
waferid As String
po1lot As String
Plant As String
po1lot2 As String
id As Long
customerName As String
FLAG As String
QTECH_CREATED_BY As String
QTECH_CREATED_DATE As Date
BagNo As String
DATECODE As String
LotNumber As String
POPrice  As Double
ItemLineText As String
FabSite As String
AssemblySite As String
TestSite As String
PPR As String

' 37新加字段
LOTID As String
JOBID As String
JobID_2 As String
WaferNO As String
WaferID_2 As String
Price As String
fab_conv_id As String
BondOrNot As String
LOTQTY As Integer
LOTDIES As Long
WAFERNOLIST As String
HTPN As String
HTPRODUCT As String
End Type





Public Type SemtechFabDetail
   
  DeviceName As String
  Batch As String
  WF As Integer
  Price As Double
  CURRENCY As String
  ShippedDt As Date
  PurchaseNo As String
  PurchaseOrderLineItem As String
  Invoice As String
  MAWBNumber As String
  Destination As String
  wafer_id As String
  FLAG As String
  QTECH_CREATED_BY As String
  QTECH_CREATED_DATE As Date
  

   ITEM As String
   Marking_Lot_ID As String
   Lot_id As String
   Good_Die_Qty As Long
   NG_Die_Qty As Long
   REMARK As String
   custom_Temp As String
   id As Long
   
   PoBatch As String
   
   

End Type



'工单Header
Public Type BillHeader
   id As Long
    ORDERNAME As String
    ORDERTYPE  As String
    EVENTTYPE As String
    ERPUSER As String
    product As String
    QTY As Long
    RequestDate As Date
    ERPCREATEDATE  As Date
    PLANSTARTDATE As Date
    PLANENDDATE As Date
    CUSTOMER As String
    SALESORDER As String
    MODIFYFLAG As Integer
    CustomerERPN  As String
    FABFACILITY As String
    IMAGERREV As String
    DESIGNID  As String
    MLEVEL235 As String
    MLEVEL260 As String
    NGFLAG  As Integer
    PARA1 As String
    PARA2 As String
    PARA3 As String
    PARA4 As String
    PARA5 As String
    PARA6 As String
    PARA8 As String
    PARA9 As String
    PARA10 As String
    
    PROTECTIVE_FILM_APLD As String
    Lot_Stauts As String
    MPN As String
    sPri As String
    sLotType As String
    
    
End Type


Public Type NpiProduct
    id As Long
    CUSTOMERSHORTNAME As String
    qtechPTNo As String
    QtechPTNo2 As String
    CustomerPTNo1 As String
    CustomerPTNo2 As String
    CustomerDieQty As String
    QtechDieQty As String
    XiangSu As String
    UsedArea As String
    StruckStr1 As String
    StruckStr2 As String
    StruckStr3 As String
    FLAG As String
    STDate As String
    TTDate As String
    PTDate As String
    CreateBy As String
    
    FzFreeUSD As String
    TestFreeUSD As String
    FzFreeRMB As String
    TestFreeRMB As String
    NreFree As String
    NreMethod As String
    UpdatePrice2 As String
    UpdatePrice1 As String
    CustomerPTNo3 As String
    CustomerPTNo4 As String
     CustomerPTNo5 As String
    CustomerPTNo6 As String
    '
    CustomerPTNo7 As String
    CustomerPTNo8 As String
    '
    PKG As String
    residual As String
    secondCode As String 'CCS ADD 20160717
    MARKINGCODE As String ' By Tony 20170814
    ProducEng   As String    ' tw  add   工程量产
    MAPPING As String
    Owner As String
    WaferPN As String
    
    
    
    
End Type



Public Type POPrice
    id As Long
    customerName As String
    CUSTOMERSHORTNAME As String
    PONo As String
    PODATE As Date
    POType As String
    pt As String
    QTY As Long
    peaseQty As Long
    Price As String
    unit As String
    File As String
    CreateBy As String
    custAA As String
    bj As String
    SingWafer As String
    SingDie As String
    DIE_PRICE As String
End Type





Public Type Baofei
    id As Long
    putInDept As String
    waferid As String
    LOTID As String
    gDDie As Long
    ngdie As Long
    customerPTNo As String
    qtechPTNo As String
    productName As String
    err_date As String
    putIn_date As String
    FLAG As String
    CreateBy As String
    TYPENAME As String
    
End Type



'工单Header
Public Type WLOBillHeader
    id As Long
    ORDERNAME As String
    ORDERTYPE  As String
    EVENTTYPE As String
    ERPUSER As String
    product As String
    CUSTOMER As String
    QTY As Long
    PieceQty As Integer
    RequestDate As Date
    ERPCREATEDATE  As Date
    PLANSTARTDATE As Date
    PLANENDDATE As Date
 
End Type

Public Type WLOBillDetail
    ORDERNAME As String
    waferid As String
    DIEQTY As Long
    WAFERSEQUENCE As Long

End Type



'工单Detail
Public Type BillDetail
    ORDERNAME As String
    waferid As String
    DIEQTY As Long
    FGDIEQTY As Long
    WAFERLOT As String
    WAFERSEQUENCE As Long
    MARKINGCODE  As String

End Type


Public Type VTData  'VT

SHIPDATETemp As Date
StockNoTemp As String
DeliveryNoTemp As String
CustDeviceTemp As String
CUSTLOTTemp As String
waferIdTemp As String
WLCSPDeviceTemp As String
WLCSPLOTTemp As String
goodDieQtyTemp As Long
ngDieQtyTemp As Long
PackingLOTNoTemp As String
TTLTemp As String
WaferQtyInTemp As String
BatchTemp As String
SAPCodeTemp As String
WorkWeekTemp As String
CartonNoTemp As String
NetWeightTemp As String
GrossWeightTemp As String
remarkTemp As String
Created_ByTemp As String
type As String
htdevice As String

End Type

Public Type ShipSideData

Created_ByTemp As String
CustomerCode As String
GULFDeviceName As String
GULFLotID As String
WaferQTY As String
ShipTo As String
        
End Type


Public Type ShippingData   'Shipping

idTemp As Long
notemp As Long
SubConPOTemp As String
itemTemp As String
QuantityTemp As Long
devicetemp As String
SPATemp As String
CSDTemp As String
lottemp As String
DateCode1Temp As String
DeliveryNameTemp As String
DeliveryAddressTemp As String
WarehouseTemp As String
LocationTemp As String
ModeOfDeliveryTemp As String
dateCodeTemp As String
soTemp As String
CarrierNotesTemp As String
lineTemp As String
ScheduleLineTemp As String
CustPNTemp As String

CountryDistributorTemp As String
customerTemp As String
customerPoTemp As String
CreatedByTemp As String




End Type


Public Type SpGR  'Special GR

PoNum As String
PoItem As String
PoLotID As String
PreviousMtrl As String
BatchID As String
MtrlNum As String
MtrlDesc As String
MtrlNumMtr As String
ConsumedQty As Long
RejectQty As Long
CurrentWaferQty As Double
DATECODE As String
TstProgram As String
CreatedDate As String
CreatedTime As String

End Type




Public SumCount As Integer  'BC 成功上传数量
Public BCResultFlag As Boolean


Public Cnn As New ADODB.Connection
Public CnnERPInt As New ADODB.Connection

Public cmd                  As New ADODB.Command
Public rs                   As New ADODB.Recordset
Dim updateRS                As New ADODB.Recordset
Dim updateRSHeader          As New ADODB.Recordset

Public CnnSPC  As New ADODB.Connection

Public C_SysName As String
Public gUserName As String
Public gUserRealName As String
Public openWOID As String
Public gCnnState As Integer

Public adoCmd As ADODB.Command
Dim adoprm1 As ADODB.Parameter
Dim adoprm2 As ADODB.Parameter
Dim adoPrm3 As ADODB.Parameter
Dim adoPrm4 As ADODB.Parameter

Public bomProductTemp As String

Dim billLotTemp     As New ADODB.Recordset

'Public Const g_Path = "Z:\others"
Public Const g_Path = "C:\others"
Public Const g_PathNewOrder = "C:\HT_OrderLog"

Public Const g_Path37 = "C:\SemTechReport"


'Public Const g_Path_GR = "W:\GR2014"

Public Const g_Path_GR = "C:\GRList"


Public woSendTemp As String

Public upLoadWoFile As Boolean

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpas_FileName As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpas_ReportTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type


Public Const MAX_PATH = 260

Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long


Public Sub AddOIRecord(rpTemp As OIRecord)
'手工建立OI，提交
Dim cmdStr As String
Dim cmdStr2 As String

On Error GoTo DealError
        
Cnn.BeginTrans
                              
cmdStr = " insert into CustomerOItbl_test(id,Po_Num,Po_Item,source_batch_id,mpn,mpn_desc," & _
         "   current_wafer_qty,Die_Qty,design_id,country_of_fab,imager_customer_rev,fabrication_facility," & _
         "  encoded_mark_id, lot_priority,protective_film_apld,shipping_mst_260,shipping_mst_level,micron_material," & _
         "   ship_site,lot_status, customershortname,Flag ,qtech_created_by, qtech_created_date) Values " & _
         "  (" & rpTemp.id & ",'" & rpTemp.PoNum & "','" & rpTemp.PoItem & "','" & rpTemp.LOTID & "','" & rpTemp.MPN & "','" & rpTemp.MPNDec & "'," & _
         " " & rpTemp.WaferQTY & "," & rpTemp.DIEQTY & ",'" & rpTemp.DESIGNID & "','" & rpTemp.CountryFab & "','" & rpTemp.ImageRev & "','" & rpTemp.FFacility & "'," & _
         " '" & rpTemp.MarkId & "','" & rpTemp.LotPriority & "','" & rpTemp.FilmApld & "','" & rpTemp.Ship260 & "','" & rpTemp.ShipLevel & "','" & rpTemp.MicMaterial & "'," & _
         " '" & rpTemp.ShipSite & "','" & rpTemp.LotStatus & "','" & rpTemp.customerName & "', '" & rpTemp.FLAG & "','" & rpTemp.CreateBy & "', sysdate)"
               
 cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerOI] (id,Po_Num,Po_Item,source_batch_id,mpn,mpn_desc," & _
         "   current_wafer_qty,Die_Qty,design_id,country_of_fab,imager_customer_rev," & _
         "   fabrication_facility,encoded_mark_id,lot_priority,protective_film_apld,shipping_mst_260," & _
         "   shipping_mst_level,micron_material,ship_site,lot_status, customershortname," & _
         "   Flag , qtech_created_by, qtech_created_date) Values " & _
         "  (" & rpTemp.id & ",'" & rpTemp.PoNum & "','" & rpTemp.PoItem & "','" & rpTemp.LOTID & "','" & rpTemp.MPN & "','" & rpTemp.MPNDec & "'," & _
         " " & rpTemp.WaferQTY & "," & rpTemp.DIEQTY & ",'" & rpTemp.DESIGNID & "','" & rpTemp.CountryFab & "','" & rpTemp.ImageRev & "'," & _
         " '" & rpTemp.FFacility & "','" & rpTemp.MarkId & "','" & rpTemp.LotPriority & "','" & rpTemp.FilmApld & "','" & rpTemp.Ship260 & "'," & _
         " '" & rpTemp.ShipLevel & "','" & rpTemp.MicMaterial & "','" & rpTemp.ShipSite & "','" & rpTemp.LotStatus & "','" & rpTemp.customerName & "'," & _
         " '" & rpTemp.FLAG & "','" & rpTemp.CreateBy & "', GETDATE() )"
                              
 AddSql (cmdStr)
 AddSql2 (cmdStr2)
 
 Cnn.CommitTrans


  
MsgBox "已成功提交!", vbInformation, "提示"

Exit Sub
DealError:

Cnn.RollbackTrans
 
 MsgBox "保存失败!", vbInformation, "提示"
 
End Sub

Public Sub GetOracleConnection()
Dim strCnn As String

strCnn = "Provider=OraOLEDB.Oracle.1;Password=KsMesDB_ht89;Persist Security Info=True;User ID=insiteqt2;Data Source=TESTMES"

On Error Resume Next
While CnnSPC.State = 0

    CnnSPC.Open strCnn
    CnnSPC.CursorLocation = adUseClient
    CnnSPC.CommandTimeout = 600
    
    If CnnSPC.State <> 0 Then
        GoTo EndCon
    End If

Wend
EndCon:

         
End Sub


Public Sub AddMap(mapTemp As MapRecord, custNameTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008

On Error GoTo DealError

'导入Mapping
Cnn.BeginTrans
                  
cmdStr = "insert into mappingDataTest (id,SubstrateId,SubstrateType,LotId,ProductId,CreateDate,MicronLotId,PassBinCount,FailBinCount, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE,customershortname,FileName) Values" & _
          " ( mappingData_SEQ.Nextval, '" & mapTemp.SUBSTRATEID & "', '" & mapTemp.substratetype & "', '" & mapTemp.LOTID & "', '" & mapTemp.PRODUCTID & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PASSBINCOUNT & "," & mapTemp.FailBinCount & ",'Y','Auto',sysdate,'" & custNameTemp & "','" & mapTemp.filename & "')"
                      
                      
                      
cmdStr2 = "INSERT INTO [ERPBASE].[dbo].[tblmappingData]([SUBSTRATEID],[SUBSTRATETYPE],[LOTID],[PRODUCTID],[CREATEDATE],[MICRONLOTID],[PASSBINCOUNT],[FAILBINCOUNT],[FLAG],[QTECH_CREATED_BY],[QTECH_CREATED_DATE],[CUSTOMERSHORTNAME] ,FileName) Values" & _
          " ( '" & mapTemp.SUBSTRATEID & "', '" & mapTemp.substratetype & "', '" & mapTemp.LOTID & "', '" & mapTemp.PRODUCTID & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PASSBINCOUNT & "," & mapTemp.FailBinCount & ",'Y','Auto',GETDATE(),'" & custNameTemp & "','" & mapTemp.filename & "')"
                      
If AddSql(cmdStr) = 0 Then
    MsgBox mapTemp.SUBSTRATEID & ":没有成功插入Mapping信息, 请联系IT", vbCritical, "警告"
    Exit Sub

End If
 
AddSql2 (cmdStr2)
 
 
 Cnn.CommitTrans
 
SumCount = SumCount + 1
   
 Exit Sub
 
DealError:

' SumCount = SumCount - 1

Cnn.RollbackTrans

End Sub

Public Function IsWaferID_OnWorking(strWaferID As String) As Boolean

    Dim strSql As String
    Dim rs As New ADODB.Recordset

    strSql = "select * from ib_waferlist where waferid = '" & strWaferID & "'"

    If rs.State = adStateOpen Then rs.Close

    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs.EOF Then
        IsWaferID_OnWorking = True
    Else
        IsWaferID_OnWorking = False
    End If
    
    rs.Close
End Function

Public Function wowaferdet(wafertemp As WoWafer) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select * from ib_waferlist a  where a.waferid = '" & wafertemp.WAFER & "' "

slectResult = QueryStr(cmdStr)
wowaferdet = slectResult
End Function
Public Function wolotdet(wafertemp As WoWafer) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select * from mappingdatatest a where a.filename = '" & wafertemp.idd & "' "

slectResult = QueryStr(cmdStr)
wolotdet = slectResult
End Function

Public Sub dellot(wafertemp As WoWafer)

 sqlTemp = "delete  from customeroitbl_test where id = '" & wafertemp.idd & "' "
              
 sqlTemp1 = "delete  from erpbase..tblCustomerOI where id = '" & wafertemp.idd & "' "
           
 AddSql (sqlTemp)
 AddSql2 (sqlTemp1)

End Sub
Public Sub delwafer(wafertemp As WoWafer)

 sqlTemp = "delete  from mappingdatatest where substrateid = '" & wafertemp.WAFER & "' "
              
 sqlTemp1 = "delete  from erpbase..tblmappingData where substrateid = '" & wafertemp.WAFER & "' "
           
 AddSql (sqlTemp)
 AddSql2 (sqlTemp1)
 
End Sub


Public Sub AddMap56(mapTemp As MapRecord, custNameTemp As String, waferid As Integer)
Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008

On Error GoTo DealError

'导入Mapping
Cnn.BeginTrans
                  
cmdStr = "insert into mappingDataTest (id,SubstrateId,SubstrateType,LotId,ProductId,CreateDate,MicronLotId,PassBinCount,FailBinCount, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE,customershortname,FileName,wafer_id) Values" & _
          " ( mappingData_SEQ.Nextval, '" & mapTemp.SUBSTRATEID & "', '" & mapTemp.substratetype & "', '" & mapTemp.LOTID & "', '" & mapTemp.PRODUCTID & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PASSBINCOUNT & "," & mapTemp.FailBinCount & ",'Y','Auto',sysdate,'" & custNameTemp & "','" & mapTemp.filename & "'," & waferid & ")"
                      
                      
cmdStr2 = "INSERT INTO [ERPBASE].[dbo].[tblmappingData]([SUBSTRATEID],[SUBSTRATETYPE],[LOTID],[PRODUCTID],[CREATEDATE],[MICRONLOTID],[PASSBINCOUNT],[FAILBINCOUNT],[FLAG],[QTECH_CREATED_BY],[QTECH_CREATED_DATE],[CUSTOMERSHORTNAME] ,FileName,wafer_id) Values" & _
          " ( '" & mapTemp.SUBSTRATEID & "', '" & mapTemp.substratetype & "', '" & mapTemp.LOTID & "', '" & mapTemp.PRODUCTID & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PASSBINCOUNT & "," & mapTemp.FailBinCount & ",'Y','Auto',GETDATE(),'" & custNameTemp & "','" & mapTemp.filename & "'," & waferid & ")"
                      
AddSql (cmdStr)
 
AddSql2 (cmdStr2)
 
 
 Cnn.CommitTrans
 
SumCount = SumCount + 1
   
 Exit Sub
 
DealError:

' SumCount = SumCount - 1

Cnn.RollbackTrans

End Sub

Public Sub AddMap95(mapTemp As MapRecord, custNameTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008

On Error GoTo DealError

'导入Mapping
Cnn.BeginTrans
                  
cmdStr = "update mappingDataTest set PASSBINCOUNT='" & mapTemp.PASSBINCOUNT & "',FAILBINCOUNT= '" & mapTemp.FailBinCount & "' where SUBSTRATEID= '" & mapTemp.SUBSTRATEID & "'  and CUSTOMERSHORTNAME='" & custNameTemp & "'"
                                       
                      
cmdStr2 = "update [ERPBASE].[dbo].[tblmappingData] set PASSBINCOUNT='" & mapTemp.PASSBINCOUNT & "',FAILBINCOUNT= '" & mapTemp.FailBinCount & "' where SUBSTRATEID= '" & mapTemp.SUBSTRATEID & "' and CUSTOMERSHORTNAME='" & custNameTemp & "'"
          
      
AddSql (cmdStr)
 
AddSql2 (cmdStr2)
 
 
 Cnn.CommitTrans
 
SumCount = SumCount + 1
   
 Exit Sub
 
DealError:

' SumCount = SumCount - 1

Cnn.RollbackTrans

End Sub

Public Sub AddTSVMap(mapTemp As MapRecord, custNameTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008
Dim cmdStr3 As String
Dim cmd As New ADODB.Command

On Error GoTo DealError

'2014-11-13 jiayun add table TSV_Mapping_tbl

'导入Mapping
Cnn.BeginTrans
                  
cmdStr = "insert into TSV_Mapping_tbl (id,SubstrateId,SubstrateType,LotId,ProductId,CreateDate,MicronLotId,PassBinCount,FailBinCount, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE,customershortname,FileName) Values" & _
          " ( mappingData_SEQ.Nextval, '" & mapTemp.SUBSTRATEID & "', '" & mapTemp.substratetype & "', '" & mapTemp.LOTID & "', '" & mapTemp.PRODUCTID & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PASSBINCOUNT & "," & mapTemp.FailBinCount & ",'Y','Auto',sysdate,'" & custNameTemp & "','" & mapTemp.filename & "')"
                      
                      
                      
cmdStr2 = "INSERT INTO [ERPBASE].[dbo].[tblmappingDataCus]([SUBSTRATEID],[SUBSTRATETYPE],[LOTID],[PRODUCTID],[CREATEDATE],[MICRONLOTID],[PASSBINCOUNT],[FAILBINCOUNT],[FLAG],[QTECH_CREATED_BY],[QTECH_CREATED_DATE],[CUSTOMERSHORTNAME] ,FileName) Values" & _
          " ( '" & mapTemp.SUBSTRATEID & "', '" & mapTemp.substratetype & "', '" & mapTemp.LOTID & "', '" & mapTemp.PRODUCTID & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PASSBINCOUNT & "," & mapTemp.FailBinCount & ",'Y','Auto',GETDATE(),'" & custNameTemp & "','" & mapTemp.filename & "')"
          
          
                      


If AddSql(cmdStr) = 0 Then
    MsgBox "没有成功插入Mapping信息, 请联系IT", vbCritical, "警告"
    Exit Sub
End If


 
AddSql2 (cmdStr2)

 cmdStr3 = "UPDATE mappingdatatest a " & _
   "SET a.QTECH_LASTUPDATE_BY = '" & gUserRealName & "',a.QTECH_LASTUPDATE_DATE = sysdate, a.passbincount = " & _
          "(SELECT b.passbincount " & _
            " FROM tsv_mapping_tbl b " & _
           " WHERE b.customershortname = 'HD' AND b.substrateid = a.substrateid), " & _
      " a.failbincount = " & _
          "(SELECT d.failbincount " & _
           "  FROM tsv_mapping_tbl d " & _
           " WHERE d.customershortname = 'HD' AND d.substrateid = a.substrateid) " & _
 "WHERE a.customershortname = 'HD' " & _
  " AND a.failbincount = 0 " & _
  " AND a.substrateid IN (SELECT c.substrateid " & _
                          " FROM tsv_mapping_tbl c " & _
                         " WHERE c.customershortname = 'HD')"

Cnn.Execute cmdStr3

 Cnn.CommitTrans
 
 
SumCount = SumCount + 1
                         
                         
   
 Exit Sub
 
DealError:

' SumCount = SumCount - 1

Cnn.RollbackTrans

End Sub

Public Sub updateGCWltMap(waferIdTemp As String, goodDieQtyTemp As Long)
Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008

On Error GoTo DealError

'2014-11-13 jiayun add table TSV_Mapping_tbl

'导入Mapping
Cnn.BeginTrans
                  
'cmdStr = "insert into TSV_Mapping_tbl (id,SubstrateId,SubstrateType,LotId,ProductId,CreateDate,MicronLotId,PassBinCount,FailBinCount, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE,customershortname,FileName) Values" & _
'          " ( mappingData_SEQ.Nextval, '" & mapTemp.SubstrateId & "', '" & mapTemp.SubstrateType & "', '" & mapTemp.lotid & "', '" & mapTemp.ProductId & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PassBinCount & "," & mapTemp.FailBinCount & ",'Y','Auto',sysdate,'" & custNameTemp & "','" & mapTemp.FileName & "')"
'
                      
cmdStr = "update mappingdatatest  a " & _
" set a.QTECH_LASTUPDATE_BY = '" & gUserRealName & "',a.QTECH_LASTUPDATE_DATE = sysdate,a.passbincount=" & goodDieQtyTemp & ",a.failbincount=4961-" & goodDieQtyTemp & " " & _
" where customershortname='GC' and remark='WLT' and a.substrateid= '" & waferIdTemp & "' "

                      
'cmdStr2 = "INSERT INTO [ERPBASE].[dbo].[tblmappingDataCus]([SUBSTRATEID],[SUBSTRATETYPE],[LOTID],[PRODUCTID],[CREATEDATE],[MICRONLOTID],[PASSBINCOUNT],[FAILBINCOUNT],[FLAG],[QTECH_CREATED_BY],[QTECH_CREATED_DATE],[CUSTOMERSHORTNAME] ,FileName) Values" & _
'          " ( '" & mapTemp.SubstrateId & "', '" & mapTemp.SubstrateType & "', '" & mapTemp.lotid & "', '" & mapTemp.ProductId & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PassBinCount & "," & mapTemp.FailBinCount & ",'Y','Auto',GETDATE(),'" & custNameTemp & "','" & mapTemp.FileName & "')"
'
If AddSql(cmdStr) = 0 Then
    MsgBox "更新失败", vbCritical, "警告"
    Exit Sub
End If
 
'AddSql2 (cmdStr2)
 
 
 Cnn.CommitTrans
 
SumCount = SumCount + 1
   
 Exit Sub
 
DealError:

' SumCount = SumCount - 1

Cnn.RollbackTrans

End Sub




Public Sub AddCustomerShipAddress(ship As CUSSHIPADDRESS)
Dim strSql  As String

strSql = "insert into erptemp..customer_information(CUSTOMER, SHIP_TO, SHIPPER, SOLD_TO, BILL_TO, SHIP_TO_AD, SOLD_BY, PAYMENT_TERMS, CURRENCY, BANK_INFORMATION, TK, PO, CREAT_BY, CREAT_DATE, SHIPPER_PACK) " & _
"values('" & ship.CUSTOMER & "','" & ship.SHIP_TO & "', '" & ship.SHIPPER & "', '" & ship.SOLD_TO & "', '" & ship.BILL_TO & "','" & ship.SHIP_TO_AD & "','" & ship.SOLD_BY & "', '" & ship.PAYMENT_TERMS & "', '" & ship.CURRENCY & "', '" & ship.BANKINFO & "','" & ship.TK & "', '" & ship.PO & "','" & gUserName & "', GETDATE(), '" & ship.SHIPPER_PACK & "')"

If AddSql2(strSql) = 0 Then

    MsgBox "保存失败, 请导出确认", vbCritical, "警告"

End If

End Sub

Public Sub updateGCCOGMap(waferIdTemp As String, goodDieQtyTemp As Long, ngQtyTemp As Long)
Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008

On Error GoTo DealError

'更新Mapping
Cnn.BeginTrans
                  

                      
cmdStr = "update mappingdatatest  a " & _
" set a.passbincount=" & goodDieQtyTemp & ",a.failbincount=" & ngQtyTemp & " " & _
" where customershortname='GC'  and a.substrateid= '" & waferIdTemp & "' "

                      

cmdStr2 = "update [ERPBASE].[dbo].[tblmappingData]   " & _
" set passbincount=" & goodDieQtyTemp & ",failbincount=" & ngQtyTemp & " " & _
" where customershortname='GC'  and substrateid= '" & waferIdTemp & "' "


AddSql (cmdStr)
 
AddSql2 (cmdStr2)
 
 
 Cnn.CommitTrans
 
SumCount = SumCount + 1
   
 Exit Sub
 
DealError:

' SumCount = SumCount - 1

Cnn.RollbackTrans

End Sub

Public Sub updateMGMap(waferIdTemp As String, goodDieQtyTemp As Long, ngDieQtyTemp As Long)
Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008

On Error GoTo DealError

'2014-11-13 jiayun add table TSV_Mapping_tbl

'导入Mapping
Cnn.BeginTrans
                  
'cmdStr = "insert into TSV_Mapping_tbl (id,SubstrateId,SubstrateType,LotId,ProductId,CreateDate,MicronLotId,PassBinCount,FailBinCount, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE,customershortname,FileName) Values" & _
'          " ( mappingData_SEQ.Nextval, '" & mapTemp.SubstrateId & "', '" & mapTemp.SubstrateType & "', '" & mapTemp.lotid & "', '" & mapTemp.ProductId & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PassBinCount & "," & mapTemp.FailBinCount & ",'Y','Auto',sysdate,'" & custNameTemp & "','" & mapTemp.FileName & "')"
'
                      
cmdStr = "update mappingdatatest  a " & _
" set a.QTECH_LASTUPDATE_BY = '" & gUserRealName & "',a.QTECH_LASTUPDATE_DATE = sysdate,a.passbincount=" & goodDieQtyTemp & ",a.failbincount=" & ngDieQtyTemp & "  " & _
" where customershortname='MG'  and a.substrateid= '" & waferIdTemp & "' "

                      
'cmdStr2 = "INSERT INTO [ERPBASE].[dbo].[tblmappingDataCus]([SUBSTRATEID],[SUBSTRATETYPE],[LOTID],[PRODUCTID],[CREATEDATE],[MICRONLOTID],[PASSBINCOUNT],[FAILBINCOUNT],[FLAG],[QTECH_CREATED_BY],[QTECH_CREATED_DATE],[CUSTOMERSHORTNAME] ,FileName) Values" & _
'          " ( '" & mapTemp.SubstrateId & "', '" & mapTemp.SubstrateType & "', '" & mapTemp.lotid & "', '" & mapTemp.ProductId & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PassBinCount & "," & mapTemp.FailBinCount & ",'Y','Auto',GETDATE(),'" & custNameTemp & "','" & mapTemp.FileName & "')"
'


cmdStr2 = "update [ERPBASE].[dbo].[tblmappingDataMG]   " & _
" set passbincount=" & goodDieQtyTemp & ",failbincount=" & ngDieQtyTemp & "  " & _
" where customershortname='MG'  and substrateid= '" & waferIdTemp & "' "


If AddSql(cmdStr) = 0 Then
    MsgBox "更新失败", vbCritical, "警告"
    Exit Sub
End If
 
AddSql2 (cmdStr2)
 
 
 Cnn.CommitTrans
 
SumCount = SumCount + 1
   
 Exit Sub
 
DealError:

' SumCount = SumCount - 1

Cnn.RollbackTrans

End Sub

Public Sub updateHW50Map(waferIdTemp As String, goodDieQtyTemp As Long, ngDieQtyTemp As Long)

Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008

On Error GoTo DealError

'2014-11-13 jiayun add table TSV_Mapping_tbl

'导入Mapping
Cnn.BeginTrans
                  
'cmdStr = "insert into TSV_Mapping_tbl (id,SubstrateId,SubstrateType,LotId,ProductId,CreateDate,MicronLotId,PassBinCount,FailBinCount, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE,customershortname,FileName) Values" & _
'          " ( mappingData_SEQ.Nextval, '" & mapTemp.SubstrateId & "', '" & mapTemp.SubstrateType & "', '" & mapTemp.lotid & "', '" & mapTemp.ProductId & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PassBinCount & "," & mapTemp.FailBinCount & ",'Y','Auto',sysdate,'" & custNameTemp & "','" & mapTemp.FileName & "')"
'
                      
cmdStr = "update mappingdatatest  a " & _
" set a.QTECH_LASTUPDATE_BY = '" & gUserRealName & "',a.QTECH_LASTUPDATE_DATE = sysdate,a.passbincount=" & goodDieQtyTemp & ",a.failbincount=" & ngDieQtyTemp & "  " & _
" where customershortname='HW50'  and a.substrateid= '" & waferIdTemp & "' "

                      
'cmdStr2 = "INSERT INTO [ERPBASE].[dbo].[tblmappingDataCus]([SUBSTRATEID],[SUBSTRATETYPE],[LOTID],[PRODUCTID],[CREATEDATE],[MICRONLOTID],[PASSBINCOUNT],[FAILBINCOUNT],[FLAG],[QTECH_CREATED_BY],[QTECH_CREATED_DATE],[CUSTOMERSHORTNAME] ,FileName) Values" & _
'          " ( '" & mapTemp.SubstrateId & "', '" & mapTemp.SubstrateType & "', '" & mapTemp.lotid & "', '" & mapTemp.ProductId & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PassBinCount & "," & mapTemp.FailBinCount & ",'Y','Auto',GETDATE(),'" & custNameTemp & "','" & mapTemp.FileName & "')"
'


cmdStr2 = "update [ERPBASE].[dbo].[tblmappingDataMG]   " & _
" set passbincount=" & goodDieQtyTemp & ",failbincount=" & ngDieQtyTemp & "  " & _
" where customershortname='HW50'  and substrateid= '" & waferIdTemp & "' "


If AddSql(cmdStr) = 0 Then
    MsgBox "更新失败", vbCritical, "警告"
    Exit Sub
End If
 
AddSql2 (cmdStr2)
 
 
 Cnn.CommitTrans
 
SumCount = SumCount + 1
   
 Exit Sub
 
DealError:

' SumCount = SumCount - 1

Cnn.RollbackTrans


End Sub

Public Sub updateAT71Map(waferIdTemp As String, goodDieQtyTemp As Long, ngDieQtyTemp As Long)
Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008

On Error GoTo DealError

'2014-11-13 jiayun add table TSV_Mapping_tbl

'导入Mapping
Cnn.BeginTrans
                  
     
cmdStr = "update mappingdatatest  a " & _
" set a.QTECH_LASTUPDATE_BY = '" & gUserRealName & "',a.QTECH_LASTUPDATE_DATE = sysdate,a.passbincount=" & goodDieQtyTemp & ",a.failbincount=" & ngDieQtyTemp & "  " & _
" where  a.substrateid= '" & waferIdTemp & "' "

                      

cmdStr2 = "update [ERPBASE].[dbo].[tblmappingDataMG]   " & _
" set passbincount=" & goodDieQtyTemp & ",failbincount=" & ngDieQtyTemp & "  " & _
" where substrateid= '" & waferIdTemp & "' "


If AddSql(cmdStr) = 0 Then
    MsgBox "更新失败", vbCritical, "警告"
    Exit Sub

 End If
AddSql2 (cmdStr2)
 
 
 Cnn.CommitTrans
 
SumCount = SumCount + 1
   
 Exit Sub
 
DealError:

' SumCount = SumCount - 1

Cnn.RollbackTrans

End Sub



Public Sub updateAB18Map(waferIdTemp As String, goodDieQtyTemp As Long, ngDieQtyTemp As Long)
Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008

On Error GoTo DealError

'2014-11-13 jiayun add table TSV_Mapping_tbl

'导入Mapping
Cnn.BeginTrans
                                 
cmdStr = "update mappingdatatest  a " & _
" set a.QTECH_LASTUPDATE_BY = '" & gUserRealName & "',a.QTECH_LASTUPDATE_DATE = sysdate,a.passbincount=" & goodDieQtyTemp & ",a.failbincount=" & ngDieQtyTemp & "  " & _
" where customershortname='AB18'  and a.substrateid= '" & waferIdTemp & "' "

    
cmdStr2 = "update [ERPBASE].[dbo].[tblmappingDataMG]   " & _
" set passbincount=" & goodDieQtyTemp & ",failbincount=" & ngDieQtyTemp & "  " & _
" where customershortname='AB18'  and substrateid= '" & waferIdTemp & "' "


If AddSql(cmdStr) = 0 Then
    MsgBox "没有成功更新Mapping,请联系IT", vbCritical, "提醒"
    Exit Sub
End If
 
AddSql2 (cmdStr2)
 
 
 Cnn.CommitTrans
 
SumCount = SumCount + 1
   
 Exit Sub
 
DealError:

' SumCount = SumCount - 1

Cnn.RollbackTrans

End Sub


Public Sub AddMap2(mapTemp As MapRecord, customerNameTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008

On Error GoTo DealError

'导入Mapping
Cnn.BeginTrans
                  
cmdStr = "insert into mappingDataTest (id,SubstrateId,SubstrateType,LotId,ProductId,CreateDate,MicronLotId,PassBinCount,FailBinCount, FLAG ,QTECH_CREATED_BY ,QTECH_CREATED_DATE,customershortname,FileName) Values" & _
          " ( mappingData_SEQ.Nextval, '" & mapTemp.SUBSTRATEID & "', '" & mapTemp.substratetype & "', '" & mapTemp.LOTID & "', '" & mapTemp.PRODUCTID & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PASSBINCOUNT & "," & mapTemp.FailBinCount & ",'Y','Auto',sysdate,'" & customerNameTemp & "','" & mapTemp.filename & "')"
                      
                      
                      
cmdStr2 = "INSERT INTO [ERPBASE].[dbo].[tblmappingData]([SUBSTRATEID],[SUBSTRATETYPE],[LOTID],[PRODUCTID],[CREATEDATE],[MICRONLOTID],[PASSBINCOUNT],[FAILBINCOUNT],[FLAG],[QTECH_CREATED_BY],[QTECH_CREATED_DATE],[CUSTOMERSHORTNAME] ,FileName) Values" & _
          " ( '" & mapTemp.SUBSTRATEID & "', '" & mapTemp.substratetype & "', '" & mapTemp.LOTID & "', '" & mapTemp.PRODUCTID & "','" & mapTemp.CreateDate & "', '" & mapTemp.MicronLotId & "', " & mapTemp.PASSBINCOUNT & "," & mapTemp.FailBinCount & ",'Y','Auto',GETDATE(),'" & customerNameTemp & "','" & mapTemp.filename & "')"
                      
AddSql (cmdStr)
 
AddSql2 (cmdStr2)
 
 
 Cnn.CommitTrans
 Exit Sub
 
DealError:

Cnn.RollbackTrans

End Sub


Public Sub AddOI(TEMP As String, temp2 As String)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = " insert into CustomerOItbl_test values(" & TEMP & ")"
cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerOI] values(" & temp2 & ")"



                            
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
 SumCount = SumCount - 1

End Sub


Public Sub AddShippingUP(TEMP As String, temp2 As String)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = " insert into CUSTOMERSHIPPINGUPTBL values(" & TEMP & ")"
cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerShippingUp] values(" & temp2 & ")"
                        
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
Cnn.CommitTrans

SumCount = SumCount + 1

Exit Sub
DealError:

Cnn.RollbackTrans
 SumCount = SumCount - 1

End Sub

Public Sub AddShippingUPDATE(dnTemp As String, jobtemp As String, jobqtytemp As Long)

    Dim cmdStr  As String

    Dim cmdStr2 As String

    On Error GoTo DealError
        
    Cnn.BeginTrans

    cmdStr = " update  CUSTOMERSHIPPINGUPTBL set quantity =(quantity + '" & jobqtytemp & "') where delivery = '" & dnTemp & "' and batchnumber = '" & jobtemp & "'"
    cmdStr2 = " update [ERPBASE].[dbo].[tblCustomerShippingUp] set quantity =(quantity + '" & jobqtytemp & "')  where delivery = '" & dnTemp & "' and batchnumber = '" & jobtemp & "'"
             
    AddSql (cmdStr)
    AddSql2 (cmdStr2)
 
    Cnn.CommitTrans

    SumCount = SumCount + 1

    Exit Sub
DealError:

    Cnn.RollbackTrans
    SumCount = SumCount - 1

End Sub




Public Sub AddBI(TEMP As String)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = " insert into CustomerBItbl values(" & TEMP & ")"

                            
AddSql (cmdStr)


Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
 SumCount = SumCount - 1

End Sub

Public Sub AddWOPRI(woTemp As String, pritemp As String, LotType As String, player As String)
Dim cmdStr As String

cmdStr = " insert into PJ_WO_PRI values('" & woTemp & "','" & pritemp & "',to_char(sysdate,'YYYY-MM-DD'),'" & LotType & "', '" & player & "')"
AddSql (cmdStr)

End Sub

Public Sub UpdateEDCData(waferIdTemp As String, typenameTemp As String, valueTemp As Integer)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
Cnn.BeginTrans

 cmdStr = " Update A_WipLotDetailsDataHistory Set WipDataValue = " & valueTemp & " " & _
" where WaferScribeNumber = '" & waferIdTemp & "' " & _
" and WipDataNameName like '槽%' and wipdatavalue is not null  " & _
" and WipDataNameName='" & typenameTemp & "' "


                            
AddSql (cmdStr)

 cmdStr = "insert into EDC_Update_His (WaferId,TypeName,valueData) values ('" & waferIdTemp & "','" & typenameTemp & "'," & valueTemp & ") "
AddSql (cmdStr)


Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
 SumCount = SumCount - 1

End Sub



Public Sub AddBC(TEMP As String, temp2 As String)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
'Cnn.BeginTrans

cmdStr = " insert into CustomerBCtbl values(" & TEMP & ")"
cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerBC] values(" & temp2 & ")"



                            
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
'Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True

End Sub

Public Sub AddONBC(TEMP As String, temp2 As String)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
'Cnn.BeginTrans

cmdStr = " insert into CustomerBCtbl (ID ,BATCHID ,APTINADOCNUMBER  ,LOTRECDATE  ," & _
" MTRLNUM,DESIGNID  ,CURRENT_WAFER_QTY,FLAG ,CREATEBY ,CreateDate,mtrldesc) values(" & TEMP & ")"
cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerBC] (ID ,BATCHID ,APTINADOCNUMBER  ,LOTRECDATE  ," & _
" MTRLNUM,DESIGNID, CURRENT_WAFER_QTY ,FLAG ,CREATEBY ,CreateDate,mtrldesc) values(" & temp2 & ")"



                            
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
'Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True


End Sub



Public Sub Add37POHeader(TEMP As SemtechPOHeader)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
'Cnn.BeginTrans

cmdStr = " insert into CUSTOMEROITBL_TEST (ID ,PO_NUM ,PO_ITEM ,SOURCE_BATCH_ID ,SOURCE_MTRL_NUM, " & _
" MPN ,MPN_DESC ,SOURCE_MTRL_SLOC,OFFSHORE_ASM_COMPANY,OFFSHORE_TEST_COMPANY," & _
" CURRENT_WAFER_QTY,DIE_QTY ,COUNTRY_OF_FAB,RETICLE_LEVEL_71 ,IMAGER_CUSTOMER_REV ," & _
" PACKAGE_TYPE , BOX_TYPE,SHIPPING_MST_260 ,SHIPPING_MST_LEVEL ,SHIP_COMMENT," & _
" CREATED_DATE  ,REF_PO  ,COUNTRY_OF_ASSEMBLY  ,DATE_CODE  ,  SHIP_SITE ," & _
" CUSTOM_PART_NO , FLAG, QTECH_CREATED_BY, QTECH_CREATED_DATE, CustomershortName ,JOBNO) values(" & _
" '" & TEMP.id & "', '" & TEMP.PurchaseOrderNo & "','" & TEMP.ITEM & "', '" & TEMP.WAFERLOT & "', '" & TEMP.PartNumber & "', " & _
" '" & TEMP.ProductionOrderNo & "', '" & TEMP.YourMaterialNumber & "','" & TEMP.LotNO & "', '" & TEMP.MfgPlant & "', '" & TEMP.ReceivingPlant & "', " & _
" '" & TEMP.Quantity & "', '" & TEMP.NetAmount & "','" & TEMP.WaferFAB & "', '" & TEMP.Version & "', '" & TEMP.WaferREV & "', " & _
" '" & TEMP.TypeService & "', '" & TEMP.UM & "','" & TEMP.CURRENCY & "', '" & TEMP.FreightCarrier & "', '" & TEMP.TermsDelivery & "', " & _
" '" & TEMP.DATE & "', '" & TEMP.UnitPrice & "','" & TEMP.TermsPayment & "', '" & TEMP.DelDate & "', '" & TEMP.ShippingAddress & "', " & _
" '" & TEMP.KeyStr & "', 'Y','" & TEMP.QTECH_CREATED_BY & "', sysdate, '37', '" & TEMP.JOBNO & "') "

cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerOI] (ID ,PO_NUM ,PO_ITEM ,SOURCE_BATCH_ID ,SOURCE_MTRL_NUM, " & _
" MPN ,MPN_DESC ,SOURCE_MTRL_SLOC,OFFSHORE_ASM_COMPANY,OFFSHORE_TEST_COMPANY," & _
" CURRENT_WAFER_QTY,DIE_QTY ,COUNTRY_OF_FAB,RETICLE_LEVEL_71 ,IMAGER_CUSTOMER_REV ," & _
" PACKAGE_TYPE , BOX_TYPE,SHIPPING_MST_260 ,SHIPPING_MST_LEVEL ,SHIP_COMMENT," & _
" CREATED_DATE  ,REF_PO  ,COUNTRY_OF_ASSEMBLY  ,DATE_CODE  ,  SHIP_SITE  ," & _
" CUSTOM_PART_NO , FLAG, QTECH_CREATED_BY, QTECH_CREATED_DATE, CustomershortName,JOBNO ) values(" & _
" '" & TEMP.id & "', '" & TEMP.PurchaseOrderNo & "','" & TEMP.ITEM & "', '" & TEMP.WAFERLOT & "', '" & TEMP.PartNumber & "', " & _
" '" & TEMP.ProductionOrderNo & "', '" & TEMP.YourMaterialNumber & "','" & TEMP.LotNO & "', '" & TEMP.MfgPlant & "', '" & TEMP.ReceivingPlant & "', " & _
" '" & TEMP.Quantity & "', '" & TEMP.NetAmount & "','" & TEMP.WaferFAB & "', '" & TEMP.Version & "', '" & TEMP.WaferREV & "', " & _
" '" & TEMP.TypeService & "', '" & TEMP.UM & "','" & TEMP.CURRENCY & "', '" & TEMP.FreightCarrier & "', '" & TEMP.TermsDelivery & "', " & _
" '" & TEMP.DATE & "', '" & TEMP.UnitPrice & "','" & TEMP.TermsPayment & "', '" & TEMP.DelDate & "', '" & TEMP.ShippingAddress & "', " & _
" '" & TEMP.KeyStr & "', 'Y','" & TEMP.QTECH_CREATED_BY & "', getdate(), '37','" & TEMP.JOBNO & "' ) "


                            
AddSql (cmdStr)
AddSql2 (cmdStr2)

 SumCount = SumCount + 1
 
'Cnn.CommitTrans

Exit Sub
DealError:

'Cnn.RollbackTrans
SumCount = SumCount - 1

End Sub


Public Sub update37po(TEMP As SemtechPOHeader)
Dim cmdStr As String
Dim cmdStr2 As String

End Sub

Public Sub Add37POHeaderICI(TEMP As SemtechPOHeader, WAFER As String)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
'Cnn.BeginTrans

cmdStr = " insert into CUSTOMEROITBL_TEST (ID ,PO_NUM ,PO_ITEM ,SOURCE_BATCH_ID ,SOURCE_MTRL_NUM, mtrl_num," & _
" MPN ,MPN_DESC ,SOURCE_MTRL_SLOC,OFFSHORE_ASM_COMPANY,OFFSHORE_TEST_COMPANY," & _
" CURRENT_WAFER_QTY,DIE_QTY ,COUNTRY_OF_FAB,RETICLE_LEVEL_72 ,IMAGER_CUSTOMER_REV ," & _
" PACKAGE_TYPE , BOX_TYPE,SHIPPING_MST_260 ,SHIPPING_MST_LEVEL ,SHIP_COMMENT," & _
" CREATED_DATE  ,REF_PO  ,COUNTRY_OF_ASSEMBLY  ,DATE_CODE  ,  SHIP_SITE   ," & _
" CUSTOM_PART_NO , FLAG, QTECH_CREATED_BY, QTECH_CREATED_DATE, CustomershortName ,BATCH_COMMENT_TEST,t_price,MTRL_DESC,test_mtrl_desc, fab_conv_id,reticle_level_71) values(" & _
" '" & TEMP.id & "', '" & TEMP.PurchaseOrderNo & "','" & TEMP.ITEM & "', '" & TEMP.po1lot & "', '" & TEMP.PartNumber & "', '" & TEMP.BagNo & "'," & _
" '" & TEMP.ProductionOrderNo & "', '" & TEMP.YourMaterialNumber & "','" & TEMP.LotNumber & "', '" & TEMP.MfgPlant & "', '" & TEMP.ReceivingPlant & "', " & _
" '" & TEMP.NetAmount & "', '" & TEMP.Quantity & "', '" & TEMP.WaferFAB & "', '" & TEMP.Version & "', '" & TEMP.WaferREV & "', " & _
" '" & TEMP.TypeService & "', '" & TEMP.UM & "','" & TEMP.CURRENCY & "', '" & TEMP.FreightCarrier & "', '" & TEMP.TermsDelivery & "', " & _
" '" & TEMP.DATE & "', '" & TEMP.UnitPrice & "','" & TEMP.TermsPayment & "', '" & TEMP.DATECODE & "', '" & TEMP.ShippingAddress & "', " & _
" '" & TEMP.KeyStr & "', 'Y','" & TEMP.QTECH_CREATED_BY & "', sysdate, '37','" & TEMP.DelDate & "'," & TEMP.POPrice & ",'" & TEMP.MaterialDes & "','" & TEMP.LotNO & "', '" & WAFER & "' ,'" & TEMP.ItemLineText & "') "

cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerOI] (ID ,PO_NUM ,PO_ITEM ,SOURCE_BATCH_ID ,SOURCE_MTRL_NUM, mtrl_num," & _
" MPN ,MPN_DESC ,SOURCE_MTRL_SLOC,OFFSHORE_ASM_COMPANY,OFFSHORE_TEST_COMPANY," & _
" CURRENT_WAFER_QTY,DIE_QTY ,COUNTRY_OF_FAB,RETICLE_LEVEL_72 ,IMAGER_CUSTOMER_REV ," & _
" PACKAGE_TYPE , BOX_TYPE,SHIPPING_MST_260 ,SHIPPING_MST_LEVEL ,SHIP_COMMENT," & _
" CREATED_DATE  ,REF_PO  ,COUNTRY_OF_ASSEMBLY  ,DATE_CODE  ,  SHIP_SITE  ," & _
" CUSTOM_PART_NO , FLAG, QTECH_CREATED_BY, QTECH_CREATED_DATE, CustomershortName,BATCH_COMMENT_TEST ,t_price,MTRL_DESC,test_mtrl_desc, fab_conv_id,reticle_level_71) values(" & _
" '" & TEMP.id & "', '" & TEMP.PurchaseOrderNo & "','" & TEMP.ITEM & "', '" & TEMP.po1lot & "', '" & TEMP.PartNumber & "', '" & TEMP.BagNo & "'," & _
" '" & TEMP.ProductionOrderNo & "', '" & TEMP.YourMaterialNumber & "','" & TEMP.LotNumber & "', '" & TEMP.MfgPlant & "', '" & TEMP.ReceivingPlant & "', " & _
" '" & TEMP.NetAmount & "','" & TEMP.Quantity & "','" & TEMP.WaferFAB & "', '" & TEMP.Version & "', '" & TEMP.WaferREV & "', " & _
" '" & TEMP.TypeService & "', '" & TEMP.UM & "','" & TEMP.CURRENCY & "', '" & TEMP.FreightCarrier & "', '" & TEMP.TermsDelivery & "', " & _
" '" & TEMP.DATE & "', '" & TEMP.UnitPrice & "','" & TEMP.TermsPayment & "', '" & TEMP.DATECODE & "', '" & TEMP.ShippingAddress & "', " & _
" '" & TEMP.KeyStr & "', 'Y','" & TEMP.QTECH_CREATED_BY & "', getdate(), '37' ,'" & TEMP.DelDate & "'," & TEMP.POPrice & ",'" & TEMP.MaterialDes & "','" & TEMP.LotNO & "', '" & WAFER & "','" & TEMP.ItemLineText & "') "


                            
AddSql (cmdStr)
AddSql2 (cmdStr2)

 SumCount = SumCount + 1
 
'Cnn.CommitTrans

Exit Sub
DealError:

'Cnn.RollbackTrans
SumCount = SumCount - 1

End Sub




Public Sub Add68POHeader(TEMP As SemtechPOHeader, customerTemp As String, waferidSTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
'Cnn.BeginTrans

cmdStr = " insert into CUSTOMEROITBL_TEST (ID ,PO_NUM ,SOURCE_BATCH_ID ,MPN_DESC,SOURCE_MTRL_SLOC, " & _
" CURRENT_WAFER_QTY ,CREATED_DATE,SHIP_SITE,FLAG, QTECH_CREATED_BY, QTECH_CREATED_DATE, CustomershortName ,REF_PO,RETICLE_LEVEL_71) values(" & _
" '" & TEMP.id & "', '" & TEMP.PurchaseOrderNo & "', '" & TEMP.WAFERLOT & "', '" & TEMP.YourMaterialNumber & "','" & TEMP.LotNO & "', " & _
" '" & TEMP.Quantity & "','" & TEMP.DATE & "','" & TEMP.ShippingAddress & "', " & _
"  'Y','" & TEMP.QTECH_CREATED_BY & "', sysdate, '" & customerTemp & "' , '" & TEMP.UnitPrice & "','" & waferidSTemp & "') "


cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerOI] (ID ,PO_NUM ,SOURCE_BATCH_ID ,MPN_DESC,SOURCE_MTRL_SLOC, " & _
" CURRENT_WAFER_QTY ,CREATED_DATE,SHIP_SITE,FLAG, QTECH_CREATED_BY, QTECH_CREATED_DATE, CustomershortName ,REF_PO,RETICLE_LEVEL_71) values(" & _
" '" & TEMP.id & "', '" & TEMP.PurchaseOrderNo & "', '" & TEMP.WAFERLOT & "', '" & TEMP.YourMaterialNumber & "','" & TEMP.LotNO & "', " & _
" '" & TEMP.Quantity & "','" & TEMP.DATE & "','" & TEMP.ShippingAddress & "', " & _
"  'Y','" & TEMP.QTECH_CREATED_BY & "', getdate(), '" & customerTemp & "' , '" & TEMP.UnitPrice & "','" & waferidSTemp & "') "


                            
AddSql (cmdStr)
AddSql2 (cmdStr2)

 SumCount = SumCount + 1
 
'Cnn.CommitTrans

Exit Sub
DealError:

'Cnn.RollbackTrans
SumCount = SumCount - 1

End Sub

Public Sub AddONForcast(TEMP As String, temp2 As String)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
'Cnn.BeginTrans

cmdStr = " insert into CUSTOMERFORECASTTBL( ID , TYPENAME , DEMAND_TYPE  ,OUT_PART_ID,START_PART_ID ," & _
"   OUT_DATE ,OUT_QTY,WORKWEEK ,SITE_ID,STAGE_ID , CTG , PTI2 ,COMMENTS,FLAG ,QTECH_CREATED_BY , QTECH_CREATED_DATE ) values(" & TEMP & ")"


cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerForeCast]( ID , TYPENAME , DEMAND_TYPE  ,OUT_PART_ID,START_PART_ID ," & _
"   OUT_DATE ,OUT_QTY,WORKWEEK ,SITE_ID,STAGE_ID , CTG , PTI2 ,COMMENTS,FLAG ,QTECH_CREATED_BY , QTECH_CREATED_DATE )  values(" & temp2 & ")"



                            
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
'Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True


End Sub

Public Sub AddONForcastBePP(TEMP As String, temp2 As String)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
'Cnn.BeginTrans

cmdStr = " insert into CUSTOMERFORECASTTBL( ID , TYPENAME , CREATE_DATE ,SITE ,STAGE ,PART_ID,PROD_PART_ID ," & _
"     CONS_ITEM ,STARTING_WEEK ,NEXT_SITE , MFG_AREA_CD,ORACLE_LOC_CD ," & _
"     PTI , PKG_CD ,PKG_GRP_CD ,SCH_COMMENTS ,QTY ,I_T,ON_HAND , FLAG ,QTECH_CREATED_BY , QTECH_CREATED_DATE ) values(" & TEMP & ")"


cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerForeCast]( ID , TYPENAME , CREATE_DATE ,SITE ,STAGE ,PART_ID,PROD_PART_ID ," & _
"     CONS_ITEM ,STARTING_WEEK ,NEXT_SITE , MFG_AREA_CD,ORACLE_LOC_CD ," & _
"     PTI , PKG_CD ,PKG_GRP_CD ,SCH_COMMENTS ,QTY ,I_T,ON_HAND , FLAG ,QTECH_CREATED_BY , QTECH_CREATED_DATE )  values(" & temp2 & ")"



                            
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
'Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True


End Sub






' add VT
Public Sub AddVT(TEMP As VTData)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = "insert into  TSV_VT_History (" & _
" SHIPDATE,DELIVERYNO,CUSTDEVICE,CUSTLOT,GOODDIEQTY," & _
" NGDIEQTY,TTL,NETWEIGHT,GROSSWEIGHT,REMARK," & _
" Flag , Created_By, created_date,id) values  " & _
" ('" & TEMP.SHIPDATETemp & "','" & TEMP.DeliveryNoTemp & "','" & TEMP.CustDeviceTemp & "','" & TEMP.CUSTLOTTemp & "'," & _
" " & TEMP.goodDieQtyTemp & "," & TEMP.ngDieQtyTemp & "," & _
" '" & TEMP.TTLTemp & "'," & _
" '" & TEMP.NetWeightTemp & "','" & TEMP.GrossWeightTemp & "','" & TEMP.remarkTemp & "'," & _
" 'Y','" & TEMP.Created_ByTemp & "',sysdate, tbl_tsv_VTData_seq.Nextval)"

                
AddSql (cmdStr)

Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True


End Sub

' add VT
Public Sub AddVTCustomer(TEMP As VTData, customerTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = "insert into  TSV_VT_History (" & _
" SHIPDATE,DELIVERYNO,CUSTDEVICE,CUSTLOT,GOODDIEQTY," & _
" NGDIEQTY,TTL,NETWEIGHT,GROSSWEIGHT,REMARK," & _
" Flag , Created_By, created_date,id,customershortname) values  " & _
" ('" & TEMP.SHIPDATETemp & "','" & TEMP.DeliveryNoTemp & "','" & TEMP.CustDeviceTemp & "','" & TEMP.CUSTLOTTemp & "'," & _
" " & TEMP.goodDieQtyTemp & "," & TEMP.ngDieQtyTemp & "," & _
" '" & TEMP.TTLTemp & "'," & _
" '" & TEMP.NetWeightTemp & "','" & TEMP.GrossWeightTemp & "','" & TEMP.remarkTemp & "'," & _
" 'Y','" & TEMP.Created_ByTemp & "',sysdate, tbl_tsv_VTData_seq.Nextval,'" & customerTemp & "')"

                
AddSql (cmdStr)

Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True


End Sub

Public Sub AddShipSideData(TEMP As ShipSideData)
Dim cmdStr As String
Dim cmdStr2 As String

'添加导入Sqlserver
On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = "insert into  erpdata.. tblSale_Shipto (" & _
" CustCode,DeviceName,LotID,WaferQty,ShipTo) values  " & _
" ('" & TEMP.CustomerCode & "','" & TEMP.GULFDeviceName & "','" & TEMP.GULFLotID & "','" & TEMP.WaferQTY & "', '" & TEMP.ShipTo & "')"

AddSql2 (cmdStr)

Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True

End Sub

Public Sub AddB2b(bT As B2B)
Dim cmdStr As String
Dim cmdStr2 As String

On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = "insert into B2B_ORDERTBL values('" & bT.EVENT_DATE & "','" & bT.EVENT & "','" & bT.OVT_COMPANY & "','" & bT.OVT_ORG & "','" & bT.SUB_NAME & "','" & bT.Stage & "','" & bT.WAFER_LOT & "','" & bT.OVT_JOB & "','" & bT.SUB_LOT & "','" & bT.Fab_Device & "', " & _
"'" & bT.SOURCE_DEVICE & "','" & bT.TARGET_DEVICE & "','" & bT.WAFER_QTY & "','" & bT.WAFER_DIE & "', '" & bT.wafer_id & "','" & bT.PO & "','" & bT.PO_RELEASE & "','" & bT.OPERATION_CODE & "','" & bT.OPERATION_DESCRIPTION & "','" & bT.RECEIVE_QTY & "','" & bT.JOB_ISSUE_QTY & "','" & bT.START_QTY & "','" & bT.COMPLETED_QTY & "', " & _
"'" & bT.GRADE_RECORD & "','" & bT.HOLD_CODE & "','" & bT.HOLD_QTY & "','" & bT.SCRAP_CODE & "','" & bT.SCRAP_QTY & "','" & bT.SCRAP_WAFER_ID & "','" & bT.Priority & "','" & bT.DATECODE & "','" & bT.ENG_NO & "','" & bT.NEXT_SUB_LOCATION & "','" & bT.LOT_TYPE & "','" & bT.E_SOD & "','" & bT.TEST_PROGRAM & "','" & bT.RMA_NO & "','" & bT.BILLING_INVOICE & "','" & bT.SHIPPING_INVOICE & "'," & _
"'" & bT.SHIPPING_DESTINATION & "','" & bT.JOB_FLAG & "','" & bT.REMARK & "','" & bT.SO & "','" & bT.SO_LINE & "','" & bT.GROSS_WEIGHT & "','" & bT.NET_WEIGHT & "','" & bT.FORWARDER & "','" & bT.AIR_WAYBILL & "','" & bT.CARTON_QTY & "', sysdate, '" & gUserName & "')"

AddSql (cmdStr)

cmdStr2 = "insert into [ERPBASE].[dbo].[B2B_ORDERTBL] values('" & bT.EVENT_DATE & "','" & bT.EVENT & "','" & bT.OVT_COMPANY & "','" & bT.OVT_ORG & "','" & bT.SUB_NAME & "','" & bT.Stage & "','" & bT.WAFER_LOT & "','" & bT.OVT_JOB & "','" & bT.SUB_LOT & "','" & bT.Fab_Device & "', " & _
"'" & bT.SOURCE_DEVICE & "','" & bT.TARGET_DEVICE & "','" & bT.WAFER_QTY & "','" & bT.WAFER_DIE & "', '" & bT.wafer_id & "','" & bT.PO & "','" & bT.PO_RELEASE & "','" & bT.OPERATION_CODE & "','" & bT.OPERATION_DESCRIPTION & "','" & bT.RECEIVE_QTY & "','" & bT.JOB_ISSUE_QTY & "','" & bT.START_QTY & "','" & bT.COMPLETED_QTY & "', " & _
"'" & bT.GRADE_RECORD & "','" & bT.HOLD_CODE & "','" & bT.HOLD_QTY & "','" & bT.SCRAP_CODE & "','" & bT.SCRAP_QTY & "','" & bT.SCRAP_WAFER_ID & "','" & bT.Priority & "','" & bT.DATECODE & "','" & bT.ENG_NO & "','" & bT.NEXT_SUB_LOCATION & "','" & bT.LOT_TYPE & "','" & bT.E_SOD & "','" & bT.TEST_PROGRAM & "','" & bT.RMA_NO & "','" & bT.BILLING_INVOICE & "','" & bT.SHIPPING_INVOICE & "'," & _
"'" & bT.SHIPPING_DESTINATION & "','" & bT.JOB_FLAG & "','" & bT.REMARK & "','" & bT.SO & "','" & bT.SO_LINE & "','" & bT.GROSS_WEIGHT & "','" & bT.NET_WEIGHT & "','" & bT.FORWARDER & "','" & bT.AIR_WAYBILL & "','" & bT.CARTON_QTY & "', GetDate(), '" & gUserName & "')"

AddSql2 (cmdStr2)

Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True

End Sub

Public Sub AddB2b2(bT As US026INVOICE)
Dim cmdStr As String
Dim cmdStr2 As String

On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = "insert into US026_INVOICE values('" & bT.SHIPPING_INVOICE & "','" & bT.SHIPPING_DESTINATION & "','" & bT.GROSS_WEIGHT & "','" & bT.NET_WEIGHT & "','" & bT.FORWAROER & "','" & bT.AIR_WAYBILL & "','" & bT.CARTON_QTY & "', sysdate, '" & gUserName & "', '" & bT.BILLING_INVOICE & "', '" & bT.FLAG & "')"

AddSql (cmdStr)

cmdStr2 = "insert into [ERPBASE].[dbo].[US026_INVOICE] values('" & bT.SHIPPING_INVOICE & "','" & bT.SHIPPING_DESTINATION & "','" & bT.GROSS_WEIGHT & "','" & bT.NET_WEIGHT & "','" & bT.FORWAROER & "','" & bT.AIR_WAYBILL & "','" & bT.CARTON_QTY & "', GETDATE(), '" & gUserName & "', '" & bT.BILLING_INVOICE & "', '" & bT.FLAG & "')"

AddSql2 (cmdStr2)

Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True

End Sub

Public Sub AddShipSideData2(TEMP As ShipSideData)
Dim cmdStr As String
Dim cmdStr2 As String

'添加导入Sqlserver
On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = "insert into  ST_TR_SEQ (" & _
" CustCode,DeviceName,LotID,WaferQty,ShipTo) values  " & _
" ('" & TEMP.CustomerCode & "','" & TEMP.GULFDeviceName & "','" & TEMP.GULFLotID & "','" & TEMP.WaferQTY & "', '" & TEMP.ShipTo & "')"

AddSql2 (cmdStr)

Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True

End Sub



' add EQ shipping 2015-04-27
Public Sub AddEQShipping(TEMP As ShippingData)
Dim cmdStr As String
Dim cmdStr2 As String
Dim cmdStr3 As String


'添加导入Sqlserver
'On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = "insert into customershippingtbl ( " & _
" id,No,SubConPO ,item ,Quantity , " & _
" Device ,SPA,CSD,Lot ,DateCode1 ," & _
" DeliveryName ,DeliveryAddress ,Warehouse ,Location  ,ModeOfDelivery  , " & _
" DateCode ,SO,CarrierNotes ,Line ,ScheduleLine , " & _
" CustPN , CountryDistributor, Customer, CustomerPO, flag, CreatedBy, CreatedDate ) values  " & _
" (" & TEMP.idTemp & "," & TEMP.notemp & ",'" & TEMP.SubConPOTemp & "','" & TEMP.itemTemp & "',  " & TEMP.QuantityTemp & ", " & _
" '" & TEMP.devicetemp & "','" & TEMP.SPATemp & "','" & TEMP.CSDTemp & "','" & TEMP.lottemp & "','" & TEMP.DateCode1Temp & "'," & _
" '" & TEMP.DeliveryNameTemp & "','" & TEMP.DeliveryAddressTemp & "','" & TEMP.WarehouseTemp & "','" & TEMP.LocationTemp & "','" & TEMP.ModeOfDeliveryTemp & "'," & _
" '" & TEMP.dateCodeTemp & "','" & TEMP.soTemp & "','" & TEMP.CarrierNotesTemp & "','" & TEMP.lineTemp & "','" & TEMP.ScheduleLineTemp & "'," & _
" '" & TEMP.CustPNTemp & "','" & TEMP.CountryDistributorTemp & "','" & TEMP.customerTemp & "','" & TEMP.customerPoTemp & "','Y','" & TEMP.CreatedByTemp & "',sysdate )"




cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerShipping] ( " & _
" id,No,SubConPO ,item ,Quantity , " & _
" Device ,SPA,CSD,Lot ,DateCode1 ," & _
" DeliveryName ,DeliveryAddress ,Warehouse ,Location  ,ModeOfDelivery  , " & _
" DateCode ,SO,CarrierNotes ,Line ,ScheduleLine , " & _
" CustPN , CountryDistributor, Customer, CustomerPO, flag, CreatedBy, CreatedDate ) values  " & _
" (" & TEMP.idTemp & "," & TEMP.notemp & ",'" & TEMP.SubConPOTemp & "','" & TEMP.itemTemp & "',  " & TEMP.QuantityTemp & ", " & _
" '" & TEMP.devicetemp & "','" & TEMP.SPATemp & "','" & TEMP.CSDTemp & "','" & TEMP.lottemp & "','" & TEMP.DateCode1Temp & "'," & _
" '" & TEMP.DeliveryNameTemp & "','" & TEMP.DeliveryAddressTemp & "','" & TEMP.WarehouseTemp & "','" & TEMP.LocationTemp & "','" & TEMP.ModeOfDeliveryTemp & "'," & _
" '" & TEMP.dateCodeTemp & "','" & TEMP.soTemp & "','" & TEMP.CarrierNotesTemp & "','" & TEMP.lineTemp & "','" & TEMP.ScheduleLineTemp & "'," & _
" '" & TEMP.CustPNTemp & "','" & TEMP.CountryDistributorTemp & "','" & TEMP.customerTemp & "','" & TEMP.customerPoTemp & "','Y','" & TEMP.CreatedByTemp & "',GETDATE() )"
        
cmdStr3 = "insert into EQ_SHIPPING_REQUEST (SUBCONPO,ITEM,QUANTITY,DEVICE,SPA,CSD,LOT,DATECODE1,DELIVERYNAME,DELIVERYADDRESS," & _
         "WAREHOUSE,LOCATION,MODEOFDELIVERY,DATECODE,SO,CARRIERNOTES,LINE,SCHEDULELINE,CUSTPN,COUNTRYANDNAMEOFDISTRIBUTOR," & _
         "CUSTOMER,CUSTOMERPO) VALUES ('" & TEMP.SubConPOTemp & "','" & TEMP.itemTemp & "'," & TEMP.QuantityTemp & ",'" & TEMP.devicetemp & "' " & _
         " ,'" & TEMP.SPATemp & "','" & TEMP.CSDTemp & "','" & TEMP.lottemp & "','" & TEMP.DateCode1Temp & "','" & TEMP.DeliveryNameTemp & "' " & _
         " ,'" & TEMP.DeliveryAddressTemp & "','" & TEMP.WarehouseTemp & "','" & TEMP.LocationTemp & "','" & TEMP.ModeOfDeliveryTemp & "' " & _
         " ,'" & TEMP.dateCodeTemp & "','" & TEMP.soTemp & "','" & TEMP.CarrierNotesTemp & "','" & TEMP.lineTemp & "','" & TEMP.ScheduleLineTemp & "' " & _
         " ,'" & TEMP.CustPNTemp & "','" & TEMP.CountryDistributorTemp & "','" & TEMP.customerTemp & "','" & TEMP.customerPoTemp & "')"



AddSql (cmdStr)
AddSql2 (cmdStr2)
AddSql (cmdStr3)

Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True


End Sub




Public Sub AddUploadDetail(TEMP As String)
Dim cmdStr As String

On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = " insert into TSV_Wo_Mt_UploadTemp values(" & TEMP & ")"
                            
AddSql (cmdStr)

 
Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans


End Sub

Public Sub delUploadDetail()
Dim cmdStr As String

On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = " delete from TSV_Wo_Mt_UploadTemp "
                            
AddSql (cmdStr)

 
Cnn.CommitTrans

Exit Sub
DealError:

Cnn.RollbackTrans


End Sub





Public Sub DelBC(batchIdTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = " update CustomerBCtbl  set flag='N',lastupdateby='Auto',lastupdatedate=sysdate  where BatchId='" & batchIdTemp & "' and Flag='Y' "
cmdStr2 = "update [ERPBASE].[dbo].[tblCustomerBC]  SET flag='N' ,lastupdateby='Auto',lastupdatedate=GETDATE()  where BatchId='" & batchIdTemp & "' and Flag='Y' "

                            
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
Cnn.CommitTrans
MsgBox "删除成功！"


Exit Sub
DealError:

Cnn.RollbackTrans
MsgBox "删除失败！"
 

End Sub

Public Sub DeltestNo(productTemp As String, testTemp As String, productNewTemp As String)
Dim cmdStr As String
Dim cmdStr1 As String
Dim cmdStr2 As String
Dim cmdStr3 As String


cmdStr = " update tblTestNo  set testedition='" & testTemp & "',lastupdateby='Auto',lastupdatedate=sysdate  where productname='" & productTemp & "' and Flag='Y' and productnamenew= '" & productNewTemp & "' "
                         
cmdStr1 = " update TBLSETPT set testNo='" & testTemp & "',lastupdateby='Auto',lastupdatedate=sysdate where productName='" & productNewTemp & "' and flag='Y' "
                         
cmdStr2 = " update tblTestNo2  set testedition='" & testTemp & "',lastupdateby='Auto',lastupdatedate=sysdate  where productname='" & productTemp & "' and Flag='Y' and productnamenew= '" & productNewTemp & "' "

cmdStr3 = " update TSVCard_EDT  set testedition='" & testTemp & "',lastupdateby='Auto',lastupdatedate=sysdate  where productname='" & productTemp & "' and Flag='Y' and productnamenew= '" & productNewTemp & "' "


                                             
AddSql (cmdStr)

AddSql (cmdStr1)

AddSql (cmdStr2)

AddSql (cmdStr3)



MsgBox "修改成功！"

End Sub


Public Sub DelGCDieQty(productTemp As String, testTemp As String, productNewTemp As Long)
Dim cmdStr As String
Dim cmdStr1 As String
Dim cmdStr2 As String
Dim cmdStr3 As String


cmdStr = " update tblCustomerDieQty  set DieQty=" & productNewTemp & ",lastupdateby='Auto',lastupdatedate=sysdate  where CustomerName='" & productTemp & "' and Flag='Y' and CustomerPT= '" & testTemp & "' "
                         
                                            
AddSql (cmdStr)




MsgBox "修改成功！"

End Sub


Public Sub UpdatePfData(valueTemp As String, typeTemp As String)
Dim cmdStr As String

cmdStr = "update tblsetpf set fieldvalue='" & valueTemp & "' where resultvalue='" & typeTemp & "' and flag='Y'  "
                         
                         
AddSql (cmdStr)

MsgBox "修改成功！"

End Sub

Public Sub UpdateTrayData(valueTemp As String, typeTemp As String, oldvalueTemp As String)
Dim cmdStr As String

cmdStr = "update TBLSETTray set fieldvalue='" & valueTemp & "' where traytype='" & typeTemp & "' and flag='Y'  and  fieldvalue='" & oldvalueTemp & "' "
                         
                         
AddSql (cmdStr)

MsgBox "修改成功！"

End Sub




Public Sub DeltestNo2(productTemp As String, testTemp As String, oldtestNoTemp As String, productTempNew As String)
Dim cmdStr As String

cmdStr = " update tblTestNo2  set testedition='" & testTemp & "',lastupdateby='Auto',lastupdatedate=sysdate  where productname='" & productTemp & "'  and testedition='" & oldtestNoTemp & "' and Flag='Y' and productname='" & productTempNew & "'"
                         
AddSql (cmdStr)

MsgBox "修改成功！"

End Sub



Public Sub delGCCTData(useridTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String

cmdStr = "   delete from  TSV_GC_CT where createby='" & useridTemp & "'  "
                                                  
AddSql (cmdStr)

End Sub


Public Sub del37CTData(useridTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String

cmdStr = "   delete from  pj_37_ct   "
                                                  
AddSql (cmdStr)

End Sub
Public Sub delQboxNoTw()
Dim cmdStr As String
Dim cmdStr2 As String

cmdStr = " delete from TWBOXTMP"
                                                  
AddSql (cmdStr)

End Sub


Public Sub delCOGIniData()
Dim cmdStr As String
Dim cmdStr2 As String

cmdStr = "   delete from  GR_COG_IniData   "
                                                  
AddSql (cmdStr)

End Sub



Public Sub delCOGRptInt02()
Dim cmdStr As String
Dim cmdStr2 As String

cmdStr = "   delete from  TSV_GR_COG_Tray_Data   "
                                                  
AddSql (cmdStr)

End Sub



Public Sub AddCOGRptInt02()
Dim cmdStr As String
Dim cmdStr2 As String

' cmdStr = "insert into TSV_GR_COG_Tray_Data" & _
'" select GR_COG_Tray_Seq.nextval ID, qboxnumber,containername,barcodeqbox,lotid,qty,GRADE,BOX_TYPE,SERIAL_NUM " & _
'" from (select distinct  a.qboxnumber,a.containername,b.barcodeqbox,substr( b.maincontainername,1,length(b.maincontainername)-4) as lotid,750 as qty ," & _
'" 'A' as GRADE,'LT' as BOX_TYPE,'1' as SERIAL_NUM " & _
'" from  GR_COG_IniData a ,TSV_QBOXTBL_GC_COGSEQ b " & _
'" Where b.maincontainername = a.containername " & _
'" and b.typename='Tray' order by a.qboxnumber,a.containername,b.barcodeqbox,4 ) X "

' tangwei:20171010   750 -700
cmdStr = "insert into TSV_GR_COG_Tray_Data" & _
" select GR_COG_Tray_Seq.nextval ID, qboxnumber,containername,barcodeqbox,lotid,qty,GRADE,BOX_TYPE,SERIAL_NUM " & _
" from (select distinct  a.qboxnumber,a.containername,b.barcodeqbox,substr( b.maincontainername,1,length(b.maincontainername)-5) as lotid,700 as qty ," & _
" 'A' as GRADE,'LT' as BOX_TYPE,'1' as SERIAL_NUM " & _
" from  GR_COG_IniData a ,TSV_QBOXTBL_GC_COGSEQ b " & _
" Where b.maincontainername = a.containername " & _
" and b.typename='Tray' order by a.qboxnumber,a.containername,b.barcodeqbox,4 ) X "



                                                  
AddSql (cmdStr)

End Sub

Public Sub AddCOGRptInt02_2()
Dim cmdStr As String
Dim cmdStr2 As String

'tangwei:20171010   750 -700
' cmdStr = "insert into TSV_GR_COG_Tray_Data" & _
'" select GR_COG_Tray_Seq.nextval ID, X.qboxnumber,X.containername,X.barcodeqbox,barcode,X.qty,X.GRADE,X.BOX_TYPE,X.SERIAL_NUM from ( " & _
'" select distinct  a.qboxnumber,a.containername,b.barcodeqbox,substr(c.beforebarcode ,1,15) barcode ,750 as qty ," & _
'" 'A' as GRADE,'LV' as BOX_TYPE,'1' as SERIAL_NUM " & _
'" from  GR_COG_IniData a ,TSV_QBOXTBL_GC_COGSEQ b ,tsv_gc_cog_qboxdetail c" & _
'" Where b.maincontainername = a.containername " & _
'" and b.typename='Lvbodai' and c.barcodeqbox=b.barcodeqbox and c.typename='Lvbodai' " & _
'" Union select distinct  a.qboxnumber,a.containername,b.barcodeqbox,substr(c.beforebarcode ,17,15) barcode ,750 as qty , " & _
'" 'A' as GRADE,'LV' as BOX_TYPE,'2' as SERIAL_NUM " & _
'" from  GR_COG_IniData a ,TSV_QBOXTBL_GC_COGSEQ b ,tsv_gc_cog_qboxdetail c " & _
'" Where b.maincontainername = a.containername " & _
'" and b.typename='Lvbodai' and c.barcodeqbox=b.barcodeqbox and c.typename='Lvbodai' order by 1,2,3,4 ) X "

cmdStr = "insert into TSV_GR_COG_Tray_Data" & _
" select GR_COG_Tray_Seq.nextval ID, X.qboxnumber,X.containername,X.barcodeqbox,barcode,X.qty,X.GRADE,X.BOX_TYPE,X.SERIAL_NUM from ( " & _
" select distinct  a.qboxnumber,a.containername,b.barcodeqbox,substr(c.beforebarcode ,1,15) barcode ,700 as qty ," & _
" 'A' as GRADE,'LV' as BOX_TYPE,'1' as SERIAL_NUM " & _
" from  GR_COG_IniData a ,TSV_QBOXTBL_GC_COGSEQ b ,tsv_gc_cog_qboxdetail c" & _
" Where b.maincontainername = a.containername " & _
" and b.typename='Lvbodai' and c.barcodeqbox=b.barcodeqbox and c.typename='Lvbodai' " & _
" Union select distinct  a.qboxnumber,a.containername,b.barcodeqbox,substr(c.beforebarcode ,17,15) barcode ,700 as qty , " & _
" 'A' as GRADE,'LV' as BOX_TYPE,'2' as SERIAL_NUM " & _
" from  GR_COG_IniData a ,TSV_QBOXTBL_GC_COGSEQ b ,tsv_gc_cog_qboxdetail c " & _
" Where b.maincontainername = a.containername " & _
" and b.typename='Lvbodai' and c.barcodeqbox=b.barcodeqbox and c.typename='Lvbodai' order by 1,2,3,4 ) X "

                                                  
AddSql (cmdStr)

End Sub


Public Sub AddCOGRptInt02_3()
Dim cmdStr As String
Dim cmdStr2 As String

' tangwei: 1500 - 1400
 cmdStr = "insert into TSV_GR_COG_Tray_Data" & _
" select GR_COG_Tray_Seq.nextval ID, X.qboxnumber,X.containername,X.barcodeqbox2,barcodeqbox,X.qty,X.GRADE,X.BOX_TYPE,X.SERIAL_NUM from ( " & _
" select distinct  a.qboxnumber,a.containername,a.qboxnumber as barcodeqbox2 ,b.barcodeqbox ,1400 as qty , " & _
" 'A' as GRADE,'LB' as BOX_TYPE, ROW_NUMBER() OVER(PARTITION BY a.qboxnumber,a.containername ORDER BY b.barcodeqbox) as SERIAL_NUM from  GR_COG_IniData a ,TSV_QBOXTBL_GC_COGSEQ b " & _
" Where b.maincontainername = a.containername and b.typename='Lvbodai' order by a.qboxnumber,a.containername,b.barcodeqbox ) X "
                                            
AddSql (cmdStr)

End Sub





Public Sub InsertGCCTData(containerTemp As String, waferid As String, dateTemp As String, useridTemp As String, dateTemp2 As String)
Dim cmdStr As String
Dim cmdStr2 As String

cmdStr = "   insert into TSV_GC_CT (wafernumber,productname,customername,createby,FabDate) values ('" & containerTemp & "','" & waferid & "','" & dateTemp & "','" & useridTemp & "','" & dateTemp2 & "')  "
                                                  
AddSql (cmdStr)

End Sub




Public Sub Insert37CTData(dateTemp As String, useridTemp As String, dateTemp2 As String)
Dim cmdStr As String
Dim cmdStr2 As String

cmdStr = "insert into ORACLEDB .. INSITEQT2.PJ_37_CT select rtrim(a.单据编号), rtrim(a.操作日期),rtrim(a.序号), rtrim(x.工单号), rtrim(x.箱号), rtrim(substring(h.CONTAINERNAME, 1, CHARINDEX('-', h.CONTAINERNAME) - 1)),rtrim(x.dn), rtrim(g.收货日期)," & _
" rtrim('07885'), rtrim(getdate()),rtrim(x.入库时间),rtrim(x.料号),rtrim(x.入库数)   from (select e.入库时间, d.箱号, e.入库单编号, e.工单号, c.dn,e.料号,d.入库数  from erpdata .. tblPackToHouse e, erpdata .. tblPackToHouseSub d, erpdata .. tblstocknumtree c where e.客户代码 = '37' " & _
" and CONVERT(varchar(100), e.入库时间, 23) >=CONVERT(varchar(100), '" & dateTemp & "', 23) and CONVERT(varchar(100), e.入库时间, 23) <= CONVERT(varchar(100),  '" & dateTemp2 & "', 23) and e.入库单编号 = d.入库单编号 and d.箱号 = c.箱号  and c.序号 not in (select q.上级序号 from erpdata .. tblstocknumtree q ) ) x " & _
"  left join erpdata .. tblstockmovesub b on  x.箱号 = b.箱号 left join erpdata .. tblstockmove a  on a.单据编号 = b.单据编号 and a.序号 = b.单据项次  and a.单据编号 like 'F%'  left join erpdata..TblQBOXNUMBER_TSV h  on h.QBOXNUMBER = x.箱号 " & _
" left join ERPBASE .. tbltorec_wafer f on f.晶圆ID = substring(h.CONTAINERNAME, 1, CHARINDEX('-', h.CONTAINERNAME) - 1) left join ERPBASE .. tbltorec g on g.到货单编号 = f.到货单编号 "
AddSql2 (cmdStr)

End Sub
Public Sub InsertQboxTemp(dataTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String

cmdStr = "insert into ORACLEDB .. INSITEQT2.TWBOXTMP SELECT distinct rtrim(a.入库单编号),rtrim(b.箱号) as 大箱号, rtrim(c.箱号) as 小箱号 ,Convert(VARCHAR(30), d.入库时间, 0 ) FROM erpdata.dbo.tblPackToHouseSub a," & _
  "erpdata.dbo.tblPackTreeInf    b, erpdata.dbo.tblPackTreeInf    c,erpdata.dbo.tblPackToHouse    d where a.入库单编号 = '" & dataTemp & "'   " & _
  "  and b.箱号 = a.箱号 and c.上级序号 = b.序号 and d.入库单编号 = a.入库单编号"

AddSql2 (cmdStr)

End Sub

Public Sub InsertQboxTempP(startTime As String, stopTime As String)
Dim cmdStr As String
Dim cmdStr2 As String

cmdStr = "insert into ORACLEDB .. INSITEQT2.TWBOXTMP SELECT distinct rtrim(a.入库单编号),rtrim(b.箱号) as 大箱号, rtrim(c.箱号) as 小箱号 ,Convert(VARCHAR(30), d.入库时间, 0 ) FROM erpdata.dbo.tblPackToHouseSub a," & _
  "erpdata.dbo.tblPackTreeInf    b, erpdata.dbo.tblPackTreeInf    c,erpdata.dbo.tblPackToHouse    d where d.入库时间 >= '" & startTime & "' and d.入库时间 <= '" & stopTime & "'  " & _
  "  and b.箱号 = a.箱号 and c.上级序号 = b.序号 and d.入库单编号 = a.入库单编号"

AddSql2 (cmdStr)

End Sub




Public Sub ModifyBC(batchIdTemp As String, NewQty As Long)
Dim cmdStr As String
Dim cmdStr2 As String


'添加导入Sqlserver
On Error GoTo DealError
        
Cnn.BeginTrans

cmdStr = " update CustomerBCtbl  set dieqty=" & NewQty & ",lastupdateby=dieqty-" & NewQty & ",lastupdatedate=sysdate  where BatchId='" & batchIdTemp & "' and Flag='Y' "
cmdStr2 = "update [ERPBASE].[dbo].[tblCustomerBC]  SET dieqty=" & NewQty & " ,lastupdateby=dieqty-" & NewQty & ",lastupdatedate=GETDATE()  where BatchId='" & batchIdTemp & "' and Flag='Y' "

                            
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
Cnn.CommitTrans
MsgBox "修改成功！"


Exit Sub
DealError:

Cnn.RollbackTrans
MsgBox "修改失败！"
 

End Sub




Public Function GetDecBCQty(idTemp As String) As ADODB.Recordset
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select dieqty from  CustomerBCtbl where batchid='" + idTemp + "'  and flag='Y'"
   
Set RSResult = getStr(cmdStr)


Set GetDecBCQty = RSResult
End Function

Public Sub AddGCHeader(mapTemp As GCHeader, id As Long, customerTemp As String)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into CustomerOItbl_test(id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,probe_ship_part_type,jobno, chromaticity, reliability_sampling, lot_priority, wafer_box_type) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.Fab_Device & "'," & _
"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',sysdate,'" & mapTemp.Ship_Out & "','" & mapTemp.TradeType & "','" & mapTemp.TAXTYPE & "', '" & mapTemp.SecondFlag & "', '" & mapTemp.LotProperty & "', '" & mapTemp.LotOwner & "', '" & mapTemp.Telephone & "')"


cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,probe_ship_part_type,jobno,chromaticity,reliability_sampling, lot_priority, wafer_box_type) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.Fab_Device & "'," & _
"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',GETDATE(),'" & mapTemp.Ship_Out & "' ,'" & mapTemp.TradeType & "','" & mapTemp.TAXTYPE & "', '" & mapTemp.SecondFlag & "', '" & mapTemp.LotProperty & "', '" & mapTemp.LotOwner & "', '" & mapTemp.Telephone & "')"


                                               
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
Cnn.CommitTrans

'Exit Sub
'
'DealError:
'
'Cnn.RollbackTrans

End Sub

Public Sub AddGCHeader95FPC(mapTemp As GCHeader, id As Long, customerTemp As String)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into CustomerOItbl_test(id,po_num,source_batch_id,mpn_desc," & _
" flag,CUSTOMERSHORTNAME,Qtech_Created_By,Qtech_Created_Date ) values (" & _
 " '" & id & " ','" & mapTemp.po_no & "','95FPC'," & _
"  '" & mapTemp.Customer_Device & "','Y','95','" & mapTemp.Created_By & "',sysdate)"


cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,source_batch_id,mpn_desc," & _
" CUSTOMERSHORTNAME,flag,Qtech_Created_By,Qtech_Created_Date) values (" & _
" " & id & ",'" & mapTemp.po_no & "','95FPC'," & _
"  '" & mapTemp.Customer_Device & "','95','Y','" & mapTemp.Created_By & "',GETDATE() )"


                                               
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
Cnn.CommitTrans

'Exit Sub
'
'DealError:
'
'Cnn.RollbackTrans

End Sub

'2016-03-23 jiayun add 37Fab
Public Sub Add37FabData(mapTemp As SemtechFabDetail)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into MAPPINGDATA37( ID,DEVICENAME,BATCH ,WF,PRICE ,CURRENCY ,SHIPPEDDT ," & _
" PURCHASENO ,PURCHASEORDERLINEITEM ,INVOICE , MAWBNUMBER  ,DESTINATION  ,WAFER_ID   , FLAG , QTECH_CREATED_BY, QTECH_CREATED_DATE) values (" & _
" " & mapTemp.id & ", '" & mapTemp.DeviceName & "','" & mapTemp.Batch & "','" & mapTemp.WF & "','" & mapTemp.Price & "','" & mapTemp.CURRENCY & "','" & mapTemp.ShippedDt & "'," & _
"  '" & mapTemp.PurchaseNo & "','" & mapTemp.PurchaseOrderLineItem & "','" & mapTemp.Invoice & "','" & mapTemp.MAWBNumber & "','" & mapTemp.Destination & "','" & mapTemp.wafer_id & "','Y','" & mapTemp.QTECH_CREATED_BY & "',sysdate)"


'cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
'" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,probe_ship_part_type) values (" & _
'" " & id & ",'" & mapTemp.PO_NO & "','" & mapTemp.Lot_ID & "','" & mapTemp.Supplier & "','" & mapTemp.ShipTo & "','" & mapTemp.FAB_Device & "'," & _
'"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',GETDATE(),'" & mapTemp.Ship_Out & "' ,'" & mapTemp.TradeType & "' )"


                                               
AddSql (cmdStr)
'AddSql2 (cmdStr2)
 
Cnn.CommitTrans

'Exit Sub
'
'DealError:
'
'Cnn.RollbackTrans

End Sub




'2016-03-23 jiayun add 37Fab
Public Sub AddMPSFabData(mapTemp As SemtechFabDetail)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into MAPPINGDATA37( ID,DEVICENAME,BATCH ,WF,PRICE ,CURRENCY ,SHIPPEDDT ," & _
" PURCHASENO ,PURCHASEORDERLINEITEM ,INVOICE , MAWBNUMBER  ,DESTINATION  ,WAFER_ID   , FLAG , QTECH_CREATED_BY, QTECH_CREATED_DATE,poBatch,customershortname) values (" & _
" " & mapTemp.id & ", '" & mapTemp.DeviceName & "','" & mapTemp.Batch & "','" & mapTemp.WF & "','" & mapTemp.Price & "','" & mapTemp.CURRENCY & "','" & mapTemp.ShippedDt & "'," & _
"  '" & mapTemp.PurchaseNo & "','" & mapTemp.PurchaseOrderLineItem & "','" & mapTemp.Invoice & "','" & mapTemp.MAWBNumber & "','" & mapTemp.Destination & "','" & mapTemp.wafer_id & "','Y','" & mapTemp.QTECH_CREATED_BY & "',sysdate,'" & mapTemp.PoBatch & "','" & mapTemp.custom_Temp & "')"


'cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
'" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,probe_ship_part_type) values (" & _
'" " & id & ",'" & mapTemp.PO_NO & "','" & mapTemp.Lot_ID & "','" & mapTemp.Supplier & "','" & mapTemp.ShipTo & "','" & mapTemp.FAB_Device & "'," & _
'"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',GETDATE(),'" & mapTemp.Ship_Out & "' ,'" & mapTemp.TradeType & "' )"


                                               
AddSql (cmdStr)
'AddSql2 (cmdStr2)
 
Cnn.CommitTrans

'Exit Sub
'
'DealError:
'
'Cnn.RollbackTrans

End Sub

Public Sub AddNormalHeader(mapTemp As GCHeader, id As Long, customerTemp As String)

    '增加抛到SqlServer
    Dim cmdStr  As String

    Dim cmdStr2 As String
                                   
    customerTemp = IIf(InStr(customerTemp, "(2)"), Replace$(customerTemp, "(2)", ""), customerTemp)

    Cnn.BeginTrans

    If customerTemp <> "AA(ON)" Then
    
        mapTemp.Coo = ""
        mapTemp.Level = ""

    End If

    cmdStr = "insert into CustomerOItbl_test(id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
       " CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,probe_ship_part_type,  RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code,shipping_mst_level,shipping_mst_260  ) values (" & _
       " " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.Fab_Device & "'," & _
       "  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',sysdate,'" & mapTemp.Ship_Out & "','" & mapTemp.TradeType & "'," & _
       "  '" & mapTemp.Attri01 & "','" & mapTemp.Attri02 & "','" & mapTemp.Attri03 & "','" & mapTemp.Attri04 & "','" & mapTemp.Attri05 & "','" & mapTemp.taxTemp & "','" & mapTemp.Attri03 & "', '" & mapTemp.Coo & "', '" & mapTemp.Level & "')"

    cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & " CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,probe_ship_part_type,RETICLE_LEVEL_71,RETICLE_LEVEL_72,RETICLE_LEVEL_73,ASSEMBLY_FACILITY,BATCH_COMMENT_ASSY,jobno,date_code ) values (" & " " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.Fab_Device & "'," & "  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',GETDATE(),'" & mapTemp.Ship_Out & "' ,'" & mapTemp.TradeType & "' ," & "  '" & mapTemp.Attri01 & "','" & mapTemp.Attri02 & "','" & mapTemp.Attri03 & "','" & mapTemp.Attri04 & "','" & mapTemp.Attri05 & "','" & mapTemp.taxTemp & "','" & mapTemp.Attri03 & "' )"
                                               
    AddSql (cmdStr)
    AddSql2 (cmdStr2)
 
    Cnn.CommitTrans

End Sub


Public Sub AddMPSHeader(mapTemp As GCHeader, id As Long, customerTemp As String)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into CustomerOItbl_test(id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,probe_ship_part_type) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.Ship_Out & "','" & mapTemp.Fab_Device & "'," & _
"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',sysdate,'" & mapTemp.Ship_Out & "','" & mapTemp.TradeType & "')"


cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,probe_ship_part_type) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.Ship_Out & "','" & mapTemp.Fab_Device & "'," & _
"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',GETDATE(),'" & mapTemp.Ship_Out & "','" & mapTemp.TradeType & "' )"


                                               
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
Cnn.CommitTrans

'Exit Sub
'
'DealError:
'
'Cnn.RollbackTrans

End Sub

'Bom add
Public Sub AddBomHeader(notemp As String, productTemp As String)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   

 cmdStr2 = " insert into [erpdata].[dbo].[TSVtblSetMRule](材料规范编号,工艺,建立日期,状态标记,是否共用标记,物料编号,产线标记) " & _
           "  values ('" & notemp & "','" & gUserName & "',GETDATE(),0,0,'" & productTemp & "',1)"

AddSql2 (cmdStr2)
 


End Sub


Public Sub AddBomChild(notemp As String, productTemp As String, ptTemp As String, qtyTemp As Double, qtyTemp2 As Integer, specTemp As String, typeTemp As String)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   

' cmdStr2 = " insert into [erpdata].[dbo].[TSVtblSetMRule](材料规范编号,工艺,建立日期,状态标记,是否共用标记,物料编号,产线标记) " & _
'           "  values ('" & notemp & "','12541 张源',GETDATE(),0,0,'" & productTemp & "',1)"


 cmdStr2 = " insert into [erpdata].[dbo].[TSVtblMRuleData]( [材料规范编号] , [工序号], [料号], [物料编号], [名称],[规格],[型号],[每只用量],[损耗] ,[单位] " & _
           "  ,[备料用量],[损耗1],[序号],[序号1] ,[材料类型]) " & _
"  select '" & notemp & "' ,'" & productTemp & "',料号, a.物料编号,物料名称,规格型号,型号," & qtyTemp & "," & qtyTemp2 & ",计量单位名称,0,0,2,'" & specTemp & "' ,'" & typeTemp & "' " & _
"  from erpdata.dbo.tblSmainM2  a  where 料号='" & ptTemp & "' "


AddSql2 (cmdStr2)
 


End Sub





Public Sub AddQR2Header(mapTemp As GCHeader, id As Long, customerTemp As String)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into customeroitbl_QR(id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,probe_ship_part_type) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.Fab_Device & "'," & _
"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',sysdate,'" & mapTemp.Ship_Out & "','" & mapTemp.TradeType & "')"
                                          
AddSql (cmdStr)

 
Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans

End Sub



'2015-04-27 CS
Public Sub AddCSHeader(mapTemp As GCHeader, id As Long, customerTemp As String)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into CustomerOItbl_test(id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,date_code,probe_ship_part_type) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.Fab_Device & "'," & _
"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',sysdate,'" & mapTemp.Ship_Out & "','" & mapTemp.DATE_CODE & "','" & mapTemp.TradeType & "')"


cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code,date_code,probe_ship_part_type) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.Fab_Device & "'," & _
"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',GETDATE(),'" & mapTemp.Ship_Out & "' ,'" & mapTemp.DATE_CODE & "','" & mapTemp.TradeType & "')"


                                               
AddSql (cmdStr)
 AddSql2 (cmdStr2)
 
Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans

End Sub


Public Sub AddEQISHeader(mapTemp As EQISHeader, id As Long, customerTemp As String)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

'cmdStr = "insert into CustomerOItbl_test(id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
'" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code) values (" & _
'" " & id & ",'" & mapTemp.PO_NO & "','" & mapTemp.Lot_ID & "','" & mapTemp.Supplier & "','" & mapTemp.ShipTo & "','" & mapTemp.FAB_Device & "'," & _
'"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',sysdate,'" & mapTemp.Ship_Out & "')"
'



cmdStr = " insert into customeroitbl_test( " & _
" id,po_num,source_batch_Id,source_mtrl_num,mtrl_num,mtrl_desc," & _
" test_mtrl_num,test_mtrl_desc,mpn,current_wafer_qty,die_qty,design_id," & _
" reticle_level_71,reticle_level_72,reticle_level_73,batch_comment_probe,assy_process_id,dark_bond_pad_assy," & _
" encoded_mark_id,planned_laser_scribe,package_type,test_site,batch_comment_assy,box_type," & _
" batch_comment_test,created_time,ref_po,date_code,special_process_lot,CUSTOM_PART_NO," & _
" comp_code , flag, qtech_created_by, qtech_created_date, CustomershortName, invflag,MPN_DESC , eqdatacode" & _
" )values(  " & id & ", '" & mapTemp.po_no & "','" & mapTemp.CompleteLotno & "','" & mapTemp.FabLotNo & "','" & mapTemp.WO_NO & "','" & mapTemp.WorkOrder_PartNo & "', " & _
"  '" & mapTemp.SPA & "',  '" & mapTemp.SPADESC & "','" & mapTemp.DieID & "'," & mapTemp.WaferQTY & "," & mapTemp.AssyQty & ",'" & mapTemp.device & "'," & _
"  '" & mapTemp.TSM_C & "','" & mapTemp.TSM_D & "','" & mapTemp.Remarks & "','" & mapTemp.AssemblyDateCode & "','" & mapTemp.Process & "','" & mapTemp.BondingDiagram & "'," & _
"  '" & mapTemp.TSM_A & "','" & mapTemp.TSM_B & "','" & mapTemp.PACKAGE & "','" & mapTemp.Vendor & "','" & mapTemp.Attention & "','" & mapTemp.LabelFormat & "'," & _
"  '" & mapTemp.ESR_No & "','" & mapTemp.Created_Datetime & "','" & mapTemp.waferid & "','" & mapTemp.DATECODE & "','" & mapTemp.ORDERTYPE & "','" & mapTemp.MarketingPartNumber & "'," & _
"  '" & mapTemp.CompanyName & "','Y', '" & mapTemp.Created_By & "',sysdate,'EQ',0,'" & mapTemp.MarketingPartNumber & "','" & mapTemp.DATECODE & "')"




'cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
'" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,comp_code) values (" & _
'" " & id & ",'" & mapTemp.PO_NO & "','" & mapTemp.Lot_ID & "','" & mapTemp.Supplier & "','" & mapTemp.ShipTo & "','" & mapTemp.FAB_Device & "'," & _
'"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',GETDATE(),'" & mapTemp.Ship_Out & "' )"



cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerOI]( " & _
" id,po_num,source_batch_Id,source_mtrl_num,mtrl_num,mtrl_desc," & _
" test_mtrl_num,test_mtrl_desc,mpn,current_wafer_qty,die_qty,design_id," & _
" reticle_level_71,reticle_level_72,reticle_level_73,batch_comment_probe,assy_process_id,dark_bond_pad_assy," & _
" encoded_mark_id,planned_laser_scribe,package_type,test_site,batch_comment_assy,box_type," & _
" batch_comment_test,created_time,ref_po,date_code,special_process_lot,CUSTOM_PART_NO," & _
" comp_code , flag, qtech_created_by, qtech_created_date, CustomershortName,MPN_DESC,eqdatacode " & _
" )values(  " & id & ", '" & mapTemp.po_no & "','" & mapTemp.CompleteLotno & "','" & mapTemp.FabLotNo & "','" & mapTemp.WO_NO & "','" & mapTemp.WorkOrder_PartNo & "', " & _
"  '" & mapTemp.SPA & "',  '" & mapTemp.SPADESC & "','" & mapTemp.DieID & "'," & mapTemp.WaferQTY & "," & mapTemp.AssyQty & ",'" & mapTemp.device & "'," & _
"  '" & mapTemp.TSM_C & "','" & mapTemp.TSM_D & "','" & mapTemp.Remarks & "','" & mapTemp.AssemblyDateCode & "','" & mapTemp.Process & "','" & mapTemp.BondingDiagram & "'," & _
"  '" & mapTemp.TSM_A & "','" & mapTemp.TSM_B & "','" & mapTemp.PACKAGE & "','" & mapTemp.Vendor & "','" & mapTemp.Attention & "','" & mapTemp.LabelFormat & "'," & _
"  '" & mapTemp.ESR_No & "','" & mapTemp.Created_Datetime & "','" & mapTemp.waferid & "','" & mapTemp.DATECODE & "','" & mapTemp.ORDERTYPE & "','" & mapTemp.MarketingPartNumber & "'," & _
"  '" & mapTemp.CompanyName & "','Y', '" & mapTemp.Created_By & "',GETDATE(),'EQ','" & mapTemp.MarketingPartNumber & "','" & mapTemp.DATECODE & "')"

                                               
AddSql (cmdStr)
 AddSql2 (cmdStr2)
 
Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans

End Sub
Public Sub AddEQISHeader_ShippingRequest(mapTemp As EQISHeader)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans



cmdStr = "insert into EQ_SHIPPING_REQUEST ( SUBCONPO,ITEM,QUANTITY,DEVICE,SPA,CSD,LOT,DATECODE1,DELIVERYNAME,DELIVERYADDRESS," & _
         "WAREHOUSE,LOCATION,MODEOFDELIVERY,DATECODE,SO,CARRIERNOTES,LINE,SCHEDULELINE,CUSTPN,COUNTRYANDNAMEOFDISTRIBUTOR," & _
         "CUSTOMER,CUSTOMERPO) VALUES ('" + mapTemp.SUBCONPO + "','" + mapTemp.ITEM + "','" + mapTemp.Quantity + "','" + mapTemp.devicetemp + "' " & _
         " ,'" + mapTemp.SPATemp + "','" + mapTemp.CSD + "','" + mapTemp.lot + "','" + mapTemp.DATECODE1 + "','" + mapTemp.DELIVERYNAME + "' " & _
         " ,'" + mapTemp.DELIVERYADDRESS + "','" + mapTemp.WAREHOUSE + "','" + mapTemp.LOCATION + "','" + mapTemp.MODEOFDELIVERY + "' " & _
         " ,'" + mapTemp.dateCodeTemp + "','" + mapTemp.SO + "','" + mapTemp.CARRIERNOTES + "','" + mapTemp.LINE + "','" + mapTemp.SCHEDULELINE + "' " & _
         " ,'" + mapTemp.CUSTPN + "','" + mapTemp.COUNTRYANDNAMEOFDISTRIBUTOR + "','" + mapTemp.CUSTOMER + "','" + mapTemp.CUSTOMERPO + "')"


                                               
AddSql (cmdStr)

 
Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans

End Sub


Public Sub AddEQHeader(mapTemp As GCHeader, id As Long, customerTemp As String)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into CustomerOItbl_test(id,po_num,source_batch_id,SHIP_SITE,Test_site,mpn,mpn_desc,FAB_CONV_ID,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,encoded_mark_id,planned_laser_scribe,batch_comment_assy,date_code,custom_part_no,EQDATACODE) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.FAB_Device2 & "','" & mapTemp.Customer_Device & "'," & _
"  '" & mapTemp.Fab_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',sysdate,'" & mapTemp.Marking_Lot_ID1 & "','" & mapTemp.Marking_Lot_ID2 & "','" & mapTemp.remarkTemp & "','" & mapTemp.DATE_CODE & "','" & mapTemp.Customer_Device & "','" & mapTemp.Veqdatecode & "')"


cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,source_batch_id,SHIP_SITE,Test_site,mpn,mpn_desc,FAB_CONV_ID,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,encoded_mark_id,planned_laser_scribe,batch_comment_assy,date_code,custom_part_no,EQDATACODE) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.FAB_Device2 & "','" & mapTemp.Customer_Device & "'," & _
"  '" & mapTemp.Fab_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',GETDATE(),'" & mapTemp.Marking_Lot_ID1 & "','" & mapTemp.Marking_Lot_ID2 & "','" & mapTemp.remarkTemp & "','" & mapTemp.DATE_CODE & "','" & mapTemp.Customer_Device & "','" & mapTemp.Veqdatecode & "')"


                                               
AddSql (cmdStr)
 AddSql2 (cmdStr2)
 
Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans

End Sub





Public Sub AddBDHeader(mapTemp As GCHeader, id As Long, customerTemp As String, PShortNameTemp As String)
'增加抛到SqlServer
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into CustomerOItbl_test(id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,mpn ) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.Fab_Device & "'," & _
"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',sysdate,'" & PShortNameTemp & "')"


cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,po_num,source_batch_id,SHIP_SITE,Test_site,FAB_CONV_ID,mpn_desc,Imager_Customer_Rev,Created_Date,mtrl_num," & _
" CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,mpn) values (" & _
" " & id & ",'" & mapTemp.po_no & "','" & mapTemp.Lot_id & "','" & mapTemp.SUPPLIER & "','" & mapTemp.ShipTo & "','" & mapTemp.Fab_Device & "'," & _
"  '" & mapTemp.Customer_Device & "','" & mapTemp.GC_Version & "','" & mapTemp.GC_Date & "','" & mapTemp.WO_NO & "','" & customerTemp & "','Y','" & mapTemp.Created_By & "',GETDATE(),'" & PShortNameTemp & "')"

                                               
AddSql (cmdStr)
 AddSql2 (cmdStr2)
 
Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans

End Sub

Public Sub AddGCDetail(mapTemp As GCDetail, customerTemp As String, oiKeyId As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

cmdStr = "insert into mappingDataTest(id,substrateid,SUBSTRATETYPE,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename )" & _
         " values( mappingData_SEQ.Nextval,'" & mapTemp.ITEM & "','" & gTax & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','Auto',sysdate," & oiKeyId & ")"
                                                               
                                                               
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values('" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','Auto',GETDATE()," & oiKeyId & ")"
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)


If customerTemp = "MG" Then

cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingDataMG] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values('" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','Auto',GETDATE()," & oiKeyId & ")"
         
 AddSql2 (cmdStr2)

End If

' 作数据备份
Dim cmdStr3 As String

cmdStr3 = " insert into wo_data(id, po_num,po_item,source_batch_id,source_mtrl_num,mtrl_num,mtrl_desc,test_mtrl_num,test_mtrl_desc,mpn,mpn_desc,source_mtrl_sloc, " & _
           " mtrl_num_mtrlgrp,probe_ship_part_type,offshore_asm_company,offshore_test_company,current_wafer_qty,die_qty,design_id,country_of_fab,fab_conv_id,fab_excr_id,reticle_level_71, " & _
           " reticle_level_72,reticle_level_73,wafer_size,imager_customer_rev,chromaticity,micro_lens_shift,temperature_spec,prb_containment_type,fabrication_facility,prb_excr_id,batch_comment_probe, " & _
           " assy_process_id,dark_bond_pad_assy,assy_serial_type,sticky_backs_to_save,optical_quality,encoded_mark_id,planned_laser_scribe,package_lid_type,package_type,pb_free_package,target_waf_thickness, " & _
           " reliability_sampling,lot_priority,wafer_box_type,test_site,assembly_facility,batch_comment_assy,tst_process_id,elec_special_test,box_type,protective_film_apld,shipping_mst_260,shipping_mst_level, " & _
           " t_price,ship_comment,batch_comment_test,created_date,created_time,unit_price,ref_po,ref_po_item,country_of_assembly,micron_material,date_code,ship_site,special_process_lot,lot_status,custom_part_no, " & _
           " flag,qtech_created_by,qtech_created_date,qtech_lastupdate_by,qtech_lastupdate_date,customershortname,downqty,invflag,wafer_visual_inspect,comp_code,eqdatacode,jobno,zx_fromsite,zx_invoice, SUBSTRATEID, SUBSTRATETYPE,PRODUCTID,MICRONLOTID,PASSBINCOUNT,FAILBINCOUNT,WAFER_ID,TIME_STATMP)   " & _
           " select   ct.id,ct.po_num,ct.po_item,ct.source_batch_id,ct.source_mtrl_num,ct.mtrl_num,ct.mtrl_desc,ct.test_mtrl_num,ct.test_mtrl_desc,ct.mpn,ct.mpn_desc,ct.source_mtrl_sloc,ct.mtrl_num_mtrlgrp, " & _
           " ct.probe_ship_part_type,ct.offshore_asm_company,ct.offshore_test_company,ct.current_wafer_qty,ct.die_qty,ct.design_id,ct.country_of_fab,ct.fab_conv_id,ct.fab_excr_id,ct.reticle_level_71,ct.reticle_level_72, " & _
           " ct.reticle_level_73,ct.wafer_size,ct.imager_customer_rev,ct.chromaticity,ct.micro_lens_shift,ct.temperature_spec,ct.prb_containment_type,ct.fabrication_facility,ct.prb_excr_id,ct.batch_comment_probe, " & _
           " ct.assy_process_id,ct.dark_bond_pad_assy,ct.assy_serial_type,ct.sticky_backs_to_save,ct.optical_quality,ct.encoded_mark_id,ct.planned_laser_scribe,ct.package_lid_type,ct.package_type,ct.pb_free_package, " & _
           " ct.target_waf_thickness,ct.reliability_sampling,ct.lot_priority,ct.wafer_box_type,ct.test_site,ct.assembly_facility,ct.batch_comment_assy,ct.tst_process_id,ct.elec_special_test,ct.box_type, " & _
           " ct.protective_film_apld,ct.shipping_mst_260,ct.shipping_mst_level,ct.t_price,ct.ship_comment,ct.batch_comment_test,ct.created_date,ct.created_time,ct.unit_price,ct.ref_po,ct.ref_po_item, " & _
           " ct.country_of_assembly,ct.micron_material,ct.date_code,ct.ship_site,ct.special_process_lot,ct.lot_status, " & _
           " ct.custom_part_no,ct.flag,ct.qtech_created_by,ct.qtech_created_date,ct.qtech_lastupdate_by,ct.qtech_lastupdate_date,ct.customershortname,ct.downqty,ct.invflag,ct.wafer_visual_inspect, " & _
           " ct.comp_code,ct.eqdatacode,ct.jobno,ct.zx_fromsite,ct.zx_invoice,mt.SUBSTRATEID, mt.SUBSTRATETYPE,mt.PRODUCTID,mt.MICRONLOTID,mt.PASSBINCOUNT,mt.FAILBINCOUNT,mt.WAFER_ID,sysdate from CustomerOItbl_test ct, MAPPINGDATATEST mt  where mt.substrateid =  '" & mapTemp.ITEM & "' and to_char(ct.id) = mt.filename"


AddSql (cmdStr3)
 
Cnn.CommitTrans

End Sub


Public Sub AddGCDetail95FPC(mapTemp As GCDetail, customerTemp As String, oiKeyId As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
 cmdStr = "insert into mappingDataTest(id,substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values( mappingData_SEQ.Nextval,'" & mapTemp.wafer_id & "','" & mapTemp.Marking_Lot_ID & "','95FPC','" & oiKeyId & "'," & mapTemp.Good_Die_Qty & ",0,'95','Y','Auto',sysdate," & oiKeyId & ")"
                                                               
                                                               
 cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values('" & mapTemp.wafer_id & "','" & mapTemp.Marking_Lot_ID & "','95FPC','" & oiKeyId & "'," & mapTemp.Good_Die_Qty & ",0,'95','Y','Auto',GETDATE()," & oiKeyId & ")"
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)

                                                             
 
Cnn.CommitTrans

'Exit Sub
'
'DealError:
'
'Cnn.RollbackTrans



End Sub


Public Sub Add37FabDetail(mapTemp As SemtechFabDetail, waferIdTemp As String, waferidNoTemp As String, qtyTemp As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
cmdStr = "insert into mappingDataTest(substrateid,lotid,passbincount,failbincount,flag,qtech_created_by,qtech_created_date,wafer_id,customershortname,id,filename )" & _
         " values( '" & waferIdTemp & "','" & mapTemp.Batch & "'," & qtyTemp & ",0,'Y','" & mapTemp.QTECH_CREATED_BY & "',sysdate,'" & waferidNoTemp & "','37',mappingData_SEQ.Nextval,'')"
                         
                                                               
                                                               
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,lotid,passbincount,failbincount,flag,qtech_created_by,qtech_created_date,wafer_id,customershortname,filename )" & _
         " values( '" & waferIdTemp & "','" & mapTemp.Batch & "'," & qtyTemp & ",0,'Y','" & mapTemp.QTECH_CREATED_BY & "',getdate(),'" & waferidNoTemp & "','37','')"
         
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)


 
Cnn.CommitTrans

'Exit Sub
'
'DealError:
'
'Cnn.RollbackTrans



End Sub


Public Sub AddMPSFabDetail(mapTemp As SemtechFabDetail, waferIdTemp As String, waferidNoTemp As String, qtyTemp As Long, customerTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
cmdStr = "insert into mappingDataTest(substrateid,lotid,passbincount,failbincount,flag,qtech_created_by,qtech_created_date,wafer_id,customershortname,id,filename )" & _
         " values( '" & waferIdTemp & "','" & mapTemp.Batch & "'," & qtyTemp & ",0,'Y','" & mapTemp.QTECH_CREATED_BY & "',sysdate,'" & waferidNoTemp & "','" & customerTemp & "',mappingData_SEQ.Nextval,'')"
                         
                                                               
                                                               
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,lotid,passbincount,failbincount,flag,qtech_created_by,qtech_created_date,wafer_id,customershortname,filename )" & _
         " values( '" & waferIdTemp & "','" & mapTemp.Batch & "'," & qtyTemp & ",0,'Y','" & mapTemp.QTECH_CREATED_BY & "',getdate(),'" & waferidNoTemp & "','" & customerTemp & "','')"
         
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)


 
Cnn.CommitTrans

'Exit Sub
'
'DealError:
'
'Cnn.RollbackTrans



End Sub



Public Sub Add37FabICIDetail(lotIdTemp As String, waferIdTemp As String, waferidNoTemp As String, qtyTemp As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
cmdStr = "insert into mappingDataTest(substrateid,lotid,passbincount,failbincount,flag,qtech_created_by,qtech_created_date,wafer_id,customershortname,id,filename )" & _
         " values( '" & waferIdTemp & "','" & lotIdTemp & "'," & qtyTemp & ",0,'Y','Auto',sysdate,'" & waferidNoTemp & "','37',mappingData_SEQ.Nextval,'')"
                         
                                                               
                                                               
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,lotid,passbincount,failbincount,flag,qtech_created_by,qtech_created_date,wafer_id,customershortname,filename )" & _
         " values( '" & waferIdTemp & "','" & lotIdTemp & "'," & qtyTemp & ",0,'Y','Auto',getdate(),'" & waferidNoTemp & "','37','')"
         
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)


 
Cnn.CommitTrans

'Exit Sub
'
'DealError:
'
'Cnn.RollbackTrans



End Sub

Public Sub update37pojobno(lotIdTemp As String, waferIdTemp As String, waferidNoTemp As String, potemp As String, TEMP As SemtechPOHeader)
Dim cmdStr As String
Dim cmdStr2 As String

Cnn.BeginTrans

cmdStr = "update CUSTOMEROITBL_TEST set " & _
      "PO_NUM = '" & TEMP.PurchaseOrderNo & "',PO_ITEM = '" & TEMP.ITEM & "',SOURCE_MTRL_NUM = '" & TEMP.PartNumber & "'," & _
      " MPN = '" & TEMP.ProductionOrderNo & "', MPN_DESC= '" & TEMP.YourMaterialNumber & "',SOURCE_MTRL_SLOC= '" & TEMP.JOBNO & "',OFFSHORE_ASM_COMPANY = '" & TEMP.MfgPlant & "'," & _
       "OFFSHORE_TEST_COMPANY = '" & TEMP.ReceivingPlant & "',CURRENT_WAFER_QTY ='" & TEMP.Quantity & "', COUNTRY_OF_FAB =  '" & TEMP.WaferFAB & "'," & _
       "RETICLE_LEVEL_71= '" & TEMP.Version & "',IMAGER_CUSTOMER_REV  = '" & TEMP.WaferREV & "',PACKAGE_TYPE=  '" & TEMP.TypeService & "',BOX_TYPE= '" & TEMP.UM & "'," & _
       "SHIPPING_MST_260= '" & TEMP.CURRENCY & "', SHIPPING_MST_LEVEL = '" & TEMP.FreightCarrier & "',SHIP_COMMENT = '" & TEMP.TermsDelivery & "',unit_price = '" & TEMP.UnitPrice & "'," & _
       "COUNTRY_OF_ASSEMBLY= '" & TEMP.TermsPayment & "',DATE_CODE = '" & TEMP.DelDate & "',SHIP_SITE = '" & TEMP.ShippingAddress & "',CUSTOM_PART_NO =  '" & TEMP.KeyStr & "'," & _
       "FLAG = 'Y',QTECH_CREATED_BY     = '" & TEMP.QTECH_CREATED_BY & "',QTECH_CREATED_DATE = sysdate,CustomershortName= '37',test_mtrl_desc= '" & TEMP.JOBNO & "'" & _
       "where id in (select c.filename from mappingDataTest c where c.substrateid = '" & waferIdTemp & "') and ( po_num is null or po_num = '' )"


cmdStr2 = "update [ERPBASE].[dbo].[tblCustomerOI] set " & _
      "PO_NUM = '" & TEMP.PurchaseOrderNo & "',PO_ITEM = '" & TEMP.ITEM & "',SOURCE_MTRL_NUM = '" & TEMP.PartNumber & "'," & _
      " MPN = '" & TEMP.ProductionOrderNo & "', MPN_DESC= '" & TEMP.YourMaterialNumber & "',SOURCE_MTRL_SLOC= '" & TEMP.JOBNO & "',OFFSHORE_ASM_COMPANY = '" & TEMP.MfgPlant & "'," & _
       "OFFSHORE_TEST_COMPANY = '" & TEMP.ReceivingPlant & "',CURRENT_WAFER_QTY ='" & TEMP.Quantity & "',COUNTRY_OF_FAB =  '" & TEMP.WaferFAB & "'," & _
       "RETICLE_LEVEL_71= '" & TEMP.Version & "',IMAGER_CUSTOMER_REV  = '" & TEMP.WaferREV & "',PACKAGE_TYPE=  '" & TEMP.TypeService & "',BOX_TYPE= '" & TEMP.UM & "'," & _
       "SHIPPING_MST_260= '" & TEMP.CURRENCY & "', SHIPPING_MST_LEVEL = '" & TEMP.FreightCarrier & "',SHIP_COMMENT = '" & TEMP.TermsDelivery & "',unit_price = '" & TEMP.UnitPrice & "'," & _
       "COUNTRY_OF_ASSEMBLY= '" & TEMP.TermsPayment & "',DATE_CODE = '" & TEMP.DelDate & "',SHIP_SITE = '" & TEMP.ShippingAddress & "',CUSTOM_PART_NO =  '" & TEMP.KeyStr & "'," & _
       "FLAG = 'Y',QTECH_CREATED_BY     = '" & TEMP.QTECH_CREATED_BY & "',QTECH_CREATED_DATE = getdate(),CustomershortName= '37',test_mtrl_desc= '" & TEMP.JOBNO & "'" & _
       "where id in (select c.filename from [ERPBASE].[dbo].[tblmappingData] c where c.substrateid = '" & waferIdTemp & "') and (PO_NUM is null or PO_NUM = '') "
       
AddSql (cmdStr)
AddSql2 (cmdStr2)
SumCount = SumCount + 1

Cnn.CommitTrans

End Sub


Public Sub update37pojobno1(lotIdTemp As String, waferIdTemp As String, waferidNoTemp As String, potemp As String, TEMP As SemtechPOHeader)
Dim cmdStr As String
Dim cmdStr2 As String

Cnn.BeginTrans

cmdStr = "update CUSTOMEROITBL_TEST set " & _
      "PO_NUM = '" & TEMP.PurchaseOrderNo & "',PO_ITEM = '" & TEMP.ITEM & "',SOURCE_MTRL_NUM = '" & TEMP.PartNumber & "'," & _
      " MPN = '" & TEMP.ProductionOrderNo & "', MPN_DESC= '" & TEMP.YourMaterialNumber & "',SOURCE_MTRL_SLOC= '" & TEMP.JOBNO & "',OFFSHORE_ASM_COMPANY = '" & TEMP.MfgPlant & "'," & _
       "OFFSHORE_TEST_COMPANY = '" & TEMP.ReceivingPlant & "',CURRENT_WAFER_QTY ='" & TEMP.Quantity & "',COUNTRY_OF_FAB =  '" & TEMP.WaferFAB & "'," & _
       "RETICLE_LEVEL_71= '" & TEMP.Version & "',IMAGER_CUSTOMER_REV  = '" & TEMP.WaferREV & "',PACKAGE_TYPE=  '" & TEMP.TypeService & "',BOX_TYPE= '" & TEMP.UM & "'," & _
       "SHIPPING_MST_260= '" & TEMP.CURRENCY & "', SHIPPING_MST_LEVEL = '" & TEMP.FreightCarrier & "',SHIP_COMMENT = '" & TEMP.TermsDelivery & "',unit_price = '" & TEMP.UnitPrice & "'," & _
       "COUNTRY_OF_ASSEMBLY= '" & TEMP.TermsPayment & "',DATE_CODE = '" & TEMP.DelDate & "',SHIP_SITE = '" & TEMP.ShippingAddress & "',CUSTOM_PART_NO =  '" & TEMP.KeyStr & "'," & _
       "FLAG = 'Y',QTECH_CREATED_BY     = '" & TEMP.QTECH_CREATED_BY & "',QTECH_CREATED_DATE = sysdate,CustomershortName= '37',test_mtrl_desc= '" & TEMP.JOBNO & "'" & _
       "where source_batch_id = '" & TEMP.LotNO & "' and po_num is null "


cmdStr2 = "update [ERPBASE].[dbo].[tblCustomerOI] set " & _
      "PO_NUM = '" & TEMP.PurchaseOrderNo & "',PO_ITEM = '" & TEMP.ITEM & "',SOURCE_MTRL_NUM = '" & TEMP.PartNumber & "'," & _
      " MPN = '" & TEMP.ProductionOrderNo & "', MPN_DESC= '" & TEMP.YourMaterialNumber & "',SOURCE_MTRL_SLOC= '" & TEMP.JOBNO & "',OFFSHORE_ASM_COMPANY = '" & TEMP.MfgPlant & "'," & _
       "OFFSHORE_TEST_COMPANY = '" & TEMP.ReceivingPlant & "',CURRENT_WAFER_QTY ='" & TEMP.Quantity & "',COUNTRY_OF_FAB =  '" & TEMP.WaferFAB & "'," & _
       "RETICLE_LEVEL_71= '" & TEMP.Version & "',IMAGER_CUSTOMER_REV  = '" & TEMP.WaferREV & "',PACKAGE_TYPE=  '" & TEMP.TypeService & "',BOX_TYPE= '" & TEMP.UM & "'," & _
       "SHIPPING_MST_260= '" & TEMP.CURRENCY & "', SHIPPING_MST_LEVEL = '" & TEMP.FreightCarrier & "',SHIP_COMMENT = '" & TEMP.TermsDelivery & "',unit_price = '" & TEMP.UnitPrice & "'," & _
       "COUNTRY_OF_ASSEMBLY= '" & TEMP.TermsPayment & "',DATE_CODE = '" & TEMP.DelDate & "',SHIP_SITE = '" & TEMP.ShippingAddress & "',CUSTOM_PART_NO =  '" & TEMP.KeyStr & "'," & _
       "FLAG = 'Y',QTECH_CREATED_BY     = '" & TEMP.QTECH_CREATED_BY & "',QTECH_CREATED_DATE = getdate(),CustomershortName= '37',test_mtrl_desc= '" & TEMP.JOBNO & "'" & _
       "where source_batch_id = '" & TEMP.LotNO & "' and ( po_num is null or PO_NUM = '') "
       
AddSql (cmdStr)
AddSql2 (cmdStr2)

SumCount = SumCount + 1

Cnn.CommitTrans

End Sub

Public Sub Add37POwaferDetail(lotIdTemp As String, waferIdTemp As String, waferidNoTemp As String, potemp As String)
Dim cmdStr As String
Dim cmdStr2 As String
       
Cnn.BeginTrans
'''
cmdStr = "insert into mappingdata37po(substrateid,lotid,flag,qtech_created_date,wafer_id,customershortname,substratetype)" & _
         " values( '" & waferIdTemp & "','" & lotIdTemp & "','Y',sysdate,'" & waferidNoTemp & "','37','" & potemp & "')"
                                                                              
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData37PoWafer] (substrateid,lotid,flag,qtech_created_date,wafer_id,customershortname ,substratetype)" & _
         " values( '" & waferIdTemp & "','" & lotIdTemp & "','Y',getdate(),'" & waferidNoTemp & "','37','" & potemp & "')"
         
AddSql (cmdStr)
AddSql2 (cmdStr2)

Cnn.CommitTrans

End Sub



Public Sub Add37POwaferDetail1(lotIdTemp As String, waferIdTemp As String, waferidNoTemp As String, potemp As String, idTemp As String, qtyTemp As String, widtemp As String)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
cmdStr = "insert into mappingdatatest(substrateid,lotid,flag,qtech_created_date,wafer_id,customershortname,filename,passbincount,failbincount)" & _
         " values( '" & widtemp & "','" & lotIdTemp & "','Y',sysdate,'" & waferidNoTemp & "','37','" & idTemp & "','" & qtyTemp & "','0')"
                         
                                                               
                                                               
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,lotid,flag,qtech_created_date,wafer_id,customershortname ,filename,passbincount,failbincount)" & _
         " values( '" & widtemp & "','" & lotIdTemp & "','Y',getdate(),'" & waferidNoTemp & "','37','" & idTemp & "','" & qtyTemp & "','0')"
         
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)


 
Cnn.CommitTrans


End Sub



Public Sub Add68POwaferDetail(lotIdTemp As String, waferIdTemp As String, waferidNoTemp As String, customerTemp As String, goodQtyTemp As Long, idTemp As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
'On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
cmdStr = "insert into mappingdatatest(substrateid,lotid,flag,qtech_created_date,wafer_id,customershortname,passbincount,failbincount,id,filename)" & _
         " values( '" & waferIdTemp & "','" & lotIdTemp & "','Y',sysdate,'" & waferidNoTemp & "','" & customerTemp & "'," & goodQtyTemp & ",0,mappingData_SEQ.Nextval," & idTemp & ")"
                         
                                                               
                                                               
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,lotid,flag,qtech_created_date,wafer_id,customershortname ,passbincount,failbincount,filename)" & _
         " values( '" & waferIdTemp & "','" & lotIdTemp & "','Y',getdate(),'" & waferidNoTemp & "','" & customerTemp & "'," & goodQtyTemp & ",0," & idTemp & ")"
         
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)


 
Cnn.CommitTrans


End Sub








Public Sub AddGCLableWLAWaferid(userNameTemp As String, poidTemp As String, gcdeviceTemp As String, lotIdTemp As String, waferIdTemp As String, woNoTemp As String)
Dim cmdStr As String
                         
                                      
cmdStr = "insert into TSV_GCLable_SETWLA (ID,PO_NO,CustomerDevice,LotID,Waferid,WO_NO,FLAG,CREATEDBY,CREATEDDATE)" & _
         " values( GC_Lable_SpilWla.Nextval,'" & poidTemp & "','" & gcdeviceTemp & "','" & lotIdTemp & "','" & waferIdTemp & "','" & woNoTemp & "','Y','" & userNameTemp & "',sysdate)"
                                                                                           
AddSql (cmdStr)

End Sub



Public Sub Add56Detail(mapTemp As GCDetail, customerTemp As String, oiKeyId As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
cmdStr = "insert into mappingDataTest(id,substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values( mappingData_SEQ.Nextval,'" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & "," & mapTemp.NG_Die_Qty & ",'" & customerTemp & "','Y','Auto',sysdate," & oiKeyId & ")"
                         
                                                               
                                                               
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values('" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & "," & mapTemp.NG_Die_Qty & ",'" & customerTemp & "','Y','Auto',GETDATE()," & oiKeyId & ")"
         
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)


If customerTemp = "MG" Then

cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingDataMG] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values('" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & "," & mapTemp.NG_Die_Qty & ",'" & customerTemp & "','Y','Auto',GETDATE()," & oiKeyId & ")"
         
 AddSql2 (cmdStr2)

End If

                                                             
 
Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans



End Sub



Public Sub AddEQDetail(mapTemp As GCDetail, customerTemp As String, oiKeyId As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
cmdStr = "insert into mappingDataTest(id,substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values( mappingData_SEQ.Nextval,'" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','Auto',sysdate," & oiKeyId & ")"
                         
                                                               
                                                               
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values('" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','Auto',GETDATE()," & oiKeyId & ")"
         
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)
                                                             
 
Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans



End Sub



Public Sub AddGCWLTDetail(mapTemp As GCDetail, customerTemp As String, oiKeyId As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
cmdStr = "insert into mappingDataTest(id,substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename,remark)" & _
         " values( mappingData_SEQ.Nextval,'" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','Auto',sysdate," & oiKeyId & ",'" & mapTemp.REMARK & "')"
                         
                                                               
                                                               
'cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
'         " values('" & mapTemp.item & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_ID & "','" & mapTemp.Wafer_ID & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','Auto',GETDATE()," & oiKeyId & ")"
'
'
                                         
AddSql (cmdStr)
'AddSql2 (cmdStr2)
                                                             
 
Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans



End Sub


Public Sub AddGCDetailZL(mapTemp As GCDetail, customerTemp As String, oiKeyId As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

mapTemp.wafer_id = Replace$(mapTemp.wafer_id, "+", "")

customerTemp = IIf(InStr(customerTemp, "(2)"), Replace$(customerTemp, "(2)", ""), customerTemp)
                                   
                                   
cmdStr = "insert into mappingDataTest(id,substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values( mappingData_SEQ.Nextval,'" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & "," & mapTemp.NG_Die_Qty & ",'" & customerTemp & "','Y','Auto',sysdate," & oiKeyId & ")"
                         
                                                               
                                                               
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values('" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & "," & mapTemp.NG_Die_Qty & ",'" & customerTemp & "','Y','Auto',GETDATE()," & oiKeyId & ")"
         
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)
                                                             
 
' 作数据备份
Dim cmdStr3 As String

cmdStr3 = " insert into wo_data(id, po_num,po_item,source_batch_id,source_mtrl_num,mtrl_num,mtrl_desc,test_mtrl_num,test_mtrl_desc,mpn,mpn_desc,source_mtrl_sloc, " & _
           " mtrl_num_mtrlgrp,probe_ship_part_type,offshore_asm_company,offshore_test_company,current_wafer_qty,die_qty,design_id,country_of_fab,fab_conv_id,fab_excr_id,reticle_level_71, " & _
           " reticle_level_72,reticle_level_73,wafer_size,imager_customer_rev,chromaticity,micro_lens_shift,temperature_spec,prb_containment_type,fabrication_facility,prb_excr_id,batch_comment_probe, " & _
           " assy_process_id,dark_bond_pad_assy,assy_serial_type,sticky_backs_to_save,optical_quality,encoded_mark_id,planned_laser_scribe,package_lid_type,package_type,pb_free_package,target_waf_thickness, " & _
           " reliability_sampling,lot_priority,wafer_box_type,test_site,assembly_facility,batch_comment_assy,tst_process_id,elec_special_test,box_type,protective_film_apld,shipping_mst_260,shipping_mst_level, " & _
           " t_price,ship_comment,batch_comment_test,created_date,created_time,unit_price,ref_po,ref_po_item,country_of_assembly,micron_material,date_code,ship_site,special_process_lot,lot_status,custom_part_no, " & _
           " flag,qtech_created_by,qtech_created_date,qtech_lastupdate_by,qtech_lastupdate_date,customershortname,downqty,invflag,wafer_visual_inspect,comp_code,eqdatacode,jobno,zx_fromsite,zx_invoice, SUBSTRATEID, SUBSTRATETYPE,PRODUCTID,MICRONLOTID,PASSBINCOUNT,FAILBINCOUNT,WAFER_ID,TIME_STATMP)   " & _
           " select   ct.id,ct.po_num,ct.po_item,ct.source_batch_id,ct.source_mtrl_num,ct.mtrl_num,ct.mtrl_desc,ct.test_mtrl_num,ct.test_mtrl_desc,ct.mpn,ct.mpn_desc,ct.source_mtrl_sloc,ct.mtrl_num_mtrlgrp, " & _
           " ct.probe_ship_part_type,ct.offshore_asm_company,ct.offshore_test_company,ct.current_wafer_qty,ct.die_qty,ct.design_id,ct.country_of_fab,ct.fab_conv_id,ct.fab_excr_id,ct.reticle_level_71,ct.reticle_level_72, " & _
           " ct.reticle_level_73,ct.wafer_size,ct.imager_customer_rev,ct.chromaticity,ct.micro_lens_shift,ct.temperature_spec,ct.prb_containment_type,ct.fabrication_facility,ct.prb_excr_id,ct.batch_comment_probe, " & _
           " ct.assy_process_id,ct.dark_bond_pad_assy,ct.assy_serial_type,ct.sticky_backs_to_save,ct.optical_quality,ct.encoded_mark_id,ct.planned_laser_scribe,ct.package_lid_type,ct.package_type,ct.pb_free_package, " & _
           " ct.target_waf_thickness,ct.reliability_sampling,ct.lot_priority,ct.wafer_box_type,ct.test_site,ct.assembly_facility,ct.batch_comment_assy,ct.tst_process_id,ct.elec_special_test,ct.box_type, " & _
           " ct.protective_film_apld,ct.shipping_mst_260,ct.shipping_mst_level,ct.t_price,ct.ship_comment,ct.batch_comment_test,ct.created_date,ct.created_time,ct.unit_price,ct.ref_po,ct.ref_po_item, " & _
           " ct.country_of_assembly,ct.micron_material,ct.date_code,ct.ship_site,ct.special_process_lot,ct.lot_status, " & _
           " ct.custom_part_no,ct.flag,ct.qtech_created_by,ct.qtech_created_date,ct.qtech_lastupdate_by,ct.qtech_lastupdate_date,ct.customershortname,ct.downqty,ct.invflag,ct.wafer_visual_inspect, " & _
           " ct.comp_code,ct.eqdatacode,ct.jobno,ct.zx_fromsite,ct.zx_invoice,mt.SUBSTRATEID, mt.SUBSTRATETYPE,mt.PRODUCTID,mt.MICRONLOTID,mt.PASSBINCOUNT,mt.FAILBINCOUNT,mt.WAFER_ID,sysdate from CustomerOItbl_test ct, MAPPINGDATATEST mt  where mt.substrateid =  '" & mapTemp.ITEM & "' and to_char(ct.id) = mt.filename"


AddSql (cmdStr3)
 
Cnn.CommitTrans


' HT_SO















Exit Sub

DealError:

Cnn.RollbackTrans



End Sub



Public Function ToNumberSystem26(N As Long) As String

Dim s As String
Dim m As Long

N = N + 18278
s = ""

While N > 0

m = N Mod 26

If m = 0 Then
    m = 26
End If

s = Chr(m + 64) + s
N = (N - m) / 26

Wend

ToNumberSystem26 = s

End Function

Public Sub AddQRDetail(mapTemp As GCDetail, customerTemp As String, oiKeyId As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
cmdStr = "insert into mappingDataTest(id,substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values( mappingData_SEQ.Nextval,'" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & "," & mapTemp.NG_Die_Qty & ",'" & customerTemp & "','Y','Auto',sysdate," & oiKeyId & ")"
                         
                                                                                                                      
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values('" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & "," & mapTemp.NG_Die_Qty & ",'" & customerTemp & "','Y','Auto',GETDATE()," & oiKeyId & ")"
         
                                
AddSql (cmdStr)
AddSql2 (cmdStr2)
                                                             

Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans



End Sub

Public Sub AddQR2Detail(mapTemp As GCDetail, customerTemp As String, oiKeyId As Long)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
cmdStr = "insert into mappingData_QR(id,substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & _
         " values( mappingData_SEQ.Nextval,'" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & "," & mapTemp.NG_Die_Qty & ",'" & customerTemp & "','Y','Auto',sysdate," & oiKeyId & ")"
                                                  
AddSql (cmdStr)

Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans

End Sub



Public Sub AddDNDetail(mapTemp As GCDetail, customerTemp As String, oiKeyId As Long, remarkTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String
                                   
On Error GoTo DealError
         
Cnn.BeginTrans

                                   
                                   
cmdStr = "insert into mappingDataTest(id,substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename,Remark)" & _
         " values( mappingData_SEQ.Nextval,'" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','Auto',sysdate," & oiKeyId & ",'" & remarkTemp & "')"
                         
                                                               
                                                               
cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename,Remark)" & _
         " values('" & mapTemp.ITEM & "','" & mapTemp.Marking_Lot_ID & "','" & mapTemp.Lot_id & "','" & mapTemp.wafer_id & "'," & mapTemp.Good_Die_Qty & ",0,'" & customerTemp & "','Y','Auto',GETDATE()," & oiKeyId & ",'" & remarkTemp & "')"
         
                                                               
                                         
AddSql (cmdStr)
AddSql2 (cmdStr2)
                                                             
 
Cnn.CommitTrans

Exit Sub

DealError:

Cnn.RollbackTrans

End Sub

'Insert into DB
Public Function AddSql(cmdStr As String) As Long
If Cnn.State = 0 Then
    ConOracle
End If

cmd.ActiveConnection = Cnn
cmd.CommandText = cmdStr
cmd.CommandType = adCmdText
cmd.Execute SD
AddSql = SD
End Function


Public Sub AddSqlERPInt(cmdStr As String)
If CnnERPInt.State = 0 Then
ConOracle
End If

cmd.ActiveConnection = CnnERPInt
cmd.CommandText = cmdStr
cmd.CommandType = adCmdText
cmd.Execute
    
End Sub

Public Sub ConOracle()
On Error Resume Next
While Cnn.State = 0

    Cnn.Open "Provider=OraOLEDB.Oracle.1;Password=KsMesDB_ht89;User ID=insiteqt2;Data Source=testmes;Persist Security Info=True"
    
    
    If Cnn.State <> 0 Then
        GoTo EndCon
    End If

Wend

While CnnERPInt.State = 0

    CnnERPInt.Open "Provider=OraOLEDB.Oracle.1;Password=erpintegration2;User ID=erpintegration2;Data Source=testmes;Persist Security Info=True"
    
    If CnnERPInt.State <> 0 Then
        GoTo EndCon
    End If

Wend

EndCon:

End Sub
Public Function GetMaxID() As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select CustomerBCtbl_SEQ.nextval ID from dual"
     
RSResult = GetSeq(cmdStr)
GetMaxID = RSResult
End Function


Public Function GetshippingMaxID() As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select CUSTOMERshippingTBL_SEQ.nextval ID from dual"
     
RSResult = GetSeq(cmdStr)
GetshippingMaxID = RSResult
End Function




Public Function Get37FabMaxID() As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select CUSTOMER37FabID_SEQ.nextval ID from dual"
     
RSResult = GetSeq(cmdStr)
Get37FabMaxID = RSResult
End Function



Public Function GetPOPriceID() As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select POPrice_SEQ.nextval ID from dual"
     
RSResult = GetSeq(cmdStr)
GetPOPriceID = RSResult
End Function



Public Function GetMaxIDONMarkCode() As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select CUSTOMERBONMark_SEQ.nextval ID from dual"
     
RSResult = GetSeq(cmdStr)
GetMaxIDONMarkCode = RSResult
End Function




Public Function GetForeCastID() As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select CUSTOMERForCastTBL_SEQ.nextval ID from dual"
     
RSResult = GetSeq(cmdStr)
GetForeCastID = RSResult
End Function



Public Function GetEQShippingMaxID() As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select EQShipping_SEQ.nextval ID from dual"
     
RSResult = GetSeq(cmdStr)
GetEQShippingMaxID = RSResult
End Function


Public Function BaoFeiGetMaxID() As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select BaoFei_SEQ.Nextval ID from dual"
     
RSResult = GetSeq(cmdStr)
BaoFeiGetMaxID = RSResult
End Function


Public Function GetGCLotIDWOId(lotIdTemp As String, woTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select id  from CustomerOItbl_test where source_batch_id='" & lotIdTemp & "' and mtrl_num='" & woTemp & "'"

RSResult = GetSeq(cmdStr)
GetGCLotIDWOId = RSResult
End Function

Public Function GetMCLotIDWOId(lotIdTemp As String, woTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select id  from CustomerOItbl_test where source_batch_id='" & lotIdTemp & "' "

RSResult = GetSeq(cmdStr)
GetMCLotIDWOId = RSResult
End Function

Public Function GetSXLotIDPOId(lotIdTemp As String, potemp As String, ptTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select id  from CustomerOItbl_test where source_batch_id='" & lotIdTemp & "'  and po_num='" & potemp & "' and mpn_desc= '" & ptTemp & "' "

RSResult = GetSeq(cmdStr)
GetSXLotIDPOId = RSResult
End Function

'2015-12-16 jiayun add
Public Function GetPOLotIDPOIdNew(lotIdTemp As String, potemp As String, custPTTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select id  from CustomerOItbl_test where source_batch_id='" & lotIdTemp & "'  and po_num='" & potemp & "' and mpn_desc='" & custPTTemp & "'  "

RSResult = GetSeq(cmdStr)
GetPOLotIDPOIdNew = RSResult
End Function




Public Function GetQR2LotIDPOId(lotIdTemp As String, potemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select id  from customeroitbl_QR where source_batch_id='" & lotIdTemp & "'  and po_num='" & potemp & "'  "

RSResult = GetSeq(cmdStr)
GetQR2LotIDPOId = RSResult
End Function


Public Function GetEQISLotIDPOId(lotIdTemp As String, potemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select id  from CustomerOItbl_test where source_batch_id='" & lotIdTemp & "'  and po_num='" & potemp & "'  "

RSResult = GetSeq(cmdStr)
GetEQISLotIDPOId = RSResult
End Function



Public Function GetAAMaping_GDieQty(lotIdTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = " select passbincount id from mappingdatatest where customershortname='AA' and substrateid='" & lotIdTemp & "' "

RSResult = GetSeq(cmdStr)
GetAAMaping_GDieQty = RSResult
End Function


Public Function Get37DieQty(ptTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = " select passbincount id from mappingdatatest where customershortname='AA' and substrateid='" & lotIDTemp & "' "



'cmdStr = " select a.customerdieqty id  from  TBLTsvNpiProduct a where a.customershortname='37' and a.flag='Y' and (  a.customerptno1= '" & ptTemp & "' or   a.customerptno2= '" & ptTemp & "' " & _
'" or a.customerptno3= '" & ptTemp & "' or   a.customerptno4='" & ptTemp & "' ) and rownum<2 "

cmdStr = " select a.customerdieqty id  from  TBLTsvNpiProduct a , SemtechFabPT b  where a.customershortname='37' and a.flag='Y' and b.flag='Y' and b.fabpt='" & ptTemp & "' " & _
" and (  a.customerptno1=b.npiproductpt or   a.customerptno2= b.npiproductpt or a.customerptno3= b.npiproductpt or   a.customerptno4=b.npiproductpt ) "



RSResult = GetSeq(cmdStr)
Get37DieQty = RSResult
End Function


Public Function Get68DieQty(ptTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = " select passbincount id from mappingdatatest where customershortname='AA' and substrateid='" & lotIDTemp & "' "



'cmdStr = " select a.customerdieqty id  from  TBLTsvNpiProduct a where a.customershortname='37' and a.flag='Y' and (  a.customerptno1= '" & ptTemp & "' or   a.customerptno2= '" & ptTemp & "' " & _
'" or a.customerptno3= '" & ptTemp & "' or   a.customerptno4='" & ptTemp & "' ) and rownum<2 "

cmdStr = " select a.customerdieqty id  from  TBLTsvNpiProduct a  where a.customershortname in ('68','70','HK006','BJ128') and a.flag='Y'  " & _
" and (  a.customerptno1='" & ptTemp & "'  or   a.customerptno2= '" & ptTemp & "' or a.customerptno3= '" & ptTemp & "'  or   a.customerptno4='" & ptTemp & "'  ) "



RSResult = GetSeq(cmdStr)
Get68DieQty = RSResult
End Function







'jiayun add 2015-07-10 市场部订单汇总
Public Function GetQtyMDMonth(facId As Integer, custId As Integer, monthId As Integer, custNameTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = " select Get_MD_RPT_MonthQty(" & facId & "," & custId & "," & monthId & ",'" & custNameTemp & "') id from dual "

RSResult = GetSeq(cmdStr)
GetQtyMDMonth = RSResult
End Function


Public Function GetQtyMDDay(facId As Integer, custId As Integer, monthId As Integer, custNameTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = " select Get_MD_RPT_DayQty(" & facId & "," & custId & "," & monthId & ",'" & custNameTemp & "') id from dual "

RSResult = GetSeq(cmdStr)
GetQtyMDDay = RSResult
End Function

Public Function GetQtyTotalWafer(custPTTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "  select NVL (sum(wafer_qty),0) id from  TblGC_TrayViewWip where target_device like '" & custPTTemp & "%' "

 
RSResult = GetSeq(cmdStr)
GetQtyTotalWafer = RSResult
End Function

Public Function GetQtyTotalWaferV2(custPTTemp As String, beginDateTemp As Date) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "  select NVL (sum(wafer_qty),0) id from  TblGC_TrayViewWip where target_device like '" & custPTTemp & "%' "

cmdStr = "  select NVL (sum(wafer_qty),0) id from  TblGC_WipHistory where target_device like '" & custPTTemp & "%' and event_date=to_date('" & beginDateTemp & "','YYYY-MM-DD') and sub_stage not in ('TEST测试','PACKAGE包装','待入库') "


 
RSResult = GetSeq(cmdStr)
GetQtyTotalWaferV2 = RSResult
End Function



Public Function GetGCTrayThLastWeekQty(custPTTemp As String, dateTemp As Date) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "  select NVL (sum(wafer_qty),0) id from  TblGC_TrayViewWip where target_device like '" & custPTTemp & "%' "

cmdStr = " select NVL (sum(wafer_qty),0) id from TblGC_WipHistory where target_device like '" & custPTTemp & "%'  and event_date=to_date('" & dateTemp & "','YYYY-MM-DD') and sub_stage not in ('TEST测试','PACKAGE包装','待入库')"



RSResult = GetSeq(cmdStr)
GetGCTrayThLastWeekQty = RSResult
End Function


Public Function GetGCTrayThLastWeekQty_Normal(custPTTemp As String, dateTemp As Date) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "  select NVL (sum(wafer_qty),0) id from  TblGC_TrayViewWip where target_device like '" & custPTTemp & "%' "

'cmdStr = " select NVL (sum(wafer_qty),0) id from TblGC_WipHistory where target_device like '" & custPTTemp & "%'  and event_date=to_date('" & dateTemp & "','YYYY-MM-DD') and sub_stage not in ('TEST测试','PACKAGE包装','待入库')"


cmdStr = " select NVL (sum(b.qty),0) id  from TblGC_WipHistory a, TSV_GCTRAY_SetWLA b  where a.target_device like '" & custPTTemp & "%'  and a.event_date=to_date('" & dateTemp & "','YYYY-MM-DD') and a.sub_stage not in ('TEST测试','PACKAGE包装','待入库') " & _
" and b.lotid=a.WAFER_LOT "
 


RSResult = GetSeq(cmdStr)
GetGCTrayThLastWeekQty_Normal = RSResult
End Function




Public Function GetGCTrayThLastWeekWoQty(custPTTemp As String, dateTemp As Date, endDateTemp As Date) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "  select NVL (sum(wafer_qty),0) id from TblGC_TrayViewWip  where target_device like '" & custPTTemp & "%' "

'cmdStr = " select NVL (sum(wafer_qty),0) id from TblGC_WipHistory  where target_device like '" & custPTTemp & "%'  and event_date='" & dateTemp & "' and sub_stage not in ('TEST测试','PACKAGE包装','待入库')"

cmdStr = " select count(b.substrateid) id from customeroitbl_test a , mappingdatatest b  where a.customershortname='GC' and b.customershortname='GC' " & _
         " and b.lotid=a.source_batch_id and b.filename=a.id and a.mpn_desc like '" & custPTTemp & "%' and a.qtech_created_date>='" & dateTemp & "' and a.qtech_created_date<'" & endDateTemp & "'"

RSResult = GetSeq(cmdStr)
GetGCTrayThLastWeekWoQty = RSResult
End Function


Public Function GetGCTrayThInvQty(custPTTemp As String, htPTTemp As String, typeTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "  select NVL (sum(wafer_qty),0) id from TblGC_TrayViewWip  where target_device like '" & custPTTemp & "%' "

'cmdStr = " select NVL (sum(wafer_qty),0) id from TblGC_WipHistory  where target_device like '" & custPTTemp & "%'  and event_date='" & dateTemp & "' and sub_stage not in ('TEST测试','PACKAGE包装','待入库')"

If typeTemp = "GD" Then

cmdStr = " select NVL(sum(goodtrayweekqty),0) id from TSV_GCTrayRptSet_Qtyint b  where b.gcpt='" & custPTTemp & "' and b.htpt='" & htPTTemp & "' "

Else
cmdStr = " select NVL(sum(ngtrayweekqty),0) id from TSV_GCTrayRptSet_Qtyint b  where b.gcpt='" & custPTTemp & "' and b.htpt='" & htPTTemp & "' "


End If



RSResult = GetSeq(cmdStr)
GetGCTrayThInvQty = RSResult
End Function


Public Function GetSemtechWeiPiQty(containerNmeTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "  select NVL (sum(wafer_qty),0) id from TblGC_TrayViewWip  where target_device like '" & custPTTemp & "%' "

'cmdStr = " select NVL (sum(wafer_qty),0) id from TblGC_WipHistory  where target_device like '" & custPTTemp & "%'  and event_date='" & dateTemp & "' and sub_stage not in ('TEST测试','PACKAGE包装','待入库')"


cmdStr = " select NVL(sum(ngtrayweekqty),0) id from TSV_GCTrayRptSet_Qtyint b  where b.gcpt='" & custPTTemp & "' and b.htpt='" & htPTTemp & "' "




RSResult = GetSeq(cmdStr)
GetSemtechWeiPiQty = RSResult
End Function




Public Function GetQty37WoWeiPiQty(lotIdTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

'查询5272站所有数据


'cmdStr = "  select sum(c.qty) as ID from a_lotwafers a, container b, historymainline c " & _
'" where a.wafernumber = '" & lotidTemp & "' and b.containerid = a.containerid " & _
'" and c.containername = b.containername and c.specname = '5272' and c.cdoname = 'MoveInLot'  and b.containername  like '%-A%' "
'
   
cmdStr = " select sum(c.qty) as ID from a_lotwafers a, container b, historymainline c " & _
" where b.containername='" & lotIdTemp & "' and b.containerid = a.containerid " & _
" and c.containername = b.containername and c.specname = '5272' and c.cdoname = 'MoveInLot'  and b.containername  like '%-A%' "
   
   
RSResult = GetSeq(cmdStr)
GetQty37WoWeiPiQty = RSResult
End Function







Public Function GetGCTrayThLastWeekWoQty_Normal(custPTTemp As String, dateTemp As Date, endDateTemp As Date) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "  select NVL (sum(wafer_qty),0) id from TblGC_TrayViewWip  where target_device like '" & custPTTemp & "%' "

'cmdStr = " select NVL (sum(wafer_qty),0) id from TblGC_WipHistory  where target_device like '" & custPTTemp & "%'  and event_date='" & dateTemp & "' and sub_stage not in ('TEST测试','PACKAGE包装','待入库')"

'cmdStr = " select count(b.substrateid) id from customeroitbl_test a , mappingdatatest b  where a.customershortname='GC' and b.customershortname='GC' " & _
'         " and b.lotid=a.source_batch_id and b.filename=a.id and a.mpn_desc like '" & custPTTemp & "%' and a.qtech_created_date>='" & dateTemp & "' and a.qtech_created_date<'" & endDateTemp & "'"


 cmdStr = "select NVL (sum(c.qty),0) id from customeroitbl_test a , mappingdatatest b , TSV_GCTRAY_SetWLA c  where a.customershortname='GC' and b.customershortname='GC' " & _
          " and b.lotid=a.source_batch_id and b.filename=a.id and a.mpn_desc like '" & custPTTemp & "%' and a.qtech_created_date>='" & dateTemp & "' and a.qtech_created_date<'" & endDateTemp & "' " & _
          " and c.lotid=b.lotid "
 


RSResult = GetSeq(cmdStr)
GetGCTrayThLastWeekWoQty_Normal = RSResult
End Function







Public Function GetQtySETNormalWafer(custPTTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "  select NVL (sum(wafer_qty),0) id from  TblGC_TrayViewWip where target_device like '" & custPTTemp & "%' "

cmdStr = "  select NVL (sum(b.qty),0) id  from  TblGC_TrayViewWip a , TSV_GCTRAY_SetWLA b where a.target_device like '" & custPTTemp & "%'  and b.lotid=a.wafer_lot "

 
RSResult = GetSeq(cmdStr)
GetQtySETNormalWafer = RSResult
End Function


Public Function GetQtySETNormalWaferV2(custPTTemp As String, beginDateTemp As Date) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "  select NVL (sum(wafer_qty),0) id from  TblGC_TrayViewWip where target_device like '" & custPTTemp & "%' "

cmdStr = "  select NVL (sum(b.qty),0) id  from  TblGC_WipHistory a , TSV_GCTRAY_SetWLA b where a.target_device like '" & custPTTemp & "%' and a.event_date=to_date('" & beginDateTemp & "','YYYY-MM-DD')  and  a.sub_stage not in ('TEST测试','PACKAGE包装','待入库') and b.lotid=a.wafer_lot "

RSResult = GetSeq(cmdStr)
GetQtySETNormalWaferV2 = RSResult
End Function

Public Function GetSpeGRNGDieQty(lotIdTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = " SELECT MAX(a.die_qty)-ISNULL(SUM(CAST(b.Consumed_Qty AS INT)),0) ID " & _
" FROM [ERPBASE].[dbo].[tblCustomerOI] a  LEFT JOIN [erpdata].[dbo].[GRdetailHistory] b ON a.SOURCE_BATCH_ID=b.Previous_Batch_ID " & _
"  WHERE a.SOURCE_BATCH_ID='" & lotIdTemp & "' "
   
RSResult = GetServerSeq(cmdStr)
GetSpeGRNGDieQty = RSResult
End Function

Public Function GetSpeGRNGPieceQty(lotIdTemp As String) As Double
Dim cmdStr As String
Dim RSResult As Double

cmdStr = " SELECT MAX(a.CURRENT_WAFER_QTY)-ISNULL(SUM(CAST(b.Current_Wafer_Qty AS float)),0) ID " & _
" FROM [ERPBASE].[dbo].[tblCustomerOI] a  LEFT JOIN [erpdata].[dbo].[GRdetailHistory] b ON a.SOURCE_BATCH_ID=b.Previous_Batch_ID " & _
"  WHERE a.SOURCE_BATCH_ID='" & lotIdTemp & "' "
   
RSResult = GetServerSeqDouble(cmdStr)
GetSpeGRNGPieceQty = RSResult
End Function



Public Function GetSXCodeID() As String
Dim cmdStr As String
Dim RSResult As String

cmdStr = "SELECT QTMCodeSeq.SXCode('a')  FROM DUAL "
     
RSResult = getStr2(cmdStr)
GetSXCodeID = RSResult
End Function

Public Function GetGTCodeID() As String
Dim cmdStr As String
Dim RSResult As String

cmdStr = "SELECT QTMCodeSeq.GTCode4('a') FROM DUAL"

RSResult = getStr2(cmdStr)
GetGTCodeID = RSResult

End Function

Public Function chkMarkingCodeLen(dT As tyWO) As Boolean
Dim iLen    As Integer
Dim iLenNPI As Integer

chkMarkingCodeLen = False

Select Case dT.CUSTOMER_CODE

    Case "GD108", "HK080"
        If Len(dT.MARKING_CODE) > 0 Then
            chkMarkingCodeLen = True
            Exit Function

        End If

End Select

iLen = Len(Trim(dT.MARKING_CODE))
iLenNPI = Get_OracleNo("select distinct MARKING_CODE from tbltsvnpiproduct where customershortname = '" & dT.CUSTOMER_CODE & "' and customerptno1 = '" & dT.Customer_Device & "' and MARKING_CODE is not null")

'If iLenNPI <> iLen Then
'    MsgBox "NPI对照表维护的打标码位数: " & iLenNPI & " 本次打标码位数: " & iLen & vbCrLf & "打标位数不匹配, 请重新确认", vbCritical, "警告"
'    Exit Function
'End If
Select Case dT.CUSTOMER_CODE

    Case "QR"

        Select Case dT.Customer_Device

            Case "MT01"
                If iLen <> 5 Then
                    MsgBox "打标码位数不正确", vbInformation, "警告"
                    Exit Function

                End If

        End Select

    Case "KR001"

        Select Case dT.HT_DEVICE

            Case "XKR00103"
                If iLen <> 12 Then
                    MsgBox "打标码位数不正确", vbInformation, "警告"
                    Exit Function

                End If

            Case esle
                If iLen <> 13 Then
                    MsgBox "打标码位数不正确", vbInformation, "警告"
                    Exit Function

                End If

        End Select

End Select

chkMarkingCodeLen = True

End Function

Public Function GetKRMark(LOTID As String, waferid As String) As String

Dim codeY As String
Dim codeWW As String
Dim codeA As String
Dim codeLotId As String
Dim codeWaferId As String
Dim codeRes As String
Dim codeTw As String
Dim codePos As String

codeTw = "0123456789ABCDEFGHJKLMNPR"
codeY = Mid(Year(Now), 4, 1)
codeWW = Right("0" & DatePart("ww", Now), 2)
codePos = Replace(codeTw, "0", "")
codeTw = Mid(codeTw, waferid, 1)

codeRes = codeY + codeWW & "A" + LOTID & "H" + codeTw

GetKRMark = codeRes

End Function

Public Function GetKRMarkP(LOTID As String, waferid As String) As String
Dim codeY       As String
Dim codeWW      As String
Dim codeA       As String
Dim codeLotId   As String
Dim codeWaferId As String
Dim codeRes     As String
Dim codeTw      As String
Dim codePos     As String

codeTw = "0123456789ABCDEFGHJKLMNPR"
codeY = Right("00" & Year(Now), 2)
codeWW = Right("00" & DatePart("ww", Now), 2)
codePos = Replace(codeTw, "0", "")
codeTw = Mid(codeTw, waferid, 1)
codeRes = codeY & codeWW & "A" & "H" + LOTID + codeTw
GetKRMarkP = codeRes

End Function


Public Function GetSX8CodeID(lotIdTemp As String, waferIdTemp As String) As String
Dim cmdStr As String
Dim RSResult As String

cmdStr = "SELECT QTMCodeSeq.SXCode8('" & lotIdTemp & "','" & waferIdTemp & "')  FROM DUAL "
     
RSResult = getStr2(cmdStr)
GetSX8CodeID = RSResult
End Function


Public Function GetSpecilGRDt(lotIdTemp As String) As String
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select  to_char(b.erpcreatedate, 'yyyy') ||   to_char(b.erpcreatedate, 'WW') dt  from ib_waferlist a,ib_wohistory b where a.waferlot='" & lotIdTemp & "' and b.ordername=a.ordername and rownum<2"
     
RSResult = getStr2(cmdStr)
GetSpecilGRDt = RSResult
End Function


Public Function GetSpecilGRTestVer(lotIdTemp As String) As String
Dim cmdStr As String
Dim RSResult As String

 cmdStr = "  select  Get_SpecilGR_TESTVER('" & lotIdTemp & "') testver from dual"

     
RSResult = getStr2(cmdStr)
GetSpecilGRTestVer = RSResult
End Function



Public Function GetWLAQbox(lotIdTemp As String) As String
Dim cmdStr As String
Dim RSResult As String

 cmdStr = "  select   trglabelseq.QTSeq_WLA('" & lotIdTemp & "') testver from dual"

     
RSResult = getStr2(cmdStr)
GetWLAQbox = RSResult
End Function




Public Function GetUploadMaxID() As Long
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select TSV_Wo_Mt_UploadTemp_seq.nextval ID from dual"
     
RSResult = GetSeq(cmdStr)
GetUploadMaxID = RSResult
End Function


Public Function Get37MerLotCounts(containerTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long


cmdStr = " select count(distinct b.wafernumber) as ID from container a ,a_lotwafers b " & _
" Where b.containerID = a.containerID and a.containername='" & containerTemp & "' "

     
RSResult = GetSeq(cmdStr)
Get37MerLotCounts = RSResult
End Function

Public Function GetSeq(cmdStr As String) As Long
    Dim resut As New ADODB.Recordset
    
    If Cnn.State = 0 Then
        ConOracle
    End If
    resut.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    GetSeq = CLng(resut.Fields("ID").Value)
End Function
Public Function GetSeqTW(cmdStr As String) As Long
    Dim resut As New ADODB.Recordset
    
    If Cnn.State = 0 Then
        ConOracle
    End If
    resut.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If resut.RecordCount > 0 Then
        If IsNull(resut.Fields(0).Value) Then
            GetSeqTW = 0
        Else
            GetSeqTW = CLng(resut.Fields(0).Value)
        End If
    
    End If
End Function

Public Function getStr2(cmdStr As String) As String
    Dim resut As New ADODB.Recordset
    
    If Cnn.State = 0 Then
        ConOracle
    End If
    resut.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If resut.RecordCount > 0 Then
    getStr2 = IIf(IsNull(resut.Fields(0).Value), "", resut.Fields(0).Value)
    Else
    getStr2 = ""
    End If
End Function



Public Function ExportFpspreadToExcel( _
    Fps As FPSpreadADO.fpSpread, _
    Optional ByVal as_FileName As String = "导出Excel", _
    Optional ByVal as_ReportTitle As String = "资料信息", _
    Optional ByVal as_Memo As String = "", _
    Optional ByVal ab_NoPageSetup As Boolean = True, _
    Optional ByVal as_Company As String = "1") As Boolean
Dim strFileName                         As String
Dim strCompanyFullName                  As String
Dim strServerDate                       As String
Dim lngMaxRow                           As Long
Dim lngMaxCol                           As Long
Dim lngErrNum                           As Long
Dim strChar                             As String
Dim strTmp                              As String
Dim blnSuccess                          As Boolean
Dim intCount                            As Integer
Dim i                                   As Integer
Dim j                                   As Integer
'Dim clsP                                As New ClsProgress
    
    On Error GoTo ErrHandle
    
  
    
    
    
    
    ExportFpspreadToExcel = False
    
    
    With Fps
        lngMaxCol = 0
        For j = 1 To .MaxCols
            .Col = j
            If Not .ColHidden Then
                lngMaxCol = lngMaxCol + 1
            End If
        Next j
        lngMaxRow = .MaxRows
    End With
    
    If lngMaxRow <= 0 Or lngMaxCol <= 0 Then
        MsgBox "没有导出的资料", vbOKOnly Or vbInformation, "提示"
        Exit Function
    End If
    
    strFileName = GetSaveFile(as_FileName)
    If Len(strFileName) = 0 Then
        Exit Function
    Else
        If Len(Dir(strFileName)) > 0 Then
            If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
                Exit Function
            Else
                On Error Resume Next
                Kill strFileName
                If Err.number <> 0 Then
                    MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                    Exit Function
                End If
            End If
        End If
    End If
    
    'clsP.Init lngMaxRow, True
    'clsP.ShowProgress 0, "正在初始化数据..."
    
    Dim ExApp                               As Excel.Application
    Dim wkbk                                As New Workbook
    Dim wkst                                As New Worksheet
    
'    Screen.MousePointer = 11
    Set ExApp = New Excel.Application
    ExApp.Visible = False
    
    Set wkbk = ExApp.Workbooks.Add
    Set wkst = wkbk.Worksheets.Add
    
    strServerDate = Now()
    
    i = lngMaxCol
    If i > 26 Then
        strChar = Chr(96 + (i \ 26)) & IIf(i Mod 26 = 0, "Z", Chr(96 + (i Mod 26)))
    Else
        strChar = Chr(96 + i)
    End If
    
'    With wkst.Range("A1:" & strChar & "1")
'        .Merge
'        .Value = strCompanyFullName
'        .Font.Size = 15
'        .RowHeight = 20.25
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlCenter
'    End With
    With wkst.Range("A1:" & strChar & "2")
        .Merge
        .Font.Size = 18
        .Font.Bold = True
        .RowHeight = 24.5
        .horizontalAlignment = xlHAlignCenterAcrossSelection
        .verticalAlignment = xlCenter
        .Value = "" & as_ReportTitle & ""
    End With
    With wkst.Range("A3:" & strChar & "3")
        .Merge
        .Font.Size = 12
        If Len(as_Memo) > 0 Then
            .Value = "(报表日期:" & FormatDateTime(strServerDate, vbLongDate) & "   共" & lngMaxRow & "笔资料)  (" & as_Memo & ")"
        Else
            .Value = "(报表日期:" & FormatDateTime(strServerDate, vbLongDate) & "   共" & lngMaxRow & "笔资料)"
        End If
        .horizontalAlignment = xlCenter
        .verticalAlignment = xlCenter
        .RowHeight = 15.5
    End With
    '础瓜
'    On Error Resume Next
'    If Len(clsLtd.LogoFileName) > 0 Then
'        wkst.Shapes.AddPicture clsLtd.LogoFileName, _
'                False, True, 10, 0, 85, 50
'    End If
'    On Error Resume Next
'    SavePicture LoadResPicture(101, vbResBitmap), DirShare & "\wus.bmp"
'    wkst.Shapes.AddPicture DirShare & "\wus.bmp", False, True, 10, 0, 85, 55
    
    On Error GoTo ErrHandle
    With wkst.Range("4:4")
        .RowHeight = 4.8
    End With
    
    With Fps
        For i = 0 To .MaxRows
            'clsP.ShowProgress i, "正在导出数据..."
            .Row = i
            intCount = 0
            For j = 1 To .MaxCols
                .Col = j
                If Not .ColHidden Then
                    intCount = intCount + 1
                    wkst.Cells(5 + i, intCount) = .text
                End If
            Next j
        Next i
    End With
    
    If i > lngMaxRow Then
        i = lngMaxRow
    End If
    
    With wkst.Range("A5:" & strChar & "5")
        .Font.Size = 12
        .Font.Bold = True
        .RowHeight = 16.25
        .horizontalAlignment = xlCenter
        .verticalAlignment = xlCenter
    End With
    
    
'    With wkst.Range("A" & i + 7 & ":" & strChar & i + 7)
'        .Merge
'        .Value = Space(38) & ":" & Space(30) & "恨:" & Space(30) & ":" & Erp_XM
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlCenter
'    End With

    With wkst.Range("A6:" & strChar & i + 7)
        .Select
        .Font.Size = 10
        .RowHeight = 14.25
        '.HorizontalAlignment = xlCenter
        .verticalAlignment = xlCenter
    End With
    wkst.Range(i + 6 & ":" & i + 6).RowHeight = 4.8
    wkst.Range("A1").Select
    ExApp.Columns.AutoFit
    
    If Not ab_NoPageSetup Then
        On Error Resume Next
        With wkst.PageSetup
            .LeftMargin = .ExApp.InchesToPoints(0)
            .RightMargin = .ExApp.InchesToPoints(0)
            .TopMargin = .ExApp.InchesToPoints(0)
            .BottomMargin = .ExApp.InchesToPoints(0)
            .HeaderMargin = .ExApp.InchesToPoints(0)
            .FooterMargin = .ExApp.InchesToPoints(0)
            .PrintHeadings = False
            .PrintGridlines = True
            .PrintComments = xlPrintInPlace
            .PrintQuality = 1200
            .CenterHorizontally = True
            .CenterVertically = False
            .Orientation = xlPortrait
            .Draft = False
            .PaperSize = xlPaperA4
            .FirstPageNumber = xlAutomatic
            .order = xlDownThenOver
            .BlackAndWhite = True
            .Zoom = 100
        End With
        On Error GoTo ErrHandle
    End If
        
    wkbk.SaveAs strFileName, xlNormal, "", "", False, False
    wkbk.Saved = True
    
    If MsgBox("导出完毕，是否要打开文件？" & Space(20), vbYesNo Or vbInformation Or vbSystemModal, "提示") = vbYes Then
        ExApp.Visible = True
    Else
        ExApp.Workbooks.Close
    End If
'    Screen.MousePointer = 0
    ExportFpspreadToExcel = True

EXITPRO:
    On Error Resume Next
    'clsP.UnLoad_Form
'    Screen.MousePointer = 0
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not clsP Is Nothing Then
'        Set clsP = Nothing
'    End If
    Exit Function
ErrHandle:
    GoTo EXITPRO
End Function
'*****************************End***************************************



Public Function GetSaveFile( _
        Optional filename As String = "ゅン", _
        Optional Filter As String = "" _
        ) As String
Dim strFileName As String
Dim strDlgTitle As String
Dim strFilter As String
Dim OFile As OPENFILENAME
Dim lngReturn As Long

Static S_OldInitPath As String

GetSaveFile = ""
On Error Resume Next
If Filter = "" Then
    Filter = "Excel*.xls" & Chr(0) & "*.xls"
End If

strFileName = LeftB(filename & Space(MAX_PATH), MAX_PATH - 1)
If Len(S_OldInitPath) <= 0 Then
    S_OldInitPath = App.Path
End If
strDlgTitle = App.Title
strFilter = Filter & Chr(0) & "All Files (*.*) " & Chr(0) & "*.*"


With OFile
    .lStructSize = Len(OFile)
    .hInstance = App.hInstance
    .lpstrFile = strFileName
    .nMaxFile = MAX_PATH
    .lpas_ReportTitle = strDlgTitle
    .nMaxFileTitle = MAX_PATH
    .lpstrFilter = strFilter
    .lpstrInitialDir = S_OldInitPath
    .lpstrDefExt = Right(Filter, 3)
    .nFilterIndex = 1
    .flags = 0
End With

lngReturn = GetSaveFileName(OFile)
If lngReturn = 0 Then
    GetSaveFile = ""
Else
    S_OldInitPath = Trim(OFile.lpstrFile)
    S_OldInitPath = Left(S_OldInitPath, InStrRev(Trim(S_OldInitPath), "\"))
    If Right(S_OldInitPath, 1) <> "\" Then
        S_OldInitPath = S_OldInitPath & "\"
    End If
    GetSaveFile = Trim(OFile.lpstrFile)
End If
End Function

Public Function ExporToExcel(strOpen As String)

Dim Rs_Data As New ADODB.Recordset
Dim Irowcount As Long
Dim Icolcount As Integer
    
    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlQuery As Excel.QueryTable
    
      If Cnn.State = 0 Then
        ConOracle
      End If
    
    With Rs_Data
        If .State = adStateOpen Then
            .Close
        End If
        .ActiveConnection = Cnn
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = strOpen
        .Open
    End With
    With Rs_Data
        If .RecordCount < 1 Then
            MsgBox "查询不到数据!", vbInformation, "提示"
            Exit Function
        End If

        Irowcount = .RecordCount
        Icolcount = .Fields.count
    End With
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlApp.Visible = True

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
    
    xlQuery.FieldNames = True
    xlQuery.Refresh
    
    With xlSheet
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.name = "堵蔨"
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
    End With
    
    xlApp.Application.Visible = True
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Function


Public Function SqlServerExporToExcel(strOpen As String)
'*********************************************************
'* 嘿ExporToExcel
'* 旧计誹EXCEL
'* ノ猭ExporToExcel(sql琩高才﹃)
'*********************************************************
' 导出SqlServer Excel

Dim Rs_Data As New ADODB.Recordset
Dim Irowcount As Long
Dim Icolcount As Integer
    
    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlQuery As Excel.QueryTable
    
'      If Cnn.State = 0 Then
'        ConOracle
'      End If
      
    If INIadoCon.State = 0 Then
        INIConnectSTART
    End If


    
    
   ' CONN_TO_ORACLE_DATABASE2
    
    With Rs_Data
        If .State = adStateOpen Then
            .Close
        End If
        .ActiveConnection = INIadoCon
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
    xlApp.Visible = True
    
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
    End With
    
    With xlSheet.PageSetup
       ' .LeftHeader = "" & Chr(10) & "&""发蔨_GB2312,盽?""&10そ嘿"   ' & Gsmc
        '.CenterHeader = "&""发蔨_GB2312,盽砏""蹦潦快龟兜&""Ш蔨,盽砏""" & Chr(10) & "&""发蔨_GB2312,盽?""&10ら 戳"
       ' .RightHeader = "" & Chr(10) & "&""发蔨_GB2312,盽砏""&10虫涵"
      '  .LeftFooter = "&""发蔨_GB2312,盽砏""&10"
       ' .CenterFooter = "&""发蔨_GB2312,盽砏""&10ら戳"
       ' .RightFooter = "&""发蔨_GB2312,盽?""&10材&P&N"
    End With
    
    xlApp.Application.Visible = True
    Set xlApp = Nothing  '"ユ临北倒Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Function


Public Function SqlServerExporToExcel_WLA(strOpen As String)
'*********************************************************
'* 嘿ExporToExcel
'* 旧计誹EXCEL
'* ノ猭ExporToExcel(sql琩高才﹃)
'*********************************************************
' 导出SqlServer Excel

Dim Rs_Data As New ADODB.Recordset
Dim Irowcount As Long
Dim Icolcount As Integer
    
    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlQuery As Excel.QueryTable
    
    If INIadoCon.State = 0 Then
        INIConnectSTART
    End If


    
    With Rs_Data
        If .State = adStateOpen Then
            .Close
        End If
        .ActiveConnection = INIadoCon
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
    xlApp.Visible = True
    
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
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.name = "宋体"
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Size = 11
        
        '砞竚堵蔨
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = False
        '砞竚蔨彩
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
        '砞竚娩妓Α
    End With

    
    xlApp.Application.Visible = True
    Set xlApp = Nothing  '"ユ临北倒Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Function

Public Function SqlServer2ExporToExcel(strOpen As String)
'*********************************************************
'* 嘿ExporToExcel
'* 旧计誹EXCEL
'* ノ猭ExporToExcel(sql琩高才﹃)
'*********************************************************
' 导出SqlServer Excel

Dim Rs_Data As New ADODB.Recordset
Dim Irowcount As Long
Dim Icolcount As Integer
    
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
    xlApp.Visible = True
    
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
    End With
    
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




Public Function JudgeFlagStauts(id As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select *  from mappingDataTest where substrateid='" + id + "' and flag='Y'"

         
slectResult = QueryStr(cmdStr)
JudgeFlagStauts = slectResult
End Function

Public Function JudgeMPSInvStatus(id As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select *  from CUSTOMERMarkingCode where mpnpart='" + id + "' and flag='Y'"

         
slectResult = QueryStr(cmdStr)
JudgeMPSInvStatus = slectResult
End Function



Public Function JudgeDummyWo1Stauts(woTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from  ib_workorder  where  ordername='" + woTemp + "' "
       
slectResult = QueryStr(cmdStr)
JudgeDummyWo1Stauts = slectResult
End Function

Public Function JudgeDummyWo2Stauts(woTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from  ib_wohistory  where  ordername='" + woTemp + "' "
       
slectResult = QueryStr(cmdStr)
JudgeDummyWo2Stauts = slectResult
End Function



Public Function JudgeFlagStautsMapping2(id As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select *  from TSV_Mapping_tbl where substrateid='" + id + "' and flag='Y'"

         
slectResult = QueryStr(cmdStr)
JudgeFlagStautsMapping2 = slectResult
End Function



Public Function JudgeFlagStautsOI(id As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeFlagStautsOI = slectResult
End Function


Public Function JudgeFlagStautsShipingUp(id As String, itemTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select *  from CUSTOMERSHIPPINGUPTBL where Delivery='" + id + "'  and itemno='" + itemTemp + "' and flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeFlagStautsShipingUp = slectResult
End Function


Public Function JudgeFlagStautsShipingUpjob(id As String, dnjobtemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select *  from CUSTOMERSHIPPINGUPTBL where Delivery='" + id + "'  and batchnumber='" + dnjobtemp + "' and flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeFlagStautsShipingUpjob = slectResult
End Function


Public Function JudgeFlagStautsBI(id As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select *  from CustomerBItbl where batch_id='" + id + "' and flag='Y'"



slectResult = QueryStr(cmdStr)
JudgeFlagStautsBI = slectResult
End Function



Public Function JudgeOracleWipWo(woTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
'根据工单号，判断在wip是否还有数据

slectResult = False

cmdStr = " SELECT  lw.wafernumber,lw.waferscribenumber FROM CONTAINER A " & _
 " inner join  CURRENTSTATUS   B on A.CurrentStatusId = B.CurrentStatusId " & _
"  inner join SPEC            S on  B.SpecId = S.SpecId and S.ObjectCategory = 'WIP' " & _
"  inner join A_SCHEDULEDATA  SD on A.ScheduleDataId = SD.ScheduleDataId and SD.ObjectType = 'WAFER' " & _
"  inner join PRODUCT         P on A.ProductId = P.ProductId and P.objecttype = 'PN' " & _
"  inner join PRODUCTBASE     PB on P.PRODUCTBASEID = PB.PRODUCTBASEID " & _
"  inner join SPECBASE        SB on S.SpecBaseId = SB.SpecBaseId " & _
"  inner join A_LOTWAFERS     lw on a.CONTAINERID = lw.CONTAINERID " & _
"  inner join OPERATION       op on s.OPERATIONID = op.OPERATIONID " & _
"  inner join WORKCENTER      wc on op.WORKCENTERID = wc.WORKCENTERID " & _
"  inner join A_LotAttributes al on a.CONTAINERID = al.CONTAINERID " & _
"  inner join mfgorder        mfg on a.mfgorderid = mfg.mfgorderid " & _
" Where A.Status = 1  and lw.workordername='" + woTemp + "' "
 

slectResult = QueryStr(cmdStr)
JudgeOracleWipWo = slectResult
End Function

Public Function JudgeOracleCloseWo(woTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
'根据工单号，判断在wip是否还有数据

slectResult = False

cmdStr = " select * from ib_waferlist where ordername='" + woTemp + "' "
 
slectResult = QueryStr(cmdStr)
JudgeOracleCloseWo = slectResult
End Function









Public Function JudgeBCExist(id As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select *  from CustomerBCtbl where batchid='" + id + "' and flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeBCExist = slectResult
End Function


Public Function JudgetestNoExist(productTemp As String, productTempNew As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select *   from  tblTestNo where  productname='" + productTemp + "' and flag='Y' and productnamenew='" + productTempNew + "'"



slectResult = QueryStr(cmdStr)
JudgetestNoExist = slectResult
End Function


Public Function JudgeGCLableLotIDData(lotIdTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from TSV_GCLable_SETWLA  where lotid='" + lotIdTemp + "'"




slectResult = QueryStr(cmdStr)
JudgeGCLableLotIDData = slectResult
End Function

Public Function JudgeGCLableWaferIDData(lotIdTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from TSV_GCLable_SETWLA  where Waferid='" + lotIdTemp + "'"


slectResult = QueryStr(cmdStr)
JudgeGCLableWaferIDData = slectResult
End Function




Public Function JudGCDieNoExist(productTemp As String, productTempNew As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select *  from  tblCustomerDieQty where CustomerName='" + productTemp + "' and  CustomerPT ='" + productTempNew + "' and  flag='Y'"

slectResult = QueryStr(cmdStr)
JudGCDieNoExist = slectResult
End Function

Public Function JudgePtExist(productTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from  TBLSETPT where productName='" + productTemp + "' and flag='Y'"
slectResult = QueryStr(cmdStr)
JudgePtExist = slectResult
End Function



Public Function JudgeTrayExist(fieldvalueTemp As String, traytypeTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = " select * from  TBLSETTray where fieldvalue='" + fieldvalueTemp + "' and traytype='" + traytypeTemp + "' and flag='Y' "

slectResult = QueryStr(cmdStr)
JudgeTrayExist = slectResult
End Function



Public Function JudgetestNoExist2(productTemp As String, oldTestNo As String, productTempNew As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select *   from  tblTestNo2 where  productname='" + productTemp + "' and TESTEDITION='" + oldTestNo + "' and flag='Y' and productnamenew='" + productTempNew + "' "


slectResult = QueryStr(cmdStr)
JudgetestNoExist2 = slectResult
End Function



Public Function JudgeFlagStautsBC(id As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from  CustomerBCtbl where batchid='" + id + "' and flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeFlagStautsBC = slectResult
End Function



Public Function JudgeFlag37POHeader(id As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from customeroitbl_test a where a.customershortname='37' and a.custom_part_no='" + id + "' and a.flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeFlag37POHeader = slectResult
End Function


Public Function JudgeFlag68POHeader(potemp As String, lotIdTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from customeroitbl_test a where a.customershortname in ('68','70') and po_num='" + potemp + "'  and SOURCE_MTRL_SLOC='" + lotIdTemp + "' and a.flag='Y'"




slectResult = QueryStr(cmdStr)
JudgeFlag68POHeader = slectResult
End Function


Public Function JudgeFlagStautsBCQty(id As String, qtyTemp As Long) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from  CustomerBItbl where batch_id='" & id & "' and die_qty=" & qtyTemp & "  and flag='Y'"
slectResult = QueryStr(cmdStr)
JudgeFlagStautsBCQty = slectResult
End Function



Public Function JudgeFlagVTData(devid As String, LOTID As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from  TSV_VT_History where deliveryno='" + devid + "' and CUSTLOT='" + LOTID + "' and flag='Y' "

slectResult = QueryStr(cmdStr)
JudgeFlagVTData = slectResult
End Function

Public Function JudgeShipSideData(GULFLotID As String) As Boolean
Dim cmdSql As String
Dim Resilt As Boolean

Resilt = False
cmdSql = "select * from erpdata.. tblSale_Shipto where LotID = '" & GULFLotID & "'"
Resilt = QuerySqlserver(cmdSql)
JudgeShipSideData = Resilt
End Function




'2012-12-04 jiayunzhang  添加 wo_no
Public Function JudgeGCHeaderId(id As String, woNoTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
'cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and flag='Y'"


cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and mtrl_num='" + woNoTemp + "'  "



slectResult = QueryStr(cmdStr)
JudgeGCHeaderId = slectResult
End Function



'37 Fab
Public Function Judge37FabData(LotType As String, wafer_type As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select * from mappingdatatest a  where a.lotid = '" + LotType + "'  and  a.wafer_id ='" + wafer_type + "' "


slectResult = QueryStr(cmdStr)
Judge37FabData = slectResult
End Function

Public Function Judge37FabDataICI(bagNoTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select * from  MAPPINGDATA37 a where a.flag='Y' and  a.Maptype='ICI' and Bagno='" + bagNoTemp + "' "


slectResult = QueryStr(cmdStr)
Judge37FabDataICI = slectResult
End Function



'37 Fab
Public Function Judge37FabDieFlag(ptTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

'cmdStr = "select * from  MAPPINGDATA37 a where a.purchaseno='" + poNoTemp + "' and a.batch='" + lotIDTemp + "' and a.flag='Y' "


'cmdStr = " select *  from  TBLTsvNpiProduct a where a.customershortname='37' and a.flag='Y' and (  a.customerptno1='" + ptTemp + "' or   a.customerptno2= '" + ptTemp + "' " & _
'" or a.customerptno3= '" + ptTemp + "' or   a.customerptno4= '" + ptTemp + "') "
'
 cmdStr = "  select *   from  TBLTsvNpiProduct a , SemtechFabPT b  where a.customershortname='37' and a.flag='Y' and b.flag='Y' and b.fabpt='" + ptTemp + "' " & _
" and (  a.customerptno1=b.npiproductpt or   a.customerptno2= b.npiproductpt or a.customerptno3= b.npiproductpt or   a.customerptno4=b.npiproductpt ) "
 


slectResult = QueryStr(cmdStr)
Judge37FabDieFlag = slectResult
End Function



Public Function JudgeMPSBankPT(ptTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False


  ' cmdStr = "  select a.qtechptno from  TBLTsvNpiProduct a  where a.customershortname in ('68','70','HK006','BJ128') and a.customerptno1='" + ptTemp + "' "
   
   cmdStr = "  select a.qtechptno from  TBLTsvNpiProduct a  where a.customerptno1='" + ptTemp + "' "
    
slectResult = QueryStr(cmdStr)
JudgeMPSBankPT = slectResult
End Function




Public Function Judge68FabDieFlag(ptTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

'cmdStr = "select * from  MAPPINGDATA37 a where a.purchaseno='" + poNoTemp + "' and a.batch='" + lotIDTemp + "' and a.flag='Y' "


'cmdStr = " select *  from  TBLTsvNpiProduct a where a.customershortname='37' and a.flag='Y' and (  a.customerptno1='" + ptTemp + "' or   a.customerptno2= '" + ptTemp + "' " & _
'" or a.customerptno3= '" + ptTemp + "' or   a.customerptno4= '" + ptTemp + "') "
'
 cmdStr = "  select *   from  TBLTsvNpiProduct a  where a.customershortname in ('68','70') and a.flag='Y' " & _
" and (  a.customerptno1='" + ptTemp + "' or   a.customerptno2= '" + ptTemp + "' or a.customerptno3= '" + ptTemp + "' or   a.customerptno4='" + ptTemp + "' ) "
 


slectResult = QueryStr(cmdStr)
Judge68FabDieFlag = slectResult
End Function





Public Function JudgeEQHeaderId(id As String, woNoTemp As String, po_Temp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
'cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and flag='Y'"


cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and mtrl_num='" + woNoTemp + "' and po_num='" + po_Temp + "'  "



slectResult = QueryStr(cmdStr)
JudgeEQHeaderId = slectResult
End Function

Public Function JudgeMCHeaderId(id As String, woNoTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
'cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and flag='Y'"


cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "'  "



slectResult = QueryStr(cmdStr)
JudgeMCHeaderId = slectResult
End Function


''2012-12-04 jiayunzhang  添加 wo_no
'Public Function JudgeSXHeaderId(id As String) As Boolean
'Dim cmdStr As String
'Dim slectResult As Boolean
'slectResult = False
''cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and flag='Y'"
'
'cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "'  "
'
'slectResult = QueryStr(cmdStr)
'JudgeSXHeaderId = slectResult
'End Function

'2014-03-20 jiayun Modify

'2016-06-13 jiayun modify add pt
Public Function JudgeSXHeaderId(id As String, potemp As String, ptTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
'cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and flag='Y'"

cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and po_num='" + potemp + "'  and flag='Y'  and mpn_desc='" + ptTemp + "' "

slectResult = QueryStr(cmdStr)
JudgeSXHeaderId = slectResult
End Function

Public Function JudgeRDA(waferid As String, cus As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select *  from RDABACKDATA where waferid='" + waferid + "' and customer='" + cus + "' "

slectResult = QueryStr(cmdStr)
JudgeRDA = slectResult

End Function


'2015-12-16 jiayun Modify 添加客户机种

Public Function JudgePOHeaderIdNew(id As String, potemp As String, custPTTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
'cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and flag='Y'"

cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and po_num='" + potemp + "'  and mpn_desc='" + custPTTemp + "'  and flag='Y' "

slectResult = QueryStr(cmdStr)
JudgePOHeaderIdNew = slectResult
End Function




Public Function JudgeQR2HeaderId(id As String, potemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
'cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and flag='Y'"

cmdStr = "select *  from customeroitbl_QR where source_batch_id='" + id + "' and po_num='" + potemp + "'  and flag='Y' "

slectResult = QueryStr(cmdStr)
JudgeQR2HeaderId = slectResult
End Function


Public Function JudgeEQISHeaderId(potemp As String, woTemp As String, LOTID As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
'cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and flag='Y'"

cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + LOTID + "' and po_num='" + potemp + "'   and flag='Y' "

slectResult = QueryStr(cmdStr)
JudgeEQISHeaderId = slectResult
End Function
Public Function JudgeEQISShippingRequest(soTemp As String, lineTemp As String, ScheduleLineTemp As String, customerPoTemp As String, DATECODETemp1 As String, DATECODE11 As String, LOT1 As String, DEVICETemp1 As String, SUBCONPO1 As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
', eqISHeaderTemp.DATECODETemp, eqISHeaderTemp.DATECODE1, eqISHeaderTemp.LOT, eqISHeaderTemp.DEVICETemp, eqISHeaderTemp.SUBCONPO


cmdStr = "select *  from EQ_SHIPPING_REQUEST where SO='" + soTemp + "' and LINE='" + lineTemp + "'   and SCHEDULELINE='" + ScheduleLineTemp + "' and CUSTOMERPO='" + customerPoTemp + "'  " & _
         " and DATECODE='" + DATECODETemp1 + "' and DATECODE1='" + DATECODE11 + "' and  LOT='" + LOT1 + "' and DEVICE='" + DEVICETemp1 + "' and SUBCONPO='" + SUBCONPO1 + "'"

slectResult = QueryStr(cmdStr)
JudgeEQISShippingRequest = slectResult
End Function



'2013-03-04 jiayunzhang  添加 PT
Public Function JudgePTHeaderId(id As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
'cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "' and flag='Y'"

cmdStr = "select *  from CustomerOItbl_test where source_batch_id='" + id + "'  "

slectResult = QueryStr(cmdStr)
JudgePTHeaderId = slectResult
End Function

Public Function JudgeGCDetailId(id As String, ITEM As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from mappingDataTest where lotid='" + id + "' and Wafer_ID='" + ITEM + "' and flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeGCDetailId = slectResult
End Function

Public Function JudgeBJNO(bj As String) As Boolean

Dim sql As String

sql = "select * from erptemp..tblBB_QUOTATION where QUOTATION = '" & bj & "'"

JudgeBJNO = QuerySqlserver(sql)

End Function

Public Function JudgeDept(dept As String) As Boolean
Dim strora As String
Dim deptTmp As String

depTmp = Mid(dept, InStr(1, dept, "_") + 1)

strora = "select FNumber  from AIS20141114094336.dbo.t_Department    where  FNumber='" + depTmp + "'"

JudgeDept = QuerySqlserver(strora)

End Function


Public Function JudgeRepeId(CUSTOMER As String) As Boolean
Dim Result As Boolean
Result = False

If (InStr(CUSTOMER, "(2)")) Then
    Result = False
Else
    Result = True
End If
JudgeRepeId = Result
End Function

Public Function JudgeRepeCount(LOTID As String, waferid As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select * from mappingDataTest where Wafer_ID = '" + waferid + "' And lotid = '" + LOTID + "' "

slectResult = QueryStr2(cmdStr)
JudgeRepeCount = slectResult
End Function

Public Function JudgeGCLableWlaID(waferIdTemp As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from TSV_GCLable_SETWLA where  Waferid='" + waferIdTemp + "' and flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeGCLableWlaID = slectResult
End Function




Public Function JudgeQR2DetailId(id As String, ITEM As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from mappingData_QR where lotid='" + id + "' and Wafer_ID='" + ITEM + "' and flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeQR2DetailId = slectResult
End Function



Public Function JudgeGCDetailIdWLD(id As String, ITEM As String) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from mappingDataTest where lotid='" + id + "' and substrateid='" + ITEM + "' and flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeGCDetailIdWLD = slectResult
End Function

Public Function QueryStr(cmdStr As String) As Boolean
    Dim resut As Boolean
    resut = False
    If Cnn.State = 0 Then
        ConOracle
    End If

    If rs.State = 1 Then
        rs.Close
    End If
    
    rs.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.RecordCount > 0 Then
        resut = True
    End If
    QueryStr = resut
End Function

Public Function QuerySqlserver(cmdStr As String) As Boolean
    Dim Result As Boolean
    Dim rs1 As ADODB.Recordset
    Result = False
    
    Set rs1 = getStrSqlServer(cmdStr)
    
    If rs1.RecordCount > 0 Then
        Result = True
    End If

    QuerySqlserver = Result
End Function

Public Function QueryStr2(cmdStr As String) As Boolean
    Dim resut As Boolean
    resut = False
    If Cnn.State = 0 Then
        ConOracle
    End If

    If rs.State = 1 Then
        rs.Close
    End If
    
    rs.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    

    
    If rs.RecordCount > 1 Then
        resut = True
    End If
    QueryStr2 = resut
End Function

Public Function GetOIData(customerTemp As String, lotIdTemp As String) As ADODB.Recordset

    Dim cmdStr   As String

    Dim RSResult As New ADODB.Recordset

    If customerTemp = "AA" Then
   
        cmdStr = "select po_num,mpn_desc,fabrication_facility,imager_customer_rev,design_id,shipping_mst_260,shipping_mst_level,encoded_mark_id,country_of_fab,micron_material,po_item," & " lot_status , mpn, protective_film_apld, lot_priority, ship_site ,PROBE_SHIP_PART_TYPE" & " from  CustomerOItbl_test where customershortname='" + customerTemp + "' and source_batch_id= '" + lotIdTemp + "'and flag='Y'"

    ElseIf customerTemp = "GC" Then

        cmdStr = "select  mtrl_num po_num, mpn_desc,fabrication_facility,'' as imager_customer_rev,design_id,shipping_mst_260,shipping_mst_level,encoded_mark_id, country_of_fab, micron_material,po_item," & " lot_status , mpn, protective_film_apld, lot_priority, '' as ship_site ,PROBE_SHIP_PART_TYPE" & " from  CustomerOItbl_test where customershortname='" + customerTemp + "' and source_batch_id= '" + lotIdTemp + "'and flag='Y' order by id desc"

    ElseIf customerTemp = "SX" Then

        cmdStr = "select  mtrl_num po_num,mpn_desc ,fabrication_facility,'' as imager_customer_rev,design_id,shipping_mst_260,shipping_mst_level,encoded_mark_id, country_of_fab, micron_material,po_item," & " lot_status , mpn, protective_film_apld, lot_priority, '' as ship_site,PROBE_SHIP_PART_TYPE " & " from  CustomerOItbl_test where customershortname='" + customerTemp + "' and source_batch_id= '" + lotIdTemp + "'and flag='Y' order by id desc"

    ElseIf customerTemp = "BD" Then

        cmdStr = "select  mtrl_num po_num, mpn_desc,fabrication_facility,'' as imager_customer_rev,design_id,shipping_mst_260,shipping_mst_level,encoded_mark_id, country_of_fab, micron_material,po_item," & " lot_status , mpn, protective_film_apld, lot_priority, '' as ship_site,PROBE_SHIP_PART_TYPE " & " from  CustomerOItbl_test where customershortname='" + customerTemp + "' and source_batch_id= '" + lotIdTemp + "'and flag='Y'"

    ElseIf customerTemp = "HY" Then

        cmdStr = "select  mtrl_num po_num,mpn_desc ,fabrication_facility,'' as imager_customer_rev,design_id,shipping_mst_260,shipping_mst_level,encoded_mark_id, country_of_fab, micron_material,po_item," & " lot_status , mpn, protective_film_apld, lot_priority, '' as ship_site ,PROBE_SHIP_PART_TYPE" & " from  CustomerOItbl_test where customershortname='" + customerTemp + "' and source_batch_id= '" + lotIdTemp + "'and flag='Y'"

    Else

        cmdStr = "select  mtrl_num po_num,mpn_desc ,fabrication_facility,'' as imager_customer_rev,design_id,shipping_mst_260,shipping_mst_level,encoded_mark_id, country_of_fab, micron_material,po_item," & " lot_status , mpn, protective_film_apld, lot_priority, '' as ship_site,PROBE_SHIP_PART_TYPE " & " from  CustomerOItbl_test where customershortname='" + customerTemp + "' and source_batch_id= '" + lotIdTemp + "'and (flag='Y' or flag = 'P')"

    End If
        
    Set RSResult = getStr(cmdStr)

    Set GetOIData = RSResult

End Function

Public Function GetOIData2(customerTemp As String, lotIdTemp As String) As ADODB.Recordset
Dim cmdStr As String

cmdStr = "select mpn_desc" & _
" from  CustomerOItbl_test where customershortname='" + customerTemp + "' and source_batch_id= '" + lotIdTemp + "'and flag='Y' order by id desc"

Set RSResult = getStr(cmdStr)


Set GetOIData2 = RSResult
End Function





Public Function GetGCPT_C(customerTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select CUSTOMERPTNew  from  TBLCUSTOMERPTAddC  where CUSTOMERPT='" + customerTemp + "' and flag='Y' "


Set RSResult = getStr(cmdStr)

Set GetGCPT_C = RSResult
End Function


Public Function GetOIDataPONum(customerTemp As String, poidTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

If customerTemp = "AA" Or customerTemp = "GT" Or customerTemp = "56" Then

cmdStr = " select qtech_created_date,mpn_desc,sum(qty) qty,count(substrateid) qty2 from ( " & _
" select distinct a.qtech_created_date,a.mpn_desc,c.lotid,c.substrateid,c.passbincount+c.failbincount as qty from customeroitbl_test a ,mappingdatatest c where a.po_num='" + poidTemp + "'  and a.customershortname='" + customerTemp + "' and a.flag='Y' and a.id=( " & _
" select max(b.id) from customeroitbl_test b where b.po_num='" + poidTemp + "' and b.customershortname='" + customerTemp + "' and b.flag='Y'" & _
" ) and c.customershortname='" + customerTemp + "' and c.lotid=a.source_batch_id ) X group by qtech_created_date,mpn_desc "


Else

cmdStr = "select to_char( aa.qtech_created_date,'YYYY-MM-DD') qtech_created_date    ," & _
"aa.MPN_DESC  mpn_desc , Sum(bb.PASSBINCOUNT + bb.FailBinCount) qty, Count(bb.SUBSTRATEID)  qty2" & _
" from customeroitbl_test aa ,mappingdatatest bb " & _
" Where bb.filename = aa.id And bb.LOTID = aa.source_batch_id " & _
" and bb.flag = 'Y' and aa.flag = 'Y' and aa.po_num  = '" & poidTemp & "' and aa.customershortname = '" & customerTemp & "' " & _
" GROUP BY to_char( aa.qtech_created_date,'YYYY-MM-DD'),AA.MPN_DESC"
End If

Set RSResult = getStr(cmdStr)

Set GetOIDataPONum = RSResult
End Function




Public Function GetBaoFeiOIData(lotIdTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
Dim customTemp As String


'cmdStr = "select a.lotid,a.passbincount,a.failbincount,b.design_id ,p.alternatename,d.product  from mappingdatatest a ,customeroitbl_test b ,ib_waferlist c ,ib_wohistory d , product p, productbase pb" & _
'" where a.customershortname in ( 'AA','HD')  and b.source_batch_id=a.lotid and c.waferid=a.substrateid " & _
'" and d.ordername=c.ordername  and a.substrateid= '" + lotIDtemp + "' " & _
'" and p.productbaseid = pb.productbaseid and d.product=pb.productname "
'


'查Mapping对应的客户  '如果是AA的，则串BC表

customTemp = GetWaferidCus(lotIdTemp)


If customTemp = "AA" Then


cmdStr = " select a.lotid,a.passbincount,a.failbincount, substr(mtrlnum,1,4)  as design_id ,p.alternatename,d.product " & _
" from mappingdatatest a ,  CustomerBCtbl b ,ib_waferlist c ,ib_wohistory d , product p, productbase pb " & _
" Where b.BatchID = a.lotid And c.waferid = a.SubstrateId" & _
" and d.ordername=c.ordername  and a.substrateid= '" + lotIdTemp + "' " & _
" and p.productbaseid = pb.productbaseid and d.product=pb.productname "


Else

 
cmdStr = "select a.lotid,a.passbincount,a.failbincount,CASE b.customershortname   WHEN 'AA' THEN   b.design_id  WHEN  'HD' THEN b.design_id   ELSE b.mpn_desc END   as design_id ,p.alternatename,d.product  from mappingdatatest a ,customeroitbl_test b ,ib_waferlist c ,ib_wohistory d , product p, productbase pb" & _
" where  b.source_batch_id=a.lotid and c.waferid=a.substrateid " & _
" and d.ordername=c.ordername  and a.substrateid= '" + lotIdTemp + "' " & _
" and p.productbaseid = pb.productbaseid and d.product=pb.productname "
 
End If

        
Set RSResult = getStr(cmdStr)

Set GetBaoFeiOIData = RSResult
End Function


Public Function GetAAGROIData(lotIdTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select po_num,po_item,source_mtrl_num,mpn,mpn_desc,mtrl_num_mtrlgrp,mtrl_num,created_date,created_time,SPGRQbox_SEQ.Nextval ID from customeroitbl_test where source_batch_id='" + lotIdTemp + "'and flag='Y' "

Set RSResult = getStr(cmdStr)

Set GetAAGROIData = RSResult
End Function



Public Function GetWoCreatedDate(lotIdTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


cmdStr = " select erpcreationdate  from mfgorder mfg where mfg.mfgordername='" + lotIdTemp + "'"
Set RSResult = getStr(cmdStr)

Set GetWoCreatedDate = RSResult
End Function

Public Function GetWoCreatedDate2(lotIdTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


cmdStr = " select max(txntimestamp) as txntimestamp  from A_Lotwafers al where al.workordername='" + lotIdTemp + "'"

 
Set RSResult = getStr(cmdStr)

Set GetWoCreatedDate2 = RSResult
End Function


'2013-04-25 jiayun add 查询OI表，供样品，重工用
Public Function GetOI2Data(customerTemp As String, lotIdTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

If customerTemp = "AA" Then
   
cmdStr = "select po_num,mpn_desc,fabrication_facility,imager_customer_rev,design_id,shipping_mst_260,shipping_mst_level,encoded_mark_id,country_of_fab,micron_material,po_item," & _
" lot_status , mpn, protective_film_apld, lot_priority, ship_site ,PROBE_SHIP_PART_TYPE " & _
" from  CustomerOItbl_test where customershortname='" + customerTemp + "' and source_batch_id= '" + lotIdTemp + "'and flag='Y'"

ElseIf customerTemp = "GC" Then

cmdStr = "select  mtrl_num po_num,substr(mpn_desc,1,length(mpn_desc)-2) as mpn_desc,fabrication_facility,'' as imager_customer_rev,design_id,shipping_mst_260,shipping_mst_level,encoded_mark_id, country_of_fab, micron_material,po_item," & _
" lot_status , mpn, protective_film_apld, lot_priority, '' as ship_site ,PROBE_SHIP_PART_TYPE " & _
" from  CustomerOItbl_test where customershortname='" + customerTemp + "' and source_batch_id= '" + lotIdTemp + "'and flag='Y' order by qtech_created_date desc  "

ElseIf customerTemp = "SX" Then

cmdStr = "select  mtrl_num po_num,mpn_desc ,fabrication_facility,'' as imager_customer_rev,design_id,shipping_mst_260,shipping_mst_level,encoded_mark_id, country_of_fab, micron_material,po_item," & _
" lot_status , mpn, protective_film_apld, lot_priority, '' as ship_site ,PROBE_SHIP_PART_TYPE " & _
" from  CustomerOItbl_test where customershortname='" + customerTemp + "' and source_batch_id= '" + lotIdTemp + "'and flag='Y' order by id desc"

ElseIf customerTemp = "MC" Then

cmdStr = "select  mtrl_num po_num,mpn_desc ,fabrication_facility,'' as imager_customer_rev,design_id,shipping_mst_260,shipping_mst_level,encoded_mark_id, country_of_fab, micron_material,po_item," & _
" lot_status , mpn, protective_film_apld, lot_priority, '' as ship_site ,PROBE_SHIP_PART_TYPE " & _
" from  CustomerOItbl_test where customershortname='" + customerTemp + "' and source_batch_id= '" + lotIdTemp + "'and flag='Y'"


Else

cmdStr = "select po_num,mpn_desc,fabrication_facility,imager_customer_rev,design_id,shipping_mst_260,shipping_mst_level,encoded_mark_id,country_of_fab,micron_material,po_item," & _
" lot_status , mpn, protective_film_apld, lot_priority, ship_site ,PROBE_SHIP_PART_TYPE " & _
" from  CustomerOItbl_test where  source_batch_id= '" + lotIdTemp + "' "
 
End If
        
Set RSResult = getStr(cmdStr)

Set GetOI2Data = RSResult
End Function



Public Function GetOINewProduct() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select po_num,source_batch_id from  CustomerOItbl_test where customershortname='AA' and  flag='Y' and qtech_created_date >sysdate-2 order by   po_num,source_batch_id "

        
Set RSResult = getStr(cmdStr)

Set GetOINewProduct = RSResult
End Function

Public Function GetGT5271() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = "select po_num,source_batch_id from  CustomerOItbl_test where customershortname='AA' and  flag='Y' and qtech_created_date >sysdate-2 order by   po_num,source_batch_id "

 cmdStr = " select containername,moveinqty  ,1 from     a_wiplothistory where productname in ('18SIC1A000CF','18E001M0000CF') and specname='5270'  and  containername like '%-A%' order by containername"

        
Set RSResult = getStr(cmdStr)

Set GetGT5271 = RSResult
End Function

Public Function GetGT5271Where(LOTID As String) As ADODB.Recordset

Dim cmdStr As String
Dim cmdStr2 As String
Dim RSResult As New ADODB.Recordset
Dim RSResult1 As New ADODB.Recordset


'cmdStr = "select po_num,source_batch_id from  CustomerOItbl_test where customershortname='AA' and  flag='Y' and qtech_created_date >sysdate-2 order by   po_num,source_batch_id "
 
 
 'Set RSResult = getSqlServerStr2(cmdStr2)
 
 
 cmdStr = " select containername,moveinqty  ,1 from     a_wiplothistory where productname in ('18SIC1A000CF','18E001M0000CF') and specname='5270'  and  containername like '%-A%' and  containername like '" + LOTID + "%'  order by containername"
'cmdStr = " select a.containername, a.moveinqty, 1 from a_wiplothistory a where a.productname in ('18SIC1A000CF', '18E001M0000CF') and a.specname = '5270' and a.containername like '%-A%' and a.containername like '" + lotid + "%' and a.containername not in  (select b.waferscribenumber||'-A' from tsv_qboxnumber_waibao_36 b where b.waferscribenumber like '" + lotid + "%' ) order by a.containername"

   
   
   
  

        
Set RSResult = getStr(cmdStr)

Set GetGT5271Where = RSResult
End Function



Public Function GetMergAAQueryInf() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select lotID,mpn,mtrl_num,test_program_rev,dateCode from TSV_AA_MergeQuery order by createdate desc "

        
Set RSResult = getStr(cmdStr)

Set GetMergAAQueryInf = RSResult
End Function




Public Function GetWipSetData(lotIdTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select lotdate,remark from  WipreportDate where lotid='" + lotIdTemp + "'"
    
Set RSResult = getStr(cmdStr)

Set GetWipSetData = RSResult
End Function





Public Function GetLotDetailData(customerTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
   
 If customerTemp = "AA" Then
    cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and  die_qty-downqty>0  order by source_batch_id "
 
 ElseIf customerTemp = "GC" Then
 
 cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'  and downqty is null order by source_batch_id "
 

 
 '2013-04-26 jiayun add
 
 ElseIf customerTemp = "PT" Or customerTemp = "DN" Then
 
 cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'  and downqty is null order by source_batch_id "
 
 Else
 
 ' cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'  and downqty is null order by source_batch_id "
 
   cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'   order by source_batch_id "
 
 End If
 
 
        
Set RSResult = getStr(cmdStr)

Set GetLotDetailData = RSResult
End Function

Public Function GetLotDetailDataNew(customerTemp As String, customerPTTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
   
 If customerTemp = "AA" Then
    cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and  die_qty-downqty>0  order by source_batch_id "
 
 ElseIf customerTemp = "GC" Then
 
 cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'  and downqty is null and mpn_desc like '" + customerPTTemp + "%' order by source_batch_id "
 

 
 '2013-04-26 jiayun add
 
 ElseIf customerTemp = "PT" Or customerTemp = "DN" Then
 
  
 ElseIf customerTemp = "HY" Then
 
'  cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0' order by source_batch_id "
'
cmdStr = " select distinct a.source_batch_id from  CustomerOItbl_test a,mappingDataTest b  where a.customershortname='" + customerTemp + "' and a.flag='Y' and a.invflag='0' " & _
            " and b.lotid=a.source_batch_id and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid) order by a.source_batch_id"
 
 
 Else
 
 ' cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'  and downqty is null order by source_batch_id "
 
   cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0' and mpn_desc like '" + customerPTTemp + "%'  order by source_batch_id "
 
 End If
 
 
        
Set RSResult = getStr(cmdStr)

Set GetLotDetailDataNew = RSResult
End Function



Public Function GetLotDetailDataForSo(customerTemp As String, soTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
'2014-03-17 jiayun modify sql

'
'cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and  die_qty-downqty>0  and po_num='" + soTemp + "' order by source_batch_id "
'
' If customerTemp = "GC" Or customerTemp = "BD" Then
'
' cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'  and downqty is null order by source_batch_id "
'
' End If
'
' '2013-10-10 jiayun add
'
'If customerTemp = "SX" Then
'
'' cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'  and downqty is null order by source_batch_id "
''
'
'cmdStr = " select distinct  a.source_batch_id from CustomerOItbl_test a ,mappingDataTest b " & _
'         " where a.customershortname = 'SX' and a.flag = 'Y' and a.invflag = '0' and a.downqty is null" & _
'         " and a.source_batch_id=b.lotid and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid)  order by a.source_batch_id "
'
'
' End If
'
' If customerTemp = "HY" Then
'
'' cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'  and downqty is null order by source_batch_id "
''
'
'cmdStr = " select distinct  a.source_batch_id from CustomerOItbl_test a ,mappingDataTest b " & _
'         " where a.customershortname = 'HY' and a.flag = 'Y' and a.invflag = '0' and a.downqty is null" & _
'         " and a.source_batch_id=b.lotid and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid)  order by a.source_batch_id "
'
'Else
'
'cmdStr = " select distinct  a.source_batch_id from CustomerOItbl_test a ,mappingDataTest b " & _
'         " where a.customershortname = '" + customerTemp + "' and a.flag = 'Y' and a.invflag = '0' and a.downqty is null" & _
'         " and a.source_batch_id=b.lotid and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid)  order by a.source_batch_id "
'
'
' End If
 
 
 If customerTemp = "AA" Then
 
   
cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and  die_qty-downqty>0  and po_num='" + soTemp + "' order by source_batch_id "
 
 ElseIf customerTemp = "GC" Then
 
 cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'  and downqty is null order by source_batch_id "
 
ElseIf customerTemp = "BD" Then

 cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'  and downqty is null order by source_batch_id "
 

ElseIf customerTemp = "SX" Then
 
 
cmdStr = " select distinct  a.source_batch_id from CustomerOItbl_test a ,mappingDataTest b " & _
         " where a.customershortname = 'SX' and a.flag = 'Y' and a.invflag = '0' and a.downqty is null" & _
         " and a.source_batch_id=b.lotid and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid)  order by a.source_batch_id "
 
 
ElseIf customerTemp = "HY" Then
 
 
 
cmdStr = " select distinct  a.source_batch_id from CustomerOItbl_test a ,mappingDataTest b " & _
         " where a.customershortname = 'HY' and a.flag = 'Y' and a.invflag = '0' and a.downqty is null" & _
         " and a.source_batch_id=b.lotid and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid)  order by a.source_batch_id "
 
Else

cmdStr = " select distinct  a.source_batch_id from CustomerOItbl_test a ,mappingDataTest b " & _
         " where a.customershortname = '" + customerTemp + "' and a.flag = 'Y' and a.invflag = '0' and a.downqty is null" & _
         " and a.source_batch_id=b.lotid and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid)  order by a.source_batch_id "

 
 End If
 
 

Set RSResult = getStr(cmdStr)

Set GetLotDetailDataForSo = RSResult
End Function


Public Function GetLotDetailDataForSoNew(customerTemp As String, soTemp As String, customerPTTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

 
If customerTemp = "AA" Then
 
cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and  die_qty-downqty>0  and po_num='" + soTemp + "' order by source_batch_id "
 
 
 ElseIf customerTemp = "GC" Then
 
 cmdStr = "select source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'   and mpn_desc like '" + customerPTTemp + "%' order by source_batch_id "
 
 cmdStr = " select distinct  a.source_batch_id from CustomerOItbl_test a ,mappingDataTest b " & _
         " where a.customershortname = '" + customerTemp + "' and (a.flag = 'Y' or a.flag = 'P') and a.invflag = '0' " & _
         " and a.source_batch_id=b.lotid and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid)  order by a.source_batch_id "
ElseIf customerTemp = "BD" Then

 cmdStr = "select distinct source_batch_id from  CustomerOItbl_test where customershortname='" + customerTemp + "' and flag='Y' and invflag='0'   order by source_batch_id "
 

ElseIf customerTemp = "SX" Then
 
 
cmdStr = " select distinct  a.source_batch_id from CustomerOItbl_test a ,mappingDataTest b " & _
         " where a.customershortname = 'SX' and a.flag = 'Y' and a.invflag = '0'  and mpn_desc like '" + customerPTTemp + "%' " & _
         " and a.source_batch_id=b.lotid and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid)  order by a.source_batch_id "
 
 
ElseIf customerTemp = "HY" Then
 
 
 
cmdStr = " select distinct  a.source_batch_id from CustomerOItbl_test a ,mappingDataTest b " & _
         " where a.customershortname = 'HY' and a.flag = 'Y' and a.invflag = '0' " & _
         " and a.source_batch_id=b.lotid and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid)  order by a.source_batch_id "
 

ElseIf customerTemp = "37(ICI)" Then



cmdStr = " select distinct  a.source_batch_id from CustomerOItbl_test a ,mappingDataTest b " & _
         " where a.customershortname in ('37','37(ICI)') and a.flag = 'Y' and a.invflag = '0' " & _
         "    and a.mtrl_num=b.substrateid  and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid)  order by a.source_batch_id "

 
 

Else

cmdStr = " select distinct  a.source_batch_id from CustomerOItbl_test a ,mappingDataTest b " & _
         " where a.customershortname = '" + customerTemp + "' and (a.flag = 'Y' or a.flag = 'P') and a.invflag = '0' " & _
         " and a.source_batch_id=b.lotid and not exists (select 1 from A_Lotwafers al where al.waferscribenumber=b.substrateid)  order by a.source_batch_id "

 
 End If
 
 

Set RSResult = getStr(cmdStr)

Set GetLotDetailDataForSoNew = RSResult
End Function


Public Function GetLotDetailDataForSoNewOn(customerPTTemp As String, opnTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


 
cmdStr = " select  distinct  b.id, b.batchid as source_batch_id from  ( select * from (select *  from CUSTOMERFORECASTTBL   order by ID desc) where   out_part_id = '" + customerPTTemp + "'  and rownum = 1 ) a ,CustomerBCtbl b " & _
"  where a.out_part_id='" + customerPTTemp + "' and a.comments='" + opnTemp + "' and a.flag='Y' and a.start_part_id=b.mtrlnum and b.batchid not in (select lotid from  On_WO_HisTory where flag='Y')   order by b.id "
       


 

Set RSResult = getStr(cmdStr)

Set GetLotDetailDataForSoNewOn = RSResult
End Function


Public Function GetONDateCode() As String

Dim tt As Date
Dim ttold As Date
Dim i As Integer
Dim dateCodeTemp As Integer

tt = CDate(Now)

If Year(Now) = "2015" Then

ttold = CDate("2015-01-02")

ElseIf Year(Now) = "2016" Then
ttold = CDate("2016-01-01")

ElseIf Year(Now) = "2017" Then
ttold = CDate("2017-01-01")
ElseIf Year(Now) = "2018" Then
ttold = CDate("2018-01-01")
ElseIf Year(Now) = "2019" Then
ttold = CDate("2019-01-01")
ElseIf Year(Now) = "2020" Then
ttold = CDate("2020-01-01")
ElseIf Year(Now) = "2021" Then
ttold = CDate("2021-01-01")
ElseIf Year(Now) = "2022" Then
ttold = CDate("2022-01-01")

End If

i = DateDiff("d", ttold, tt)

Dim j As Integer
Dim k As Integer


j = i \ 7
k = i Mod 7

If k = 0 Then
dateCodeTemp = j

Else

dateCodeTemp = j + 1

End If

If Format(Now, "YYYY-MM-DD") = "2016-01-01" Then
   GetONDateCode = "1601"
   
ElseIf Format(Now, "YYYY-MM-DD") = "2016-12-31" Then
   GetONDateCode = "1652"
 
Else
 
 GetONDateCode = Right(Year(Now), 2) & Right("0" & dateCodeTemp, 2)
 
End If
 

End Function


Public Function GetONOPN(ptTemp As String) As String

Dim cmdStr As String
Dim RSResult  As String

cmdStr = "  select comments as newonopnpart from CUSTOMERFORECASTTBL  where out_part_id='" + ptTemp + "' and typename='FEDS' and flag='Y' and rownum<2"

RSResult = getStr2(cmdStr)

 GetONOPN = RSResult
End Function


Public Function GetONOPN_WSG(ptTemp As String) As String

Dim cmdStr As String
Dim RSResult  As String

cmdStr = " select out_part_id as newonopnpart from CUSTOMERFORECASTTBL where comments='" + ptTemp + "'  order by id desc "

RSResult = getStr2(cmdStr)

 GetONOPN_WSG = RSResult
End Function



Public Function GetLot(JOBID As String) As String

Dim cmdStr As String
Dim RSResult  As String


cmdStr = " select distinct ct.source_batch_id  from customeroitbl_test  ct where ct.test_mtrl_desc = '" + JOBID + "' and ct.source_batch_id  is not null"


RSResult = getStr2(cmdStr)

 GetLot = RSResult
End Function




Public Function GetONWoMarkingCode(ptTemp As String) As String

Dim cmdStr As String
Dim RSResult  As String

cmdStr = "  select MarkingCodeFirst from CUSTOMERMPNAttributes  where part='" + ptTemp + "'  and flag='Y' "

RSResult = getStr2(cmdStr)

 GetONWoMarkingCode = RSResult
End Function



Public Function GetONCS(lotIdTemp As String) As String

Dim cmdStr As String
Dim RSResult  As String

cmdStr = "   select designid  id from  CustomerBCtbl where  batchid='" + lotIdTemp + "' and flag='Y'"
       

 RSResult = getStr2(cmdStr)

 GetONCS = RSResult
End Function


Public Function GetONBCPlace(lotIdTemp As String) As String

Dim cmdStr As String
Dim RSResult  As String

cmdStr = "   select substr(mtrlnum,instr(mtrlnum,'-')+1,2)  id from  CustomerBCtbl where  batchid='" + lotIdTemp + "' and flag='Y'"
       

 RSResult = getStr2(cmdStr)

 GetONBCPlace = RSResult
End Function



Public Function JudgeCustomerPTNum(lotIdTemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = " select source_batch_id,count( distinct mpn_desc) from   CustomerOItbl_test " & _
         " where  source_batch_id='" + lotIdTemp + "' and customershortname<>'AA' and customershortname is not null " & _
         " group by source_batch_id having count( distinct mpn_desc)>1"


slectResult = QueryStr(cmdStr)
JudgeCustomerPTNum = slectResult
End Function


Public Function JudgeMofidyQboxStatus(containerTemp As String, qboxtemp As String) As Boolean
'耞琌
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False


cmdStr = " select wafernumber from tsv_qboxnumber_details b " & _
         " where   b.containername='" + containerTemp + "'  and  b.qboxnumber='" + qboxtemp + "' and customername='36' "


slectResult = QueryStr(cmdStr)
JudgeMofidyQboxStatus = slectResult
End Function

Public Function getStr(cmdStr As String) As ADODB.Recordset
    Dim resut As New ADODB.Recordset
   If Cnn.State = 0 Then
    ConOracle
    End If
    resut.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    Set getStr = resut
End Function

Public Function execSql(cmdStr As String) As ADODB.Recordset
    Dim resut As New ADODB.Recordset
   If Cnn.State = 0 Then
    ConOracle
    End If
    resut.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    Set execSql = resut
End Function



Public Function getSqlServerStr2(cmdStr As String) As ADODB.Recordset
    Dim resut As New ADODB.Recordset

    If INIadoCon.State = 0 Then
    INIConnectSTART
    End If


    resut.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Set getSqlServerStr2 = resut
End Function



Public Function getStrSqlServer(cmdStr As String) As ADODB.Recordset
    Dim resut As New ADODB.Recordset
    If INIadoCon.State = 0 Then
        INIConnectSTART
    End If

    resut.Open cmdStr, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Set getStrSqlServer = resut
End Function

Public Function GetFps(sqlTemp As String, customerTemp As String) As ADODB.Recordset

    Dim cmdStr   As String

    Dim RSResult As New ADODB.Recordset

    If customerTemp = "AA" Or customerTemp = "AA(ON)" Then
   
        cmdStr = " select id, substrateid,'',passbincount+failbincount,passbincount,lotid,ONMarkingCodeSeq.QTSeq(substrateid) ,'' " & "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname in ('AA','AA(ON)')  and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by substrateid  "

    ElseIf customerTemp = "CN" Then
   
        cmdStr = " select id, substrateid,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','') ,'' " & "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by substrateid  "

    ElseIf customerTemp = "GC" Or customerTemp = "SX" Or customerTemp = "BD" Then

        cmdStr = " select a.id, a.substrateid ,'',a.passbincount+a.failbincount,a.passbincount,a.lotid,replace(productid,'_','')  ,'' " & "  from   mappingDataTest a ,customeroitbl_test b   where a.lotid in ('" & sqlTemp & "' ) and  a.customershortname='" + customerTemp + "'  and a.flag='Y'  and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)" & " and b.source_batch_id=a.lotid and a.filename=b.id  order by 2  "

    ElseIf customerTemp = "HY" Then

        cmdStr = " select id, lotid || substr('0'||wafer_id,-2,2) ,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','')  ,'' " & "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y'  and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid) order by 2  "

    ElseIf customerTemp = "MC" Then

        cmdStr = " select id, substrateid ,'',passbincount,failbincount,lotid,replace(productid,'_','')  ,'' " & "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' )  and  customershortname='" + customerTemp + "' and flag='Y'  and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid) order by 2 "

    ElseIf customerTemp = "HD" Or customerTemp = "MG" Then
        cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','')  ,'' " & "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by 2  "

    ElseIf customerTemp = "GT" Or customerTemp = "SI" Then
        cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','')  ,'' " & "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by 2  "

    ElseIf customerTemp = "CS" Then
        cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','')  ,'' " & "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by 2  "

    ElseIf customerTemp = "37(ICI)" Then

        cmdStr = "select a.id, a.substrateid ,'',a.passbincount+a.failbincount,a.passbincount,b.source_batch_id,replace(productid,'_','')  ,'' " & " from   mappingDataTest a , CustomerOItbl_test b " & " where b.source_batch_id in ('" & sqlTemp & "' )  and a.substrateid=b.mtrl_num  and  a.customershortname='37' and a.flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by 2 "

    ElseIf customerTemp = "37" Then

        cmdStr = " select a.id, a.substrateid ,'',a.passbincount+a.failbincount,a.passbincount,a.lotid,replace(productid,'_','')  ,'' " & "  from   mappingDataTest a,customeroitbl_test c where  a.FileName = to_char(c.ID) and a.lotid in ('" & sqlTemp & "' )  and c.po_num is not null " & "  and  a.customershortname='" + customerTemp + "' and a.flag='Y'   and not exists (select 1 from a_lotwafers b where b.waferscribenumber = a.substrateid) order by 2  "

    Else

        cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','') ,'' " & "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and (flag='Y' or flag = 'P') and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by 2   "

    End If
        
    Set RSResult = getStr(cmdStr)

    Set GetFps = RSResult

End Function

Public Function GetFpsForSX(sqlTemp As String, customerTemp As String, ptTemp As String, htTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

If customerTemp = "AA" Or customerTemp = "AA(ON)" Then
   
cmdStr = " select id, substrateid,'',passbincount+failbincount,passbincount,lotid,ONMarkingCodeSeq.QTSeq(substrateid) ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname in ('AA','AA(ON)')  and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by substrateid  "

ElseIf customerTemp = "CN" Then
   
cmdStr = " select id, substrateid,'',passbincount+failbincount,passbincount,lotid,productid ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by substrateid  "

ElseIf customerTemp = "GC" Or customerTemp = "HJ" Or customerTemp = "BD" Then



cmdStr = " select a.id, a.lotid || substr('0'||a.wafer_id,-2,2) ,'',a.passbincount + a.failbincount,a.passbincount,a.lotid,replace(productid,'_','') ,'' " & _
         "  from   mappingDataTest a ,customeroitbl_test b   where a.lotid in ('" & sqlTemp & "' ) and  a.customershortname='" + customerTemp + "'  and b.mpn_desc='" + ptTemp + "'  and a.flag='Y'  and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)" & _
         " and b.source_batch_id=a.lotid and a.filename=b.id  order by 2  "

ElseIf customerTemp = "SX" Or customerTemp = "TJ003" Then

cmdStr = " select a.id, a.substrateid ,'',a.passbincount + a.failbincount,a.passbincount,a.lotid,replace(productid,'_','') ,'' " & _
         "  from   mappingDataTest a ,customeroitbl_test b   where a.lotid in ('" & sqlTemp & "' ) and  a.customershortname='" + customerTemp + "'  and b.mtrl_num='" + htTemp + "'  and a.flag='Y'  and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)" & _
         " and b.source_batch_id=a.lotid and a.filename=b.id  order by 2  "

ElseIf customerTemp = "HY" Then

cmdStr = " select id, lotid || substr('0'||wafer_id,-2,2) ,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','')  ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y'  and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid) order by 2  "


ElseIf customerTemp = "MC" Then

cmdStr = " select id, substrateid ,'',passbincount,failbincount,lotid,replace(productid,'_','')  ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' )  and  customershortname='" + customerTemp + "' and flag='Y'  and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid) order by 2 "

ElseIf customerTemp = "HD" Or customerTemp = "MG" Then
cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','')  ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by 2  "


ElseIf customerTemp = "GT" Or customerTemp = "SI" Then
cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','') ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by 2  "


ElseIf customerTemp = "CS" Then
cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','')  ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by 2  "


ElseIf customerTemp = "37(ICI)" Then


 cmdStr = "select a.id, a.substrateid ,'',a.passbincount+a.failbincount,a.passbincount,b.source_batch_id,replace(productid,'_','')  ,'' " & _
" from   mappingDataTest a , CustomerOItbl_test b " & _
" where b.source_batch_id in ('" & sqlTemp & "' )  and a.substrateid=b.mtrl_num  and  a.customershortname='37' and a.flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by 2 "

Else

cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,replace(productid,'_','')  ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by 2   "
End If
 
Set RSResult = getStr(cmdStr)

Set GetFpsForSX = RSResult
End Function




Public Function GetFps37WaferRec() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset



cmdStr = " select 'PO_RECEIPT' as Event, a.purchaseno,a.purchaseorderlineitem,a.devicename,a.wf,a.batch from MAPPINGDATA37 a  where a.status=0 order by a.qtech_created_date "

Set RSResult = getStr(cmdStr)

Set GetFps37WaferRec = RSResult
End Function


Public Function GetFps37POCommit() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


cmdStr = " select distinct 'COMMIT',d.po_num ,d.po_item,d.mpn,a.erpcreatedate,d.source_batch_id,d.mtrl_num from ib_wohistory a,  " & _
" ib_waferlist b ,mappingdatatest c ,customeroitbl_test d ,mfgorder e,container f where a.ordername = b.ordername and c.substrateid = b.waferid " & _
" and to_char(d.id) = c.filename and e.mfgordername = a.ordername and e.mfgorderid = f.mfgorderid and d.customershortname = '37' " & _
" and a.erpcreatedate > sysdate - 3 and d.po_num is not null order by a.erpcreatedate desc "


Set RSResult = getStr(cmdStr)

Set GetFps37POCommit = RSResult
End Function

Public Function GetFps37POStart() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


'cmdStr = " select 'COMMIT'as Event,po_num,po_item, mpn,' ' as commitdate from customeroitbl_test a where a.customershortname='37' and a.qtech_created_date>to_date('2016-03-26','YYYY-MM-DD') " & _
'" and a.lot_status is null order by a.id "


'cmdStr = " select distinct  'START' as wostart,c.po_num,c.po_item,c.mpn,to_char(a.erpcreatedate,'YYYYMMDD'),c.SOURCE_MTRL_SLOC  from ib_wohistory a ,ib_waferlist b ,customeroitbl_test c ,mappingdata37 d" & _
'" Where b.OrderName = a.OrderName and a.customer='37' and a.erpcreatedate>to_Date('2016-03-26','YYYY-MM-DD') and c.source_batch_id=b.waferlot " & _
'" and d.batch=c.source_batch_id and d.purchaseno=c.po_num and a.lot_status is null "


'cmdStr = " select distinct  'START' as wostart,c.po_num,c.po_item,c.mpn,to_char(a.erpcreatedate,'YYYYMMDD'),c.source_batch_id  from ib_wohistory a ,ib_waferlist b ,customeroitbl_test c " & _
'" Where b.OrderName = a.OrderName and a.customer='37' and a.erpcreatedate>to_Date('2016-03-26','YYYY-MM-DD') and c.source_batch_id=b.waferlot " & _
'" and a.erpcreatedate>=to_date(to_char(sysdate - 7,'YYYY-MM-DD'),'YYYY-MM-DD') and a.lot_status is null   order by c.po_num "
'

cmdStr = " select distinct 'START',d.po_num ,d.po_item,d.mpn,a.erpcreatedate,f.firstname from ib_wohistory a, ib_waferlist b ,mappingdatatest c ," & _
"customeroitbl_test d ,mfgorder e,container f where a.ordername = b.ordername and c.substrateid = b.waferid  and to_char(d.id) = c.filename" & _
" and e.mfgordername = a.ordername and e.mfgorderid = f.mfgorderid and d.customershortname = '37' and a.erpcreatedate > sysdate - 3" & _
" and d.po_num is not null order by a.erpcreatedate desc "
  
 
 

Set RSResult = getStr(cmdStr)

Set GetFps37POStart = RSResult
End Function

Public Function GetFpsAARTWo(sqlTemp As String, customerTemp As String, woTypeTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'2012-11-07 jiayun 修改sql

If (customerTemp = "AA" Or customerTemp = "AA(ON)") And (woTypeTemp = "ST" Or woTypeTemp = "ET") Then
   
'cmdStr = " select id, substrateid,'',passbincount+failbincount,passbincount,lotid,(select b.planned_laser_scribe from CustomerOItbl_test b where  b.source_batch_id= a.lotid) || Get_AA_MarkingCode(substrateid) ,'' " & _
'         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='AA' and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by substrateid  "


cmdStr = " select id, substrateid,'',passbincount+failbincount,passbincount,lotid,ONSTMarkingCodeSeq.QTSeq(substrateid,lotid) ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname in ('AA','AA(ON)')  and flag='Y' and not exists (select 1 from a_lotwafers b where b.waferscribenumber=a.substrateid)  order by substrateid  "

End If
 
        
Set RSResult = getStr(cmdStr)

Set GetFpsAARTWo = RSResult
End Function








Public Function GetProductBom(productTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

productTemp = Trim(Replace(Replace(productTemp, Chr(13), ""), Chr(10), ""))

cmdStr = " SELECT  b.料号 FROM [erpdata].[dbo].[TSVtblSetMRule] a, [erpdata].[dbo].[TSVtblMRuleData] b " & _
         " Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & productTemp & "'"
 
        
Set RSResult = getSqlServerStr2(cmdStr)

Set GetProductBom = RSResult
End Function


Public Function GetBondedTax(batchIdTemp As String, txtPO) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select ct.PO_NUM from customeroitbl_test ct where ct.source_batch_id='" & batchIdTemp & "' and ct.mtrl_num ='" & txtPO & "' "

Set RSResult = getStr(cmdStr)
Set GetBondedTax = RSResult
End Function
Public Function GetLastMI(waferIdTemp As String, lotIdTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select a.productid from mappingdatatest a where a.wafer_id='" & waferIdTemp & "' and a.lotid='" & lotIdTemp & "'"

Set RSResult = getStr(cmdStr)
Set GetLastMI = RSResult
End Function
Public Function GetIdCount(waferIdTemp As String, lotIdTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select a.* from mappingdatatest a where a.wafer_id='" & waferIdTemp & "' and a.lotid='" & lotIdTemp & "'"

Set RSResult = getStr(cmdStr)
Set GetIdCount = RSResult
End Function


Public Function Get37MergeLotDetails(containerTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


         
 cmdStr = " select b.wafernumber,c.po_num,b.ndpw from container a ,a_lotwafers b ,customeroitbl_test c " & _
" Where b.containerID = a.containerID and a.containername='" & containerTemp & "'" & _
" and instr(a.containername,b.waferscribenumber)<1 and c.source_batch_id=b.wafernumber "

 
        
Set RSResult = getStr(cmdStr)

Set Get37MergeLotDetails = RSResult
End Function




Public Function GetProductBomERpSign(productTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

productTemp = Trim(Replace(Replace(productTemp, Chr(13), ""), Chr(10), ""))
cmdStr = " SELECT  b.料号 FROM [erpdata].[dbo].[TSVtblSetMRule] a, [erpdata].[dbo].[TSVtblMRuleData] b " & _
         " Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & productTemp & "' and  a.审核日期 is not null  "
 
        
Set RSResult = getSqlServerStr2(cmdStr)

Set GetProductBomERpSign = RSResult
End Function
Public Function GetNpiCustmer(pt As String, ct As String, MPN As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = "select t.* from tbltsvnpiproduct t where t.qtechptno2  = '" & pt & "' and (t.customerptno1 = '" & ct & "' or  t.customerptno2 = '" & ct & "' or  t.customerptno3 = '" & ct & "' or  t.customerptno4 = '" & ct & "' or " & _
'"t.customerptno1 = '" & mpn & "' or  t.customerptno2 = '" & mpn & "' or  t.customerptno3 = '" & mpn & "' or  t.customerptno4 = '" & mpn & "' )"

cmdStr = "select t.* from tbltsvnpiproduct t where t.qtechptno2  = '" & pt & "' and (t.customerptno1 = '" & ct & "' or t.customerptno1 = '" & MPN & "' )"

Set RSResult = getStr(cmdStr)

Set GetNpiCustmer = RSResult
End Function

Public Function ResizeOrderQty(OrderID As String)
' 重新抛工单
Dim STr_Sql1 As String
Dim str_sql2 As String
Dim Str_sql3 As String
Dim STr_sql4 As String
Dim STr_sql5 As String

' 1. 删除工单数据
STr_Sql1 = "delete from A_Lotwafers al where al.workordername in ('" + OrderID + "')"
str_sql2 = "delete from container conn where conn.mfgorderid in  (select mfg.mfgorderid from mfgorder mfg where mfg.mfgordername  in ('" + OrderID + "')) "
Str_sql3 = "delete from mfgorder mfg where mfg.mfgordername in ('" + OrderID + "')"

getStr (STr_Sql1)
getStr (str_sql2)
getStr (Str_sql3)

MsgBox "已清除残留数据"
' 2. 重新插入数据
STr_sql4 = "insert into ib_workorder select * from ib_wohistory where ordername=('" + OrderID + "')"

STr_sql5 = "delete from ib_wohistory where ordername in ('" + OrderID + "')"

getStr (STr_sql4)
getStr (STr_sql5)
MsgBox "重新抛转数据开始, 请等待1分钟查看数据是否正常, 本次重置数据已完成"

End Function

Public Function GetProductJDObject(productTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

productTemp = Trim(Replace(Replace(productTemp, Chr(13), ""), Chr(10), ""))
'cmdStr = " SELECT  b.料号 FROM [erpdata].[dbo].[TSVtblSetMRule] a, [erpdata].[dbo].[TSVtblMRuleData] b " & _
'         " Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & productTemp & "' "
'
cmdStr = "select   b.FCostObjID from   AIS20141114094336.dbo.t_ICItem a,AIS20141114094336.dbo.cb_CostObj_Product  b " & _
         " Where b.FProductID = a.FItemID and a.F_101='" & productTemp & "' and a.FDeleted<>'1' "
         
         
' cmdStr = "select   b.FCostObjID from   AIS20141114094336.dbo.t_ICItem a,AIS20141114094336.dbo.cb_CostObj_Product  b " & _
'         " Where b.FProductID = a.FItemID and a.F_102='" & productTemp & "' and a.FDeleted<>'1' "



Set RSResult = getSqlServerStr2(cmdStr)

Set GetProductJDObject = RSResult
End Function

'CCS ADD 20160717 检查产品对照表是否建立
Public Function GetProductCheckTT(productTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

productTemp = Trim(Replace(Replace(productTemp, Chr(13), ""), Chr(10), ""))
cmdStr = "SELECT QTECHPTNO2  FROM tbltsvnpiproduct WHERE QTECHPTNO2= '" & productTemp & "' "



Set GetProductCheckTT = getStr(cmdStr)


End Function

' Public Function GetNpiProductCheck(CustName As String, qtechPt As String, custPt1 As String, custPt2 As String, qtechPt2 As String) As ADODB.Recordset

' Dim cmdStr As String
' Dim RSResult As New ADODB.Recordset

' cmdStr = " select qtechptno   from  TBLTsvNpiProduct where customershortname='" & CustName & "' and qtechptno='" & qtechPt & "' and customerptno1='" & custPt1 & "' " & _
         ' " and customerptno2='" & custPt2 & "'  and qtechptno2='" & qtechPt2 & "'  and flag='Y' "

' Set RSResult = getStr(cmdStr)

' Set GetNpiProductCheck = RSResult
' End Function


Public Function GetNpiProductCheck(CustName As String, qtechPt As String, custPt1 As String, custPt2 As String, custPt3 As String, id As Integer, qtechPt2 As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select qtechptno   from  TBLTsvNpiProduct where customershortname='" & CustName & "' and qtechptno='" & qtechPt & "' and customerptno1='" & custPt1 & "' " & _
         " and customerptno2='" & custPt2 & "'  and  ( customerptno3='" & qtechPt2 & "' or customerptno3 is null)  AND ID <>  '" & id & "' " & _
         "  union select qtechptno  from   TBLTsvNpiProduct where qtechptno2='" & qtechPt2 & "' and  customershortname='" & CustName & "' AND ID <>  '" & id & "'  "

Set RSResult = getStr(cmdStr)

Set GetNpiProductCheck = RSResult
End Function


Public Function GetNpiProductCheck1(CustName As String, qtechPt2 As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select qtechptno   from  TBLTsvNpiProduct where customershortname='" & CustName & "' and  qtechptno2='" & qtechPt2 & "'   "

Set RSResult = getStr(cmdStr)

Set GetNpiProductCheck1 = RSResult
End Function


Public Function GetWOData(woTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = "select mfgordername from mfgorder mfg where mfg.mfgordername='" & woTemp & "' "

cmdStr = " select ordername  from erpintegration2.ib_wohistory where ordername='" & woTemp & "' "
             
Set RSResult = getStr(cmdStr)

Set GetWOData = RSResult
End Function

Public Function GetWoData2(woTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = "select mfgordername from mfgorder mfg where mfg.mfgordername='" & woTemp & "' "

cmdStr = " select ordername  from erpintegration2.ib_workorder where ordername='" & woTemp & "' "
             
Set RSResult = getStr(cmdStr)

Set GetWoData2 = RSResult
End Function


Public Function GetWoLableStatus(customerTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "  select customername from TSV_OpenWO_LableStatus  where customername='" & customerTemp & "' and flag='Y' "
              
Set RSResult = getStr(cmdStr)

Set GetWoLableStatus = RSResult
End Function




Public Function CheckWLOWo(woTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

 
 cmdStr = "select ordername from erpintegration2.WLO_IB_WORKORDER where ordername='" & woTemp & "' "
        
Set RSResult = getStr(cmdStr)

Set CheckWLOWo = RSResult
End Function


Public Function GetWLOWo(woTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

 
 cmdStr = "SELECT 工单号 From [ERPBASE].[dbo].[tblllplan] where 工单号='" & woTemp & "' and 产线标记=2 and 实领数量>0 "
  
        
Set RSResult = getSqlServerStr2(cmdStr)

Set GetWLOWo = RSResult
End Function

Public Function GetWLOWoBomLing(woTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


 cmdStr = "SELECT 工单号  From [ERPBASE].[dbo].[tblllplan] where  工单号='" & woTemp & "' and 产线标记=2 and 实领数量 is null "
    
     
Set RSResult = getSqlServerStr2(cmdStr)

Set GetWLOWoBomLing = RSResult
End Function



Public Function GetProduct_Check(productTemp As String) As ADODB.Recordset
'查料号是否存在
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

productTemp = Trim(Replace(Replace(productTemp, Chr(13), ""), Chr(10), ""))
cmdStr = "  select * from  PRODUCTBASE where productname='" & productTemp & "'"

 
        
Set RSResult = getStr(cmdStr)

Set GetProduct_Check = RSResult
End Function



Public Function GetWOCreateDate(woTemp As String) As ADODB.Recordset
'查询开立工单据的日期
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select min(txntimestamp) createDate  from A_Lotwafers al where al.workordername='" & woTemp & "' "
 
        
Set RSResult = getStr(cmdStr)

Set GetWOCreateDate = RSResult
End Function


Public Function GetFps2(sqlTemp As String, customerTemp As String) As ADODB.Recordset
'分批　查mapping数据
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'If customerTemp = "AA" Then
'
'cmdStr = " select id, substrateid,'',passbincount+failbincount,passbincount,lotid,(select b.planned_laser_scribe from CustomerOItbl_test b where substr(b.source_batch_id,1,9) = substr(a.lotid,1,9)) || Get_AA_MarkingCode(substrateid) ,'' " & _
'         "  from   mappingDataTest a where lotid like '" & sqlTemp & "%'  and  customershortname='" + customerTemp + "' and flag='Y' order by substrateid  "
'
'ElseIf customerTemp = "GC" Then
'
'cmdStr = " select id, lotid || substr('0'||wafer_id,-2,2) ,'',passbincount,failbincount,lotid,'' ,'' " & _
'         "  from   mappingDataTest a where lotid like '" & sqlTemp & "%'  and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "
'
'ElseIf customerTemp = "PT" Then
'
'cmdStr = " select id, substrateid ,'',passbincount,failbincount,lotid,'' ,'' " & _
'         "  from   mappingDataTest a where lotid like '" & sqlTemp & "%'  and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "
'
'ElseIf customerTemp = "DN" Then
'
'cmdStr = " select id, substrateid ,'',passbincount,failbincount,lotid,'' ,'' " & _
'         "  from   mappingDataTest a where lotid like '" & sqlTemp & "%'  and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "
'
'ElseIf customerTemp = "MC" Then
'
'cmdStr = " select id, substrateid ,'',passbincount,failbincount,lotid,'' ,'' " & _
'         "  from   mappingDataTest a where lotid like '" & sqlTemp & "%'  and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "
'
'
'Else
'
'cmdStr = " select id, substrateid ,'',passbincount,failbincount,lotid,'' ,'' " & _
'         "  from   mappingDataTest a where lotid like '" & sqlTemp & "%'  and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "
'
'
'cmdStr = " select id, substrateid ,'',passbincount,failbincount,lotid,'' ,'' " & _
'         "  from   mappingDataTest a where lotid like '" & sqlTemp & "%'  and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "
'
'
'End If

If customerTemp = "AA" Then

cmdStr = " select id, substrateid,'',passbincount+failbincount,passbincount,lotid,(select b.planned_laser_scribe from CustomerOItbl_test b where substr(b.source_batch_id,1,9) = substr(a.lotid,1,9)) || Get_AA_MarkingCode(substrateid) ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' )  and  customershortname='" + customerTemp + "' and flag='Y' order by substrateid  "

ElseIf customerTemp = "GC" Then

cmdStr = " select id, lotid || substr('0'||wafer_id,-2,2) ,'',passbincount+failbincount,passbincount,lotid,productid ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' )  and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "

ElseIf customerTemp = "PT" Then

cmdStr = " select id, substrateid ,'',passbincount,failbincount,lotid,productid ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' )  and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "

ElseIf customerTemp = "DN" Then

cmdStr = " select id, substrateid ,'',passbincount,failbincount,lotid,productid ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' )  and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "

ElseIf customerTemp = "MC" Then

cmdStr = " select id, substrateid ,'',passbincount,failbincount,lotid,productid ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' )  and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "


ElseIf customerTemp = "HY" Then

cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,productid,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "

ElseIf customerTemp = "CN" Then

cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,productid ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "

ElseIf customerTemp = "SI" Or customerTemp = "GT" Or customerTemp = "CS" Then

cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,productid,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "

ElseIf customerTemp = "HD" Then

cmdStr = " select id, substrateid ,'',passbincount+failbincount,passbincount,lotid,productid ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' ) and  customershortname='" + customerTemp + "' and flag='Y' order by 2  "

ElseIf customerTemp = "37" Then


cmdStr = " select a.id, a.substrateid ,'',a.passbincount+a.failbincount,a.passbincount,a.lotid,productid ,'' " & _
         "  from   mappingDataTest a,customeroitbl_test c where  a.FileName = to_char(c.ID) and a.lotid in ('" & sqlTemp & "' )  and c.po_num is not null " & _
       "  and  a.customershortname='" + customerTemp + "' and a.flag='Y'   and not exists (select 1 from a_lotwafers b where b.waferscribenumber = a.substrateid) order by 2  "


Else

cmdStr = " select a.id, a.substrateid ,'',a.passbincount+a.failbincount,a.passbincount,a.lotid,productid ,'' " & _
         "  from   mappingDataTest a,customeroitbl_test c where  a.FileName = to_char(c.ID) and a.lotid in ('" & sqlTemp & "' )  " & _
       "  and  a.customershortname='" + customerTemp + "' and a.flag='Y'   and not exists (select 1 from a_lotwafers b where b.waferscribenumber = a.substrateid) order by 2  "


        
End If


        
Set RSResult = getStr(cmdStr)

Set GetFps2 = RSResult
End Function



Public Function GetFps2GCWld(sqlTemp As String, customerTemp As String) As ADODB.Recordset
'分批　查mapping数据
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset





cmdStr = " select id, substrateid ,'',passbincount,failbincount,lotid,'' ,'' " & _
         "  from   mappingDataTest a where lotid in ('" & sqlTemp & "' )  and  customershortname='" + customerTemp + "' and flag='Y' and substrateid like '%-%' order by 2  "





        
Set RSResult = getStr(cmdStr)

Set GetFps2GCWld = RSResult
End Function



Public Function GetFps2AA_00B(sqlTemp As String, customerTemp As String) As ADODB.Recordset
'分批　查mapping数据
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

Dim strKey As String

strKey = Left$(sqlTemp, 9)

If customerTemp = "AA" Then
   
cmdStr = " select id, substrateid,'',passbincount+failbincount,passbincount,lotid,(select b.planned_laser_scribe from CustomerOItbl_test b where source_batch_id = '" & sqlTemp & "') || Get_AA_MarkingCode(substrateid) ,'' " & _
         "  from   mappingDataTest a where lotid  like '" & strKey & "%'  and  customershortname='" + customerTemp + "' and flag='Y' order by substrateid  "


End If
 
        
Set RSResult = getStr(cmdStr)

Set GetFps2AA_00B = RSResult
End Function


Public Function GetFpsWaferDetail() As ADODB.Recordset
'分批　查mapping数据
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select ID, SUBSTRATEID, '', BINCOUNT, PASSBINCOUNT, LOTID, '', ''  From TSV_WO_DetailTemp "
      
Set RSResult = getStr(cmdStr)

Set GetFpsWaferDetail = RSResult
End Function

Public Function GetFpsWLOWaferDetail() As ADODB.Recordset
'分批　查mapping数据
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select id ,substrateid,passbincount from  WLO_WO_DetailTemp  order by id"
      
Set RSResult = getStr(cmdStr)

Set GetFpsWLOWaferDetail = RSResult
End Function





Public Function GetFpsUploadTempData() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select Id ,WaferId ,FinishFlag ,TotalDieQty ,GoodDirQty ,WaferLot ,MarkingCode ,'' from  TSV_Wo_Mt_UploadTemp where flag='Y'  "
        
Set RSResult = getStr(cmdStr)

Set GetFpsUploadTempData = RSResult
End Function







Public Function GetFpsBom(ptTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'  cmdStr = " SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.单位1 " & _
'           " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b" & _
'           " Where a.材料规范编号 = b.材料规范编号 AND (a.物料编号='" & ptTemp & "'   OR a.是否共用标记=1) ORDER BY b.材料规范编号 DESC"
  
 
cmdStr = " SELECT [ID],[MID],[PT],[WLID],[Name],cast(Qty as varchar),[SHRateQty],[Unit],[PT1],[WLID1],[Name1],cast(Qty1 as varchar),[SHRateQty1],[Unit1] " & _
            " FROM [erpdata].[dbo].[TSVtblBillBomInitData] ORDER BY id"
            
 
        
Set RSResult = getStrSqlServer(cmdStr)

Set GetFpsBom = RSResult
End Function


Public Function GetFpsBomNew(ptTemp As String, Userid As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'  cmdStr = " SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.单位1 " & _
'           " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b" & _
'           " Where a.材料规范编号 = b.材料规范编号 AND (a.物料编号='" & ptTemp & "'   OR a.是否共用标记=1) ORDER BY b.材料规范编号 DESC"
  
 
cmdStr = " SELECT [ID],[MID],[PT],[WLID],[Name],cast(Qty as varchar),[SHRateQty],[Unit],[PT1],[WLID1],[Name1],cast(Qty1 as varchar),[SHRateQty1],[Unit1] " & _
            " FROM [erpdata].[dbo].[TSVtblBillBomInitData] where employ='" & Userid & "'  ORDER BY id"
            
 
        
Set RSResult = getStrSqlServer(cmdStr)

Set GetFpsBomNew = RSResult
End Function



Public Function GetFpsBomSetUp() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'  cmdStr = " SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.单位1 " & _
'           " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b" & _
'           " Where a.材料规范编号 = b.材料规范编号 AND (a.物料编号='" & ptTemp & "'   OR a.是否共用标记=1) ORDER BY b.材料规范编号 DESC"
  
 
'cmdStr = " SELECT [ID],[MID],[PT],[WLID],[Name],cast(Qty as varchar),[SHRateQty],[Unit],[PT1],[WLID1],[Name1],cast(Qty1 as varchar),[SHRateQty1],[Unit1] " & _
'            " FROM [erpdata].[dbo].[TSVtblBillBomInitData] ORDER BY id"
'
            
   
cmdStr = "  SELECT a.id,a.ProductName,a.childBomNameID,b.物料编号,a.CreateDate" & _
         "  FROM [erpdata].[dbo].[TSVtblBomSetup] a, [erpdata].[dbo].[TSVtblSetMRule] b  Where b.材料规范编号 = a.childBomNameID order by a.id desc "
  
            
 
        
Set RSResult = getStrSqlServer(cmdStr)

Set GetFpsBomSetUp = RSResult
End Function



Public Function GetFpsWLOBom(ptTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'  cmdStr = " SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.单位1 " & _
'           " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b" & _
'           " Where a.材料规范编号 = b.材料规范编号 AND (a.物料编号='" & ptTemp & "'   OR a.是否共用标记=1) ORDER BY b.材料规范编号 DESC"
  
 
cmdStr = " SELECT [ID],[MID],[PT],[WLID],[Name],cast(Qty as varchar),[SHRateQty],[Unit],[PT1],[WLID1],[Name1],cast(Qty1 as varchar),[SHRateQty1],[Unit1] " & _
            " FROM [erpdata].[dbo].[WLOtblBillBomInitData] ORDER BY id"
 
        
Set RSResult = getStrSqlServer(cmdStr)

Set GetFpsWLOBom = RSResult
End Function

Public Function GetMonthChar(monThTemp As Integer) As String
If monThTemp = 10 Then
GetMonthChar = "A"

ElseIf monThTemp = 11 Then
GetMonthChar = "B"

ElseIf monThTemp = 12 Then
GetMonthChar = "C"
    
Else
GetMonthChar = monThTemp
    
End If


End Function


Public Function GetFpsReworkBom(ptTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'  cmdStr = " SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.单位1 " & _
'           " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b" & _
'           " Where a.材料规范编号 = b.材料规范编号 AND (a.物料编号='" & ptTemp & "'   OR a.是否共用标记=1) ORDER BY b.材料规范编号 DESC"
  
 
cmdStr = " SELECT [ID],'',[MID],[PT],[WLID],[Name],cast(Qty as varchar),[SHRateQty],[Unit],[PT1],[WLID1],[Name1],cast(Qty1 as varchar),[SHRateQty1],[Unit1] " & _
            " FROM [erpdata].[dbo].[TSVtblBillBom2InitData] ORDER BY id"
 
        
Set RSResult = getStrSqlServer(cmdStr)

Set GetFpsReworkBom = RSResult
End Function


Public Function GetFpsReworkBomNew(ptTemp As String, Userid As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'  cmdStr = " SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.单位1 " & _
'           " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b" & _
'           " Where a.材料规范编号 = b.材料规范编号 AND (a.物料编号='" & ptTemp & "'   OR a.是否共用标记=1) ORDER BY b.材料规范编号 DESC"
  
 
cmdStr = " SELECT [ID],'',[MID],[PT],[WLID],[Name],cast(Qty as varchar),[SHRateQty],[Unit],[PT1],[WLID1],[Name1],cast(Qty1 as varchar),[SHRateQty1],[Unit1] " & _
            " FROM [erpdata].[dbo].[TSVtblBillBom2InitData] where employ ='" & Userid & "' ORDER BY id"
 
        
Set RSResult = getStrSqlServer(cmdStr)

Set GetFpsReworkBomNew = RSResult
End Function



Public Function GetFpsBomQuery(ptTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'  cmdStr = " SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.单位1 " & _
'           " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b" & _
'           " Where a.材料规范编号 = b.材料规范编号 AND (a.物料编号='" & ptTemp & "'   OR a.是否共用标记=1) ORDER BY b.材料规范编号 DESC"
  
 
'cmdStr = " SELECT [ID],[MID],[PT],[WLID],[Name],cast(Qty as varchar),[SHRateQty],[Unit],[PT1],[WLID1],[Name1],cast(Qty1 as varchar),[SHRateQty1],[Unit1] " & _
'            " FROM [erpdata].[dbo].[TSVtblBillBomInitData] ORDER BY id"
 
        
Set RSResult = getStrSqlServer(ptTemp)

Set GetFpsBomQuery = RSResult
End Function





Public Function GetSeqChar() As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select nvl(max(substr(ordername,-4)),0) id from IB_WOHISTORY  where substr(ordername,5,4) =  to_char(sysdate,'YYMM')"
     
RSResult = GetSeq(cmdStr)
GetSeqChar = RSResult
End Function


Public Function GetString(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select  Get_PFType_LotId('" & LOTID & "')   from dual"
       
RSResult = getStr2(cmdStr)
GetString = RSResult
End Function


Public Function GetWaferidCus(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select customershortname from mappingdatatest a where a.substrateid='" & LOTID & "'"


       
RSResult = getStr2(cmdStr)
GetWaferidCus = RSResult
End Function

Public Function GetOICustomerPTNum(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String



cmdStr = " select a.mpn_desc  from  customeroitbl_test  a  where a.source_batch_id='" & LOTID & "'  and " & _
         " a.id in (select max(b.id)  from  customeroitbl_test  b  where b.source_batch_id=a.source_batch_id) "
         
       
RSResult = getStr2(cmdStr)
GetOICustomerPTNum = RSResult
End Function



Public Function GetAADelQboxNo1(LOTID As String) As String
'查询箱号
Dim cmdStr As String
Dim RSResult As String


         
cmdStr = "select distinct a.qboxnumber  from  tsv_qboxnumber_details a" & _
" where a.containername='" & LOTID & "' and A.Customername='AA' and (a.containername like '%-E' or a.containername like '%-F') and a.qboxnumber like 'SP216%' "
         
       
RSResult = getStr2(cmdStr)
GetAADelQboxNo1 = RSResult
End Function

Public Function GetAADelQboxNo2(LOTID As String) As String
'查询箱号
Dim cmdStr As String
Dim RSResult As String


         
cmdStr = "select distinct a.qboxnumber  from  tsv_qboxnumber_details a" & _
" where a.containername='" & LOTID & "' and A.Customername='AA' and (a.containername like '%-E' or a.containername like '%-F') and a.qboxnumber like  'Q16%' "
         
       
RSResult = getStr2(cmdStr)
GetAADelQboxNo2 = RSResult
End Function




Public Function GetNpiCustomerPTNum(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String



cmdStr = " select qtechptno2 from  TBLTsvNpiProduct a where a.customerptno1='" & LOTID & "'  or  a.customerptno2='" & LOTID & "' or  a.customerptno3='" & LOTID & "' or  a.customerptno4='" & LOTID & "' "
       
RSResult = getStr2(cmdStr)
GetNpiCustomerPTNum = RSResult
End Function





Public Function GetGC400NString(LOTID As String, qboxnumberTemp As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String


'cmdStr = " SELECT workorderattr3 ||','|| sum(QTY1) ||','||wafernumber||','||'SH' ||','||imager_customer_rev|| ','||count(waferscribenumber1) ||','|| " & _
'         " LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ') ||','||wafernumber||'.'||min(waferscribenumber1)  as txt" & _
'         " FROM (select  f.workorderattr3 ,sum(b.qty) as QTY1,a.wafernumber,'SH' as comp_code,g.imager_customer_rev,substr(a.waferscribenumber, -2, 2) waferscribenumber1," & _
'         " ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN " & _
'         " from a_lotwafers a,container b,a_lotattributes c,product d,productbase e,mfgorder f,customeroitbl_test g " & _
'         " Where a.containerid = b.containerid And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
'         " and f.mfgordername = a.workordername and b.productid = d.productid and b.containername in ('" & lotid & "') and c.customername = 'GC' " & _
'         " and g.source_batch_id=a.wafernumber GROUP BY f.workorderattr3 ,a.wafernumber,g.comp_code,g.imager_customer_rev ,a.waferscribenumber) " & _
'         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
'         " GROUP BY workorderattr3,QTY1,wafernumber,comp_code,imager_customer_rev  order by wafernumber"
         
'      cmdStr = " SELECT workorderattr3 ||','|| sum(QTY1) ||','||wafernumber||','||'SH' ||','||imager_customer_rev|| ','||count(waferscribenumber1) ||','|| " & _
'         " LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ') ||','||wafernumber||'.'||min(waferscribenumber1)  as txt" & _
'         " FROM (select  f.workorderattr3 ,sum(b.qty) as QTY1,a.wafernumber,'SH' as comp_code,g.imager_customer_rev,substr(a.waferscribenumber, -2, 2) waferscribenumber1," & _
'         " ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN " & _
'         " from a_lotwafers a,container b,a_lotattributes c,product d,productbase e,mfgorder f,customeroitbl_test g ,  mappingdatatest h  " & _
'         " Where a.containerid = b.containerid And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
'         " and f.mfgordername = a.workordername and b.productid = d.productid and b.containername in ('" & lotid & "') and c.customername = 'GC'  and h.lotid=g.source_batch_id and h.filename=g.id and h.substrateid in ('H8W67605',    'H8W67606',  'H8W67608' )  " & _
'         " and g.source_batch_id=a.wafernumber GROUP BY f.workorderattr3 ,a.wafernumber,g.comp_code,g.imager_customer_rev ,a.waferscribenumber) " & _
'         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
'         " GROUP BY workorderattr3,QTY1,wafernumber,comp_code,imager_customer_rev  order by wafernumber"
         
'            cmdStr = " SELECT workorderattr3 ||','|| sum(QTY1) ||','||wafernumber||','||'SH' ||','||imager_customer_rev|| ','||count(waferscribenumber1) ||','|| " & _
'         " LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ') ||','||wafernumber||'.'||min(waferscribenumber1)  as txt" & _
'         " FROM (select  f.workorderattr3 ,sum(b.qty) as QTY1,a.wafernumber,'SH' as comp_code,g.imager_customer_rev,substr(a.waferscribenumber, -2, 2) waferscribenumber1," & _
'         " ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN " & _
'         " from a_lotwafers a,container b,a_lotattributes c,product d,productbase e,mfgorder f,customeroitbl_test g ,  mappingdatatest h  " & _
'         " Where a.containerid = b.containerid And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
'         " and f.mfgordername = a.workordername and b.productid = d.productid and b.containername in ('" & lotid & "') and c.customername = 'GC'  and h.lotid=g.source_batch_id and h.filename=g.id  " & _
'         " and g.source_batch_id=a.wafernumber and h.substrateid=a.waferscribenumber GROUP BY f.workorderattr3 ,a.wafernumber,g.comp_code,g.imager_customer_rev ,a.waferscribenumber) " & _
'         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
'         " GROUP BY workorderattr3,QTY1,wafernumber,comp_code,imager_customer_rev  order by wafernumber"
         
         

'            cmdStr = " SELECT workorderattr3||'-2.5' || ','||wafernumber||',' ||imager_customer_rev||','  " & _
'         " ||'SH'||','||count(waferscribenumber1)||','|| LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ';')), ';') ||','||wafernumber || '.' || min(waferscribenumber1) ||','|| '" & qboxnumberTemp & "' as txt " & _
'         " FROM (select  f.workorderattr3 ,sum(b.qty) as QTY1,a.wafernumber,'SH' as comp_code,g.imager_customer_rev,substr(a.waferscribenumber, -2, 2) waferscribenumber1," & _
'         " ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN " & _
'         " from a_lotwafers a,container b,a_lotattributes c,product d,productbase e,mfgorder f,customeroitbl_test g ,  mappingdatatest h  " & _
'         " Where a.containerid = b.containerid And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
'         " and f.mfgordername = a.workordername and b.productid = d.productid and b.containername in ('" & LOTID & "') and c.customername = 'GC'  and h.lotid=g.source_batch_id and h.filename=g.id  " & _
'         " and g.source_batch_id=a.wafernumber and h.substrateid=a.waferscribenumber GROUP BY f.workorderattr3 ,a.wafernumber,g.comp_code,g.imager_customer_rev ,a.waferscribenumber) " & _
'         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
'         " GROUP BY workorderattr3,QTY1,wafernumber,comp_code,imager_customer_rev  order by wafernumber"
'
' cmdStr = "SELECT workorderattr3 ||','|| wafernumber||','|| LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ')||','|| count(waferscribenumber1) ||','||cdate  " & _
'          " FROM (select f.workorderattr3||'/'||d.alternatename as workorderattr3, a.wafernumber, " & _
'          " substr(a.waferscribenumber, -2, 2) waferscribenumber1, ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN,to_char(sysdate,'YYYY-MM-DD') as cdate " & _
'          " from a_lotwafers a, container  b,a_lotattributes c,product d, productbase e, mfgorder f " & _
'          " Where a.containerid = b.containerid  And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
'          " and f.mfgordername = a.workordername and b.productid = d.productid " & _
'          " and b.containername in ('" & lotid & "')) " & _
'          "  START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
'          "  GROUP BY workorderattr3, wafernumber,cdate order by wafernumber"
'
         
              cmdStr = " SELECT workorderattr3||'-2.5' || ','||wafernumber||',' ||imager_customer_rev||','  " & _
         " ||'SH'||','||count(waferscribenumber1)||','|| LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ';')), ';') ||','||wafernumber || '.' || min(waferscribenumber1) ||','|| '" & qboxnumberTemp & "' as txt " & _
         " FROM (select  f.workorderattr3 ,sum(A.NDPW) as QTY1,a.wafernumber,'SH' as comp_code,g.imager_customer_rev,substr(a.waferscribenumber, -2, 2) waferscribenumber1," & _
         " ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN " & _
         " from a_lotwafers a,a_lotattributes c,mfgorder f,customeroitbl_test g ,  mappingdatatest h  " & _
         " WHERE f.mfgordername = a.workordername and A.WAFERSCRIBENUMBER in ('" & LOTID & "') and c.customername = 'GC'  and h.lotid=g.source_batch_id and h.filename=g.id  " & _
         " and g.source_batch_id=a.wafernumber and h.substrateid=a.waferscribenumber GROUP BY f.workorderattr3 ,a.wafernumber,g.comp_code,g.imager_customer_rev ,a.waferscribenumber) " & _
         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
         " GROUP BY workorderattr3,QTY1,wafernumber,comp_code,imager_customer_rev  order by wafernumber"
         
    
RSResult = getStr2(cmdStr)
GetGC400NString = RSResult
End Function




Public Function Get37InQboxSeqTxt(LOTID As String, qboxnumberTemp As String, containerTemp As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String


cmdStr = " select get_37_LableID('INQbox','" & containerTemp & "',max(a.htlotid),'','') as inQboxID " & _
" from  TSV_Tray_details a where trayqboxnumber in ('" & LOTID & "') " & _
" group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htlotid,a.htdatecode"

RSResult = getStr2(cmdStr)
Get37InQboxSeqTxt = RSResult
End Function










Public Function Get37OutQboxHTTxt(LOTID As String, qboxnumberTemp As String, containerTemp As String) As String

Dim cmdStr As String
Dim RSResult As String


cmdStr = " select  trglabelseq.QTSeq_NotMesQbox('" & containerTemp & "')  from dual "


RSResult = getStr2(cmdStr)
Get37OutQboxHTTxt = RSResult
End Function



'Public Function Get37Dn() As ADODB.Recordset
''DN
'Dim cmdStr As String
'Dim RSResult As New ADODB.Recordset
'
'cmdStr = "select distinct delivery  as DNName, delivery  as DNid  from   CUSTOMERSHIPPINGUPTBL  a where a.flag='Y' order by a.delivery desc"
'
'Set RSResult = getStr(cmdStr)
'Set Get37Dn = RSResult
'End Function

Public Function Get37Dnbatch() As ADODB.Recordset
'DN
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select batchnumber as DNBatch from   CUSTOMERSHIPPINGUPTBL  a where a.flag='Y' order by a.delivery desc"

Set RSResult = getStr(cmdStr)
Set Get37Dnbatch = RSResult
End Function


Public Function Get37CustomerList() As ADODB.Recordset
'DN
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset



cmdStr = " select distinct customername as DNName,  customername as DNid     from CUSTOMERQboxSetData "


Set RSResult = getStr(cmdStr)
Set Get37CustomerList = RSResult
End Function


Public Function Get37PTList(tempDN As String) As ADODB.Recordset
'DN
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


cmdStr = "select q.PRODUCT as DNName,  q.PRODUCT as DNid from [erpdata].[dbo].[tblPackToHouseRec] d , [erpdata].[dbo].[tblTSVworkorder] q where 工单号 in (SELECT a.BatchNumber FROM [ERPBASE].[dbo].[tblCustomerShippingUp] a where a.Delivery = '" & tempDN & "' ) and d.大工单 = q.ORDERNAME"
'cmdStr = " select distinct pt as DNName,  pt as DNid     from CUSTOMERQboxSetData "


Set RSResult = getSqlStr(cmdStr)
Set Get37PTList = RSResult
End Function




Public Function Get37QboxTypeList() As ADODB.Recordset
'DN
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset



cmdStr = " select distinct qboxType as DNName,  qboxType as DNid     from CUSTOMERQboxSetData "


Set RSResult = getStr(cmdStr)
Set Get37QboxTypeList = RSResult
End Function






Public Function Get37QboxList(tempDN As String) As ADODB.Recordset
'DN
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select  箱号 as DNName, 箱号 as DNID from [erpdata].[dbo].[tblPackTreeInf] c where c.序号 in (  select 上级序号 from [erpdata].[dbo].[tblPackTreeInf] a ,[erpdata].[dbo].[tblPackToHouseRec] b where b.工单号 in (SELECT g.BatchNumber FROM [ERPBASE].[dbo].[tblCustomerShippingUp] g where g.Delivery = '" & tempDN & "' ) and a.入库单编号 = b.入库单编号 )"

'cmdStr = " select distinct ltrim(CONTAINERNAME) as DNName , ltrim(CONTAINERNAME) as  DNID from  [erpdata].[dbo].[TblTSV_OutBOX_DETAILS] order by 1 "


Set RSResult = getSqlStr(cmdStr)
Set Get37QboxList = RSResult
End Function







Public Function GetWaiBaoString(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String


'cmdStr = " SELECT workorderattr3 ||','|| sum(QTY1) ||','||wafernumber||','||'SH' ||','||imager_customer_rev|| ','||count(waferscribenumber1) ||','|| " & _
'         " LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ') ||','||wafernumber||'.'||min(waferscribenumber1)  as txt" & _
'         " FROM (select  f.workorderattr3 ,sum(b.qty) as QTY1,a.wafernumber,'SH' as comp_code,g.imager_customer_rev,substr(a.waferscribenumber, -2, 2) waferscribenumber1," & _
'         " ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN " & _
'         " from a_lotwafers a,container b,a_lotattributes c,product d,productbase e,mfgorder f,customeroitbl_test g " & _
'         " Where a.containerid = b.containerid And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
'         " and f.mfgordername = a.workordername and b.productid = d.productid and b.containername in ('" & lotid & "') and c.customername = 'GC' " & _
'         " and g.source_batch_id=a.wafernumber GROUP BY f.workorderattr3 ,a.wafernumber,g.comp_code,g.imager_customer_rev ,a.waferscribenumber) " & _
'         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
'         " GROUP BY workorderattr3,QTY1,wafernumber,comp_code,imager_customer_rev  order by wafernumber"
         
'      cmdStr = " SELECT workorderattr3 ||','|| sum(QTY1) ||','||wafernumber||','||'SH' ||','||imager_customer_rev|| ','||count(waferscribenumber1) ||','|| " & _
'         " LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ') ||','||wafernumber||'.'||min(waferscribenumber1)  as txt" & _
'         " FROM (select  f.workorderattr3 ,sum(b.qty) as QTY1,a.wafernumber,'SH' as comp_code,g.imager_customer_rev,substr(a.waferscribenumber, -2, 2) waferscribenumber1," & _
'         " ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN " & _
'         " from a_lotwafers a,container b,a_lotattributes c,product d,productbase e,mfgorder f,customeroitbl_test g ,  mappingdatatest h  " & _
'         " Where a.containerid = b.containerid And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
'         " and f.mfgordername = a.workordername and b.productid = d.productid and b.containername in ('" & lotid & "') and c.customername = 'GC'  and h.lotid=g.source_batch_id and h.filename=g.id and h.substrateid in ('H8W67605',    'H8W67606',  'H8W67608' )  " & _
'         " and g.source_batch_id=a.wafernumber GROUP BY f.workorderattr3 ,a.wafernumber,g.comp_code,g.imager_customer_rev ,a.waferscribenumber) " & _
'         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
'         " GROUP BY workorderattr3,QTY1,wafernumber,comp_code,imager_customer_rev  order by wafernumber"
         
'            cmdStr = " SELECT workorderattr3 ||','|| sum(QTY1) ||','||wafernumber||','||'SH' ||','||imager_customer_rev|| ','||count(waferscribenumber1) ||','|| " & _
'         " LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ') ||','||wafernumber||'.'||min(waferscribenumber1)  as txt" & _
'         " FROM (select  f.workorderattr3 ,sum(b.qty) as QTY1,a.wafernumber,'SH' as comp_code,g.imager_customer_rev,substr(a.waferscribenumber, -2, 2) waferscribenumber1," & _
'         " ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN " & _
'         " from a_lotwafers a,container b,a_lotattributes c,product d,productbase e,mfgorder f,customeroitbl_test g ,  mappingdatatest h  " & _
'         " Where a.containerid = b.containerid And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
'         " and f.mfgordername = a.workordername and b.productid = d.productid and b.containername in ('" & lotid & "') and c.customername = 'GC'  and h.lotid=g.source_batch_id and h.filename=g.id  " & _
'         " and g.source_batch_id=a.wafernumber and h.substrateid=a.waferscribenumber GROUP BY f.workorderattr3 ,a.wafernumber,g.comp_code,g.imager_customer_rev ,a.waferscribenumber) " & _
'         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
'         " GROUP BY workorderattr3,QTY1,wafernumber,comp_code,imager_customer_rev  order by wafernumber"
         
         

'            cmdStr = " SELECT workorderattr3||'-2.5' || ','||wafernumber||',' ||imager_customer_rev||','  " & _
'         " ||'SH'||','||count(waferscribenumber1)||','|| LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ';')), ';') ||','||wafernumber || '.' || min(waferscribenumber1) as txt " & _
'         " FROM (select  f.workorderattr3 ,sum(b.qty) as QTY1,a.wafernumber,'SH' as comp_code,g.imager_customer_rev,substr(a.waferscribenumber, -2, 2) waferscribenumber1," & _
'         " ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN " & _
'         " from a_lotwafers a,container b,a_lotattributes c,product d,productbase e,mfgorder f,customeroitbl_test g ,  mappingdatatest h  " & _
'         " Where a.containerid = b.containerid And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
'         " and f.mfgordername = a.workordername and b.productid = d.productid and b.containername in ('" & lotid & "') and c.customername = 'GC'  and h.lotid=g.source_batch_id and h.filename=g.id  " & _
'         " and g.source_batch_id=a.wafernumber and h.substrateid=a.waferscribenumber GROUP BY f.workorderattr3 ,a.wafernumber,g.comp_code,g.imager_customer_rev ,a.waferscribenumber) " & _
'         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
'         " GROUP BY workorderattr3,QTY1,wafernumber,comp_code,imager_customer_rev  order by wafernumber"
         
 cmdStr = "SELECT workorderattr3 ||','|| wafernumber||','|| LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ')||','|| count(waferscribenumber1) ||','||cdate  " & _
          " FROM (select d.alternatename as workorderattr3, a.wafernumber, " & _
          " substr(a.waferscribenumber, -2, 2) waferscribenumber1, ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN,to_char(sysdate,'YYYY-MM-DD') as cdate " & _
          " from a_lotwafers a, container  b,a_lotattributes c,product d, productbase e, mfgorder f " & _
          " Where a.containerid = b.containerid  And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
          " and f.mfgordername = a.workordername and b.productid = d.productid " & _
          " and b.containername in ('" & LOTID & "')) " & _
          "  START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
          "  GROUP BY workorderattr3, wafernumber,cdate order by wafernumber"
          
          
          
'  cmdStr = " SELECT workorderattr3 ||','|| wafernumber||','|| LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ')||','|| count(waferscribenumber1) ||','||sum(qty)||','||cdate " & _
'           " FROM (select f.workorderattr3||'/'||d.alternatename as workorderattr3, a.wafernumber, b.qty , " & _
'           " substr(a.waferscribenumber, -2, 2) waferscribenumber1, ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN,to_char(sysdate,'YYYY-MM-DD') as cdate " & _
'" from a_lotwafers a, container  b,a_lotattributes c,product d, productbase e, mfgorder f " & _
'" Where a.containerid = b.containerid And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
'" and f.mfgordername = a.workordername and b.productid = d.productid " & _
'" and b.containername in ('" & lotid & "')) " & _
'" START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
'" GROUP BY workorderattr3, wafernumber,cdate order by wafernumber "

         
         
    
RSResult = getStr2(cmdStr)
GetWaiBaoString = RSResult
End Function






Public Function GetLiePian45String(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

cmdStr = " SELECT workorderattr3 || ',' || waferlot || ',' || LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ') || ',' || " & _
      " count(waferscribenumber1) || ',' || cdate FROM ( " & _
" select d.alternatename||'/'||b.customerpn as workorderattr3,a.waferlot,  substr(a.waferid, -2, 2) waferscribenumber1,  ROW_NUMBER() OVER(PARTITION BY a.waferlot ORDER BY a.waferid) RN, " & _
" to_char(sysdate, 'YYYY-MM-DD') as cdate " & _
" from ib_waferlist a , ib_wohistory b ,productbase c, product d where a.waferid in ('" & LOTID & "') " & _
" and b.ordername=a.ordername and c.productname=b.product " & _
" and d.productbaseid=c.productbaseid) START WITH RN = 1 " & _
" CONNECT BY RN - 1 = PRIOR RN AND waferlot = PRIOR waferlot " & _
" GROUP BY workorderattr3, waferlot, cdate order by waferlot "
 
          
RSResult = getStr2(cmdStr)
GetLiePian45String = RSResult
End Function





Public Function GetWaiBao36String(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

          
  cmdStr = " SELECT workorderattr3 ||','|| wafernumber||','|| LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ')||','|| count(waferscribenumber1) ||','||sum(qty)||','||cdate " & _
           " FROM (select  d.alternatename||'/'||f.workorderattr3  as workorderattr3, a.wafernumber, b.qty , " & _
           " substr(a.waferscribenumber, -2, 2) waferscribenumber1, ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN,to_char(sysdate,'YYYY-MM-DD') as cdate " & _
" from a_lotwafers a, container  b,a_lotattributes c,product d, productbase e, mfgorder f " & _
" Where a.containerid = b.containerid And b.containerid = c.containerid And d.productbaseid = e.productbaseid " & _
" and f.mfgordername = a.workordername and b.productid = d.productid " & _
" and b.containername in ('" & LOTID & "')) " & _
" START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
" GROUP BY workorderattr3, wafernumber,cdate order by wafernumber "

         
         
    
RSResult = getStr2(cmdStr)
GetWaiBao36String = RSResult
End Function


Public Function GetWaiBaoQRString(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

          

  cmdStr = " SELECT workorderattr3||','||wafernumber||','||LTRIM(MAX(SYS_CONNECT_BY_PATH(waferscribenumber1, ' ')), ' ') ||','||sum(PCS)||','||sum(QTY1)||','||FAB_CONV_ID||','||Pdate " & _
" FROM ( select distinct e.workorderattr3,a.wafernumber,substr(a.waferscribenumber, -2, 2) waferscribenumber1 " & _
" ,count(a.waferscribenumber) PCS,SUM(b.qty) AS QTY1,d.FAB_CONV_ID,to_char(sysdate, 'yyyy-mm-dd') as Pdate " & _
" ,ROW_NUMBER() OVER(PARTITION BY a.wafernumber ORDER BY a.waferscribenumber) RN " & _
" from a_lotwafers a,container b,a_lotattributes c,customeroitbl_test d,mappingdatatest f,mfgorder e " & _
" Where a.containerid = b.containerid And b.containerid = c.containerid And a.WAFERNUMBER = d.SOURCE_BATCH_ID " & _
" and a.WORKORDERNAME=e.MFGORDERNAME AND b.CONTAINERNAME in ('" & LOTID & "') " & _
" and c.customername = 'QR' and f.filename=d.id and f.customershortname='QR' " & _
" and f.substrateid=a.waferscribenumber GROUP BY workorderattr3,wafernumber,waferscribenumber,FAB_CONV_ID " & _
" ) START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN AND wafernumber = PRIOR wafernumber " & _
" GROUP BY workorderattr3,wafernumber,Pdate,FAB_CONV_ID,pcs order by wafernumber "



RSResult = getStr2(cmdStr)
GetWaiBaoQRString = RSResult
End Function









Public Function GetMPNString(LOTID As String) As String
'取MPN
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select mpn from customeroitbl_test where customershortname='AA' and source_batch_id='" & LOTID & "'"

RSResult = getStr2(cmdStr)
GetMPNString = RSResult
End Function

Public Function GetGCWLDMaringCode(LOTID As String) As String
'取MPN
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select productid from mappingdatatest  a where a.customershortname='GC' and substrateid='" & LOTID & "' and rownum<2"


RSResult = getStr2(cmdStr)
GetGCWLDMaringCode = RSResult
End Function


Public Function GetMtrlNumString(LOTID As String) As String
'取MPN
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select mtrl_num from customeroitbl_test where customershortname='AA' and source_batch_id='" & LOTID & "'"

RSResult = getStr2(cmdStr)
GetMtrlNumString = RSResult
End Function

Public Function GetDateCodeString(LOTID As String) As String
'DateCode
Dim cmdStr As String
Dim RSResult As String

cmdStr = " select  to_char(erpcreatedate, 'yyyy') || to_char(erpcreatedate, 'WW') AS DATA_CODE1 from ( " & _
         " select max(b.erpcreatedate) erpcreatedate  from ib_waferlist a ,ib_wohistory b where a.ordername=b.ordername and a.waferlot='" & LOTID & "') X "

RSResult = getStr2(cmdStr)
GetDateCodeString = RSResult
End Function

Public Function GetWaferFQty(LOTID As String) As String
'DateCode
Dim cmdStr As String
Dim RSResult As String

cmdStr = "  select factorystartqty from container   where containername like '" & LOTID & "%' and containername='" & LOTID & "-F' "
 
RSResult = getStr2(cmdStr)
GetWaferFQty = RSResult
End Function

Public Function GetPJPOData(potemp As String, devicetemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


cmdStr = " select a.po_num from TSV_MD_POPrice a where a.po_num = '" & potemp & "'  and a.pt = '" & devicetemp & "' "
             
Set RSResult = getStr(cmdStr)

Set GetPJPOData = RSResult
End Function
Public Function GetWaferEQty(LOTID As String) As String
'DateCode
Dim cmdStr As String
Dim RSResult As String

cmdStr = "  select factorystartqty from container   where containername like '" & LOTID & "%' and containername='" & LOTID & "-E' "
 
RSResult = getStr2(cmdStr)
GetWaferEQty = RSResult
End Function

Public Function GetWaferAQty(LOTID As String) As String
'DateCode
Dim cmdStr As String
Dim RSResult As String

cmdStr = "  select factorystartqty from container   where containername like '" & LOTID & "%' and containername='" & LOTID & "-A' "
 
RSResult = getStr2(cmdStr)
GetWaferAQty = RSResult
End Function



Public Function GetWaferCustomerFQty(LOTID As String) As String
'DateCode
Dim cmdStr As String
Dim RSResult As String

cmdStr = " select failbincount from mappingdatatest where substrateid='" & LOTID & "'"


RSResult = getStr2(cmdStr)
GetWaferCustomerFQty = RSResult
End Function

Public Function GetWaferCustomerMapQty(LOTID As String) As String
'DateCode
Dim cmdStr As String
Dim RSResult As String

cmdStr = " select passbincount+failbincount as qty from mappingdatatest where substrateid='" & LOTID & "'"


RSResult = getStr2(cmdStr)
GetWaferCustomerMapQty = RSResult
End Function



Public Function GetWafeRejectCodeQty(LOTID As String) As String
'DateCode
Dim cmdStr As String
Dim RSResult As String

cmdStr = "  select sum(X.waferrejectsqty) qty from (select distinct  lossreasonid,waferrejectsqty from  A_WIPLOTREJECTSHISTORY where waferscribenumber='" & LOTID & "' )X "


RSResult = getStr2(cmdStr)
GetWafeRejectCodeQty = RSResult
End Function


Public Function GetBaofeiRejectMailName(idTemp As Long) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

cmdStr = " select b.mailname from  TBLTSVBaoFei a , AutoMailList b where a.id=" & idTemp & " and b.systemname='BaoFeiSys_Reject'  and b.sendtype='To' and b.deptname=a.created_by "
       
RSResult = getStr2(cmdStr)
GetBaofeiRejectMailName = RSResult
End Function




Public Function GetTrayString(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select  Get_TrayType_LotId('" & LOTID & "')   from dual"
       
RSResult = getStr2(cmdStr)
GetTrayString = RSResult
End Function

Public Function GetON_HTKJPT(ptTemp As String) As String
'AA取厂内料号
Dim cmdStr As String
Dim RSResult As String

cmdStr = " select  qtechptno2 from  TBLTsvNpiProduct a where a.customerptno1='" & ptTemp & "' and flag='Y' "
       
RSResult = getStr2(cmdStr)
GetON_HTKJPT = RSResult
End Function

Public Function GetHWMonthMaxQty() As Long
'取HW本月最大号
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select count(*)+1 as ID from mappingdatatest where customershortname='MG' and to_char(qtech_created_date,'YYYY-MM')=to_char(sysdate,'YYYY-MM') "
       
RSResult = GetSeq(cmdStr)
GetHWMonthMaxQty = RSResult
End Function


Public Function GetTestNoString(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select  Get_TestNo_LotId('" & LOTID & "')   from dual"
       
RSResult = getStr2(cmdStr)
GetTestNoString = RSResult
End Function


Public Function GetFirstPtString(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

'cmdStr = " select '18' ||  b.sc3pname01  fristStr from CustomerOItbl_test a  ,tblPDM b where a.source_batch_id='" & LotId & "' and a.mpn=b.fg_material and rownum<2 "
'2013-12-09 jiayun add

cmdStr = " select '18' ||  b.sc3pname01  fristStr from CustomerOItbl_test a  ,tblPDM b where a.source_batch_id='" & LOTID & "' and a.mpn=b.fg_material  and a.mtrl_num=b.turnkey_material and rownum<2 "
       
RSResult = getStr2(cmdStr)
GetFirstPtString = RSResult
End Function



Public Function GetGTTxtProduct(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

''cmdStr = " select '18' ||  b.sc3pname01  fristStr from CustomerOItbl_test a  ,tblPDM b where a.source_batch_id='" & LotId & "' and a.mpn=b.fg_material and rownum<2 "
''2013-12-09 jiayun add
'
'cmdStr = " select '18' ||  b.sc3pname01  fristStr from CustomerOItbl_test a  ,tblPDM b where a.source_batch_id='" & lotID & "' and a.mpn=b.fg_material  and a.mtrl_num=b.turnkey_material and rownum<2 "
'
       
cmdStr = "select g.mpn_desc from  customeroitbl_test g,mappingdatatest x where  g.source_batch_id=x.lotid and x.substrateid='" & LOTID & "' and g.customershortname in ('GT','SI')"
       
       
RSResult = getStr2(cmdStr)
GetGTTxtProduct = RSResult
End Function



Public Function GetGTTxtPackNo(LOTID As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

''cmdStr = " select '18' ||  b.sc3pname01  fristStr from CustomerOItbl_test a  ,tblPDM b where a.source_batch_id='" & LotId & "' and a.mpn=b.fg_material and rownum<2 "
''2013-12-09 jiayun add
'
'cmdStr = " select '18' ||  b.sc3pname01  fristStr from CustomerOItbl_test a  ,tblPDM b where a.source_batch_id='" & lotID & "' and a.mpn=b.fg_material  and a.mtrl_num=b.turnkey_material and rownum<2 "
'
       
'cmdStr = "select g.mpn_desc from  customeroitbl_test g,mappingdatatest x where  g.source_batch_id=x.lotid and x.substrateid='" & lotID & "' and g.customershortname in ('GT','SI')"
'
cmdStr = "select  get_GT_VT_qbox('" & LOTID & "') as packNo from dual "
       
       
RSResult = getStr2(cmdStr)
GetGTTxtPackNo = RSResult
End Function


Public Function GetGTqboxNo(waferid As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

cmdStr = " select CONTAINERID from container where containername like '" & waferid & "%' "
       
       
RSResult = getStr2(cmdStr)
GetGTqboxNo = RSResult
End Function


Public Function GetAllPtString(productNameTemp As String, pfstatusTemp As String, trayTemp As String, testnoTemp As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

'cmdStr = " select productname from  TBLSETPT  where productname like '" & productNameTemp & "%' and pfstaus='" & pfstatusTemp & "'  and traystaus='" & trayTemp & "' and testno='" & testnoTemp & "' and flag='Y' "
'

cmdStr = "select  Get_ProductName_LotId('" & productNameTemp & "','" & pfstatusTemp & "','" & trayTemp & "','" & testnoTemp & "')   from dual"



RSResult = getStr2(cmdStr)
GetAllPtString = RSResult
End Function

Public Function GetCustomerPtNum(lotIdTemp As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

'cmdStr = " select productname from  TBLSETPT  where productname like '" & productNameTemp & "%' and pfstaus='" & pfstatusTemp & "'  and traystaus='" & trayTemp & "' and testno='" & testnoTemp & "' and flag='Y' "
'

'cmdStr = "select  Get_ProductName_LotId('" & productNameTemp & "','" & pfstatusTemp & "','" & trayTemp & "','" & testnoTemp & "')   from dual"


'cmdStr = "select b.qtechpt from CustomerOItbl_test a,TBLSETQtechPT b  where a.source_batch_id='" & lotidTemp & "' and a.mpn_desc=b.customerpt"
 
cmdStr = "select Get_OpenWO_PT('" & lotIdTemp & "') qtechpt from dual"


RSResult = getStr2(cmdStr)
GetCustomerPtNum = RSResult
End Function

Public Function GetCustomerPtNum37(lotIdTemp As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

 
cmdStr = "select pj_material.po_37_material('" & lotIdTemp & "') qtechpt from dual"


RSResult = getStr2(cmdStr)
GetCustomerPtNum37 = RSResult
End Function


Public Function GetWoDept(lotIdTemp As String) As String
'取工单后四位
Dim cmdStr As String
Dim RSResult As String

'cmdStr = " select productname from  TBLSETPT  where productname like '" & productNameTemp & "%' and pfstaus='" & pfstatusTemp & "'  and traystaus='" & trayTemp & "' and testno='" & testnoTemp & "' and flag='Y' "
'

'cmdStr = "select  Get_ProductName_LotId('" & productNameTemp & "','" & pfstatusTemp & "','" & trayTemp & "','" & testnoTemp & "')   from dual"


'cmdStr = "select b.qtechpt from CustomerOItbl_test a,TBLSETQtechPT b  where a.source_batch_id='" & lotidTemp & "' and a.mpn_desc=b.customerpt"
 
cmdStr = "select Get_Product_Dept('" & lotIdTemp & "') qtechpt from dual"


RSResult = getStr2(cmdStr)
GetWoDept = RSResult
End Function

Public Function GetWoIDTemp(lotIdTemp As String) As String

Dim cmdStr As String
Dim RSResult As String

cmdStr = "select  nvl( max(sequenceid) ,0)+1 from  TSV_WO_SEQ_TAB   where flag='Y' and wotype='" & lotIdTemp & "' and ymonth=to_char(sysdate,'YYMM') "

RSResult = getStr2(cmdStr)
GetWoIDTemp = RSResult
End Function


Public Function GetSeqID() As Long
'取Header Seq
Dim cmdStr As String
Dim RSResult As Long

cmdStr = "select IB_WOHISTORY_TEST_seq.nextval ID from dual"
     
RSResult = GetSeq(cmdStr)
GetSeqID = RSResult
End Function



Public Function GetMPSBillID() As Long
'取Header Seq
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "select IB_WOHISTORY_TEST_seq.nextval ID from dual"

cmdStr = " select count(*)+1 ID  from  CUSTOMERMarkingCode where to_char(qtech_created_date,'YYYY-MM-DD')=  to_char(sysdate,'YYYY-MM-DD') "

     
RSResult = GetSeq(cmdStr)
GetMPSBillID = RSResult
End Function






Public Sub AddBillHeaderWoDummy(dataTemp As BillHeader)
'增加Header
Dim cmdStr As String
Dim cmdStrSql As String

Dim cmdStr2 As String
Dim UpLotId As String

Dim i As Integer
Dim detailTemp As BillDetail
UpLotId = ""

'新增加到Bom领料表中
Dim woTemp As String
Dim qtyWaferTemp As Long
woTemp = dataTemp.ORDERNAME
qtyWaferTemp = 0

Dim woDeptTemp As String



On Error GoTo DealError

woDeptTemp = dataTemp.PARA8

woDeptTemp = Right(woDeptTemp, Len(woDeptTemp) - InStr(woDeptTemp, "_"))

         
Cnn.BeginTrans
         
cmdStr = "insert into ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",to_date('" & dataTemp.ERPCREATEDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANENDDATE & "','yyyy-mm-dd')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & dataTemp.PARA8 & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"

 
 cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",convert(datetime,'" & dataTemp.ERPCREATEDATE & "'),convert(datetime,'" & dataTemp.PLANSTARTDATE & "'),convert(datetime,'" & dataTemp.PLANENDDATE & "')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & woDeptTemp & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"
         
         
'cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
'         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
'         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
'         " Values (" & dataTemp.id & ",'" & dataTemp.OrderName & "','" & dataTemp.orderType & "' ,'CREATED','" & dataTemp.ERPUser & "','" & dataTemp.Product & "'," & dataTemp.Qty & ",convert(datetime,'" & dataTemp.ERPCreateDate & "'),to_date('" & dataTemp.ERPCreateDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanStartDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanEndDate & "','yyyy-mm-dd')," & _
'         " '" & dataTemp.Customer & "','" & dataTemp.SalesOrder & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FabFacility & "','" & dataTemp.ImagerRev & "','" & dataTemp.Designid & "','" & dataTemp.MLevel235 & "','" & dataTemp.Mlevel260 & "','" & dataTemp.NGFlag & "','" & dataTemp.Para1 & "'," & _
'         "  '" & dataTemp.Para2 & "','" & dataTemp.Para3 & "','" & dataTemp.Para4 & "','" & dataTemp.Para5 & "','" & dataTemp.Para6 & "','" & dataTemp.RequestDate & "','" & dataTemp.Protective_Film_Apld & "','" & dataTemp.Lot_Stauts & "'," & _
'         " '" & dataTemp.MPN & "')"
 
 
 
'
' AddSql (cmdStr)
AddSqlERPInt (cmdStr)
  
AddSql2 (cmdStrSql)


'With FrmApplyWO.fps(0)
'
'For i = 1 To .MaxRows
'
'    .Row = i
'    .Col = 8
'    If .Text = 1 Then
'
'    detailTemp.OrderName = FrmApplyWO.Text2.Text
'    .Row = i
'    .Col = 2
'    detailTemp.WaferId = .Text
'
'    .Col = 4
'    detailTemp.DieQty = .Text
'
'    .Col = 5
'    detailTemp.FGDieQty = .Text
'
'    .Col = 6
'    detailTemp.WaferLot = .Text
'
'    If InStr(1, UpLotId, detailTemp.WaferLot) = 0 Then
'        UpLotId = UpLotId & "," & detailTemp.WaferLot
'
'    End If
'
'
'    .Col = 7
'    detailTemp.MarkingCode = .Text
'
'
'   cmdStr2 = "insert into ib_waferlist(ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.OrderName & "'," & _
'             " '" & detailTemp.WaferId & "'," & detailTemp.DieQty & "," & detailTemp.DieQty & ",'" & detailTemp.WaferLot & "',100,'" & detailTemp.MarkingCode & "')"
'
'   cmdStrSql = "insert into [erpdata].[dbo].[tblTSVwaferlist](ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.OrderName & "'," & _
'             " '" & detailTemp.WaferId & "'," & detailTemp.DieQty & "," & detailTemp.DieQty & ",'" & detailTemp.WaferLot & "',100,'" & detailTemp.MarkingCode & "')"
'
'
'
''    AddSql (cmdStr2)
'    AddSqlERPInt (cmdStr2)
'
'
'    AddSql2 (cmdStrSql)
'
'    qtyWaferTemp = qtyWaferTemp + 1
'
'
'    End If
'
'Next i
'
'
'End With


'校验数据
Dim ArrayLot() As String
Dim j As Integer
Dim detailCount As Integer
Dim detailQty As Long

Dim headerCount As Integer
Dim headerQty As Long



'    If InStr(1, UpLotId, ",") > 0 Then
'        ArrayLot = Split(UpLotId, ",")
'
'        For i = 1 To UBound(ArrayLot)
'            '算出明细表中的笔数，及数量;算出Head表中的Wafer数，及数量；
'            '如果都相等，说明这些单据已完成，压上状态，并Update数量
'            '如果<Header表，说明未完成，不压状态，但Update数量
'            If FrmApplyWO.CmbCustomer.Text = "AA" Then
'
'            Set updateRS = GetDetailData(ArrayLot(i))
'            detailCount = CInt(updateRS.fields("num").Value)
'            detailQty = CLng(updateRS.fields("sumQty").Value)
'
'            Set updateRSHeader = GetHeaderData(ArrayLot(i))
'            headerCount = CInt(updateRSHeader.fields("current_wafer_qty").Value)
'            headerQty = CLng(updateRSHeader.fields("die_qty").Value)
'
'            If detailCount = headerCount And detailQty = headerQty Then
'                '关闭Header状态
'               Call updateHeaderDate(CStr(ArrayLot(i)), "N", detailQty)
'
'            ElseIf detailCount < headerCount And detailQty < headerQty Then
'                 '更新数量
'                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
'
'            ElseIf detailCount > headerCount Or detailQty > headerQty Then
'                 '异常
''                 GoTo DealError
'
'            End If
'
'            Else
'                Call updateHeaderDate(CStr(ArrayLot(i)), "N", detailQty)
'
'            End If
'
'
'        Next
'    End If


Dim bomStrTemp As String



'bomStrTemp = "INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',a.qty*" & qtyWaferTemp & " + (a.SHRateQty*a.qty*" & qtyWaferTemp & ")/100 ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a WHERE a.qty>0"
'
'
'
'
'AddSql2 (bomStrTemp)




Cnn.CommitTrans
MsgBox "工单：" & FrmDummy.Text2.text & "建立成功 !", vbInformation, "提示"


Exit Sub

DealError:

Cnn.RollbackTrans


End Sub


Public Sub AddBillHeader(dataTemp As BillHeader)
'增加Header
Dim cmdStr As String
Dim cmdStrSql As String

Dim cmdStr2 As String
Dim UpLotId As String

Dim i As Integer
Dim detailTemp As BillDetail
UpLotId = ""

'新增加到Bom领料表中
Dim woTemp As String
Dim qtyWaferTemp As Long
Dim qtyDieTemp As Long
Dim woDeptTemp As String

woTemp = dataTemp.ORDERNAME
qtyWaferTemp = 0
qtyDieTemp = 0

woDeptTemp = dataTemp.PARA8

woDeptTemp = Right(woDeptTemp, Len(woDeptTemp) - InStr(woDeptTemp, "_"))


On Error GoTo DealError
         
Cnn.BeginTrans
         
cmdStr = "insert into ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",to_date('" & dataTemp.ERPCREATEDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANENDDATE & "','yyyy-mm-dd')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & dataTemp.PARA8 & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"

 
 cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",convert(datetime,'" & dataTemp.ERPCREATEDATE & "'),convert(datetime,'" & dataTemp.PLANSTARTDATE & "'),convert(datetime,'" & dataTemp.PLANENDDATE & "')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & woDeptTemp & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"
         
         
'cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
'         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
'         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
'         " Values (" & dataTemp.id & ",'" & dataTemp.OrderName & "','" & dataTemp.orderType & "' ,'CREATED','" & dataTemp.ERPUser & "','" & dataTemp.Product & "'," & dataTemp.Qty & ",convert(datetime,'" & dataTemp.ERPCreateDate & "'),to_date('" & dataTemp.ERPCreateDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanStartDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanEndDate & "','yyyy-mm-dd')," & _
'         " '" & dataTemp.Customer & "','" & dataTemp.SalesOrder & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FabFacility & "','" & dataTemp.ImagerRev & "','" & dataTemp.Designid & "','" & dataTemp.MLevel235 & "','" & dataTemp.Mlevel260 & "','" & dataTemp.NGFlag & "','" & dataTemp.Para1 & "'," & _
'         "  '" & dataTemp.Para2 & "','" & dataTemp.Para3 & "','" & dataTemp.Para4 & "','" & dataTemp.Para5 & "','" & dataTemp.Para6 & "','" & dataTemp.RequestDate & "','" & dataTemp.Protective_Film_Apld & "','" & dataTemp.Lot_Stauts & "'," & _
'         " '" & dataTemp.MPN & "')"
 
 
 
'
' AddSql (cmdStr)
AddSqlERPInt (cmdStr)

Call addLogTxt(woTemp, " 插入表:ib_workorder ")

  
 AddSql2 (cmdStrSql)
 
 
Call addLogTxt(woTemp, " 插入SqlServer表:tblTSVworkorder ")



With FrmApplyWO2.Fps(0)

For i = 1 To .MaxRows

    .Row = i
    .Col = 8
    If .text = 1 Then

    detailTemp.ORDERNAME = FrmApplyWO2.Text2.text
    .Row = i
    .Col = 2
    detailTemp.waferid = .text
    
    .Col = 4
    detailTemp.DIEQTY = .text
    
    .Col = 5
    detailTemp.FGDIEQTY = .text
    
'    .Col = 6
'    detailTemp.WaferLot = UCase(Trim(FrmApplyWO2.TxtSourceBatchId.Text))

' 2015-03-19 jiayun modify

    .Col = 6
    
    If upLoadWoFile = True Then
       detailTemp.WAFERLOT = .text
    
    Else
    
    detailTemp.WAFERLOT = UCase(Trim(FrmApplyWO2.TxtSourceBatchId.text))
    
    End If

    
    If InStr(1, UpLotId, detailTemp.WAFERLOT) = 0 Then
        UpLotId = UpLotId & "," & detailTemp.WAFERLOT
    
    End If
    
    
    .Col = 7
    detailTemp.MARKINGCODE = .text
    
    
   cmdStr2 = "insert into ib_waferlist(ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"

   cmdStrSql = "insert into [erpdata].[dbo].[tblTSVwaferlist](ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"



'    AddSql (cmdStr2)
    AddSqlERPInt (cmdStr2)
     
    
    AddSql2 (cmdStrSql)
    
    qtyWaferTemp = qtyWaferTemp + 1
 
    
    End If

Next i


End With


'校验数据
Dim ArrayLot() As String
Dim j As Integer
Dim detailCount As Integer
Dim detailQty As Long

Dim headerCount As Integer
Dim headerQty As Long



    If InStr(1, UpLotId, ",") > 0 Then
        ArrayLot = Split(UpLotId, ",")
        
        For i = 1 To UBound(ArrayLot)
            '算出明细表中的笔数，及数量;算出Head表中的Wafer数，及数量；
            '如果都相等，说明这些单据已完成，压上状态，并Update数量
            '如果<Header表，说明未完成，不压状态，但Update数量
            If FrmApplyWO2.CmbCustomer.text = "AA" Then
            
            Set updateRS = GetDetailData(ArrayLot(i))
            detailCount = CInt(updateRS.Fields("num").Value)
            detailQty = CLng(updateRS.Fields("sumQty").Value)
            
            Set updateRSHeader = GetHeaderData(ArrayLot(i))
            headerCount = CInt(updateRSHeader.Fields("current_wafer_qty").Value)
            headerQty = CLng(updateRSHeader.Fields("die_qty").Value)
        
            If detailCount = headerCount And detailQty = headerQty Then
                '关闭Header状态
               Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
                
            ElseIf detailCount < headerCount And detailQty < headerQty Then
                 '更新数量
                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
                 
            ElseIf detailCount > headerCount Or detailQty > headerQty Then
                 '异常
'                 GoTo DealError
            
            End If
            
            Else
                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
            
            End If
        
           
        Next
    End If


Dim bomStrTemp As String

qtyDieTemp = dataTemp.QTY

'bomStrTemp = "INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',a.qty*" & qtyWaferTemp & " + (a.SHRateQty*a.qty*" & qtyWaferTemp & ")/100 ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a WHERE a.qty>0"


Call addLogTxt(woTemp, " 准备 插入SqlServer表:tblllplan " & "料号：" & dataTemp.product)

'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',CAST( ( CAST(a.qty AS DECIMAL(18,8))*erpdata.dbo.Get_TSV_BomQtyNew2(b.序号," & qtyWaferTemp & "," & qtyDieTemp & ") + (a.SHRateQty*CAST(a.qty AS DECIMAL(18,8))*erpdata.dbo.Get_TSV_BomQtyNew2(b.序号," & qtyWaferTemp & "," & qtyDieTemp & "))/100) AS  DECIMAL(18,3)  ) ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a ,[erpdata].[dbo].[TSVtblMRuleData] b  WHERE a.qty>0 and a.employ='" & gUserName & "' AND a.mid=b.材料规范编号 AND a.wlid=b.物料编号 "

'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & _
'" SELECT distinct '" + woTemp + "',X.物料编号,'1','主选材料', " & _
'" CAST( (  CAST(X.用量 AS DECIMAL(18,8)) *erpdata.dbo.Get_TSV_BomQtyNew2(Y.序号," & qtyWaferTemp & "," & qtyDieTemp & ") " & _
'"  + (X.损耗* CAST(X.用量 AS DECIMAL(18,8)) *erpdata.dbo.Get_TSV_BomQtyNew2(Y.序号," & qtyWaferTemp & "," & qtyDieTemp & "))/100) AS  DECIMAL(18,3)) ,1 " & _
'" from (SELECT b.材料规范编号, b.物料编号,b.每只用量 as 用量,b.损耗 " & _
'" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
'" Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & dataTemp.product & "' AND b.料号 NOT LIKE '18%' AND b.料号 NOT LIKE '19%' AND b.料号1 NOT LIKE '18%' AND b.料号1 NOT LIKE '19%' Union " & _
'" SELECT  b.材料规范编号, b.物料编号,b.每只用量 as 用量,b.损耗 " & _
'" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
'" Where a.材料规范编号 = b.材料规范编号 AND a.材料规范编号=( select childBomNameID  from [erpdata].[dbo].[TSVtblBomSetup] c where c.ProductName='" & dataTemp.product & "' and c.Flag='Y'))  X ,[erpdata].[dbo].[TSVtblMRuleData] Y " & _
'" Where X.材料规范编号 = Y.材料规范编号 And X.物料编号 = Y.物料编号 "


'2015-12-07 jiayun add
dataTemp.product = Trim(Replace(Replace(dataTemp.product, Chr(13), ""), Chr(10), ""))

bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & _
" SELECT distinct  '" + woTemp + "',X.物料编号,'1','主选材料', " & _
" CAST( (CAST(X.用量 AS DECIMAL(18,8)) * " & qtyWaferTemp & " ) AS  DECIMAL(18,3))  ,1 " & _
" from ( SELECT b.材料规范编号, b.物料编号,sum(b.每只用量) as 用量 " & _
" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
" Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & dataTemp.product & "' " & _
" group by b.材料规范编号, b.物料编号 )  X "



AddSql2 (bomStrTemp)

Call addLogTxt(woTemp, " 插入SqlServer表:tblllplan OK")


Cnn.CommitTrans

Call addLogTxt(woTemp, " 成功保存工单！ ")

'Call UpdateDataToJD(dataTemp.OrderName, dataTemp.Product)

MsgBox "工单：" & FrmApplyWO2.Text2.text & "建立成功 !", vbInformation, "提示"


Exit Sub

DealError:

Call addLogTxt(woTemp, " 保存工单失败！ ")

Cnn.RollbackTrans


End Sub


Public Sub AddBillHeaderNotToErp(dataTemp As BillHeader)
'增加Header
Dim cmdStr As String
Dim cmdStrSql As String

Dim cmdStr2 As String
Dim UpLotId As String
Dim woDeptTemp As String



Dim i As Integer
Dim detailTemp As BillDetail
UpLotId = ""

'新增加到Bom领料表中
Dim woTemp As String
Dim qtyWaferTemp As Long
Dim qtyDieTemp As Long


woTemp = dataTemp.ORDERNAME
qtyWaferTemp = 0
qtyDieTemp = 0

woDeptTemp = dataTemp.PARA8

woDeptTemp = Right(woDeptTemp, Len(woDeptTemp) - InStr(woDeptTemp, "_"))


On Error GoTo DealError
         
Cnn.BeginTrans
         
cmdStr = "insert into ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA9,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",to_date('" & dataTemp.ERPCREATEDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANENDDATE & "','yyyy-mm-dd')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & dataTemp.PARA8 & "','" & dataTemp.PARA9 & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"

 
 cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA9,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",convert(datetime,'" & dataTemp.ERPCREATEDATE & "'),convert(datetime,'" & dataTemp.PLANSTARTDATE & "'),convert(datetime,'" & dataTemp.PLANENDDATE & "')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & woDeptTemp & "','" & dataTemp.PARA9 & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"
         
         
'cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
'         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
'         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
'         " Values (" & dataTemp.id & ",'" & dataTemp.OrderName & "','" & dataTemp.orderType & "' ,'CREATED','" & dataTemp.ERPUser & "','" & dataTemp.Product & "'," & dataTemp.Qty & ",convert(datetime,'" & dataTemp.ERPCreateDate & "'),to_date('" & dataTemp.ERPCreateDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanStartDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanEndDate & "','yyyy-mm-dd')," & _
'         " '" & dataTemp.Customer & "','" & dataTemp.SalesOrder & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FabFacility & "','" & dataTemp.ImagerRev & "','" & dataTemp.Designid & "','" & dataTemp.MLevel235 & "','" & dataTemp.Mlevel260 & "','" & dataTemp.NGFlag & "','" & dataTemp.Para1 & "'," & _
'         "  '" & dataTemp.Para2 & "','" & dataTemp.Para3 & "','" & dataTemp.Para4 & "','" & dataTemp.Para5 & "','" & dataTemp.Para6 & "','" & dataTemp.RequestDate & "','" & dataTemp.Protective_Film_Apld & "','" & dataTemp.Lot_Stauts & "'," & _
'         " '" & dataTemp.MPN & "')"
 
 
 
'
' AddSql (cmdStr)
AddSqlERPInt (cmdStr)

Call addLogTxt(woTemp, " 插入表:ib_workorder ")


  
 AddSql2 (cmdStrSql)


Call addLogTxt(woTemp, " 插入SqlServer表:tblTSVworkorder ")



With FrmNotToERPApplyWO2.Fps(0)

For i = 1 To .MaxRows

    .Row = i
    .Col = 8
    If .text = 1 Then

    detailTemp.ORDERNAME = FrmNotToERPApplyWO2.Text2.text
    .Row = i
    .Col = 2
    detailTemp.waferid = .text
    
    .Col = 4
    detailTemp.DIEQTY = .text
    
    .Col = 5
    detailTemp.FGDIEQTY = .text
    
    .Col = 6
    '2015-03-19 jiayun modify
    
     detailTemp.WAFERLOT = .text
    
'    If upLoadWoFile = True Then
'        detailTemp.WaferLot = .Text
'    Else
'        detailTemp.WaferLot = UCase(Trim(FrmNotToERPApplyWO2.TxtSourceBatchId.Text))
'    End If
    
   ' detailTemp.WaferLot = UCase(Trim(FrmNotToERPApplyWO2.TxtSourceBatchId.Text))
    
'    detailTemp.WaferLot = .Text
    
    If InStr(1, UpLotId, detailTemp.WAFERLOT) = 0 Then
        UpLotId = UpLotId & "," & detailTemp.WAFERLOT
    
    End If
    
    
    .Col = 7
    detailTemp.MARKINGCODE = .text
    
    
   cmdStr2 = "insert into ib_waferlist(ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"

   cmdStrSql = "insert into [erpdata].[dbo].[tblTSVwaferlist](ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"



'    AddSql (cmdStr2)
    AddSqlERPInt (cmdStr2)
     
    
    AddSql2 (cmdStrSql)
    
    qtyWaferTemp = qtyWaferTemp + 1
 
    
    End If

Next i


End With


'校验数据
Dim ArrayLot() As String
Dim j As Integer
Dim detailCount As Integer
Dim detailQty As Long

Dim headerCount As Integer
Dim headerQty As Long


'    If InStr(1, UpLotId, ",") > 0 Then
'        ArrayLot = Split(UpLotId, ",")
'
'        For i = 1 To UBound(ArrayLot)
'            '算出明细表中的笔数，及数量;算出Head表中的Wafer数，及数量；
'            '如果都相等，说明这些单据已完成，压上状态，并Update数量
'            '如果<Header表，说明未完成，不压状态，但Update数量
'            If FrmNotToERPApplyWO2.CmbCustomer.Text = "AA" Then
'
'            Set updateRS = GetDetailData(ArrayLot(i))
'            detailCount = CInt(updateRS.fields("num").Value)
'            detailQty = CLng(updateRS.fields("sumQty").Value)
'
'            Set updateRSHeader = GetHeaderData(ArrayLot(i))
'            headerCount = CInt(updateRSHeader.fields("current_wafer_qty").Value)
'            headerQty = CLng(updateRSHeader.fields("die_qty").Value)
'
'            If detailCount = headerCount And detailQty = headerQty Then
'                '关闭Header状态
'               Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
'
'            ElseIf detailCount < headerCount And detailQty < headerQty Then
'                 '更新数量
'                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
'
'            ElseIf detailCount > headerCount Or detailQty > headerQty Then
'                 '异常
''                 GoTo DealError
'
'            End If
'
'            Else
'                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
'
'            End If
'
'
'        Next
'    End If

Dim bomStrTemp As String

qtyDieTemp = dataTemp.QTY

'bomStrTemp = "INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',a.qty*" & qtyWaferTemp & " + (a.SHRateQty*a.qty*" & qtyWaferTemp & ")/100 ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a WHERE a.qty>0"



'2013-11-13 jiayun modify
'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',CAST( ( CAST(a.qty AS DECIMAL(18,8))*erpdata.dbo.Get_TSV_BomQtyNew2(b.序号," & qtyWaferTemp & "," & qtyDieTemp & ") + (a.SHRateQty*CAST(a.qty AS DECIMAL(18,8))*erpdata.dbo.Get_TSV_BomQtyNew2(b.序号," & qtyWaferTemp & "," & qtyDieTemp & "))/100) AS  DECIMAL(18,3)  ) ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a ,[erpdata].[dbo].[TSVtblMRuleData] b  WHERE a.qty>0 AND a.mid=b.材料规范编号 AND a.wlid=b.物料编号 "


'Call addLogTxt(woTemp, " 准备 插入SqlServer表:tblllplan " & "料号：" & dataTemp.product)

'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',CAST( ( CAST(a.qty AS DECIMAL(18,8))*erpdata.dbo.Get_TSV_BomQtyNew2(b.序号," & qtyWaferTemp & "," & qtyDieTemp & ") + (a.SHRateQty*CAST(a.qty AS DECIMAL(18,8))*erpdata.dbo.Get_TSV_BomQtyNew2(b.序号," & qtyWaferTemp & "," & qtyDieTemp & "))/100) AS  DECIMAL(18,3)  ) ,1  FROM  [erpdata].[dbo].[TSVtblBillBom2InitData] a ,[erpdata].[dbo].[TSVtblMRuleData] b  WHERE a.qty>0 and a.employ='" & gUserName & "' AND a.mid=b.材料规范编号 AND a.wlid=b.物料编号 "


'AddSql2 (bomStrTemp)

'Call addLogTxt(woTemp, " 插入SqlServer表:tblllplan OK")



Cnn.CommitTrans


'Call UpdateDataToJD(dataTemp.OrderName, dataTemp.Product)

MsgBox "工单：" & FrmNotToERPApplyWO2.Text2.text & "建立成功 !", vbInformation, "提示"

Call addLogTxt(woTemp, " 成功保存工单！ ")



Exit Sub

DealError:

Call addLogTxt(woTemp, " 保存工单失败！ ")

Cnn.RollbackTrans


End Sub



Public Sub AddBillHeaderSplit(dataTemp As BillHeader)
'增加Header
Dim cmdStr As String
Dim cmdStrSql As String

Dim cmdStr2 As String
Dim UpLotId As String

Dim i As Integer
Dim detailTemp As BillDetail
UpLotId = ""

'新增加到Bom领料表中
Dim woTemp As String
Dim qtyWaferTemp As Long
Dim qtyDieTemp As Long
Dim woDeptTemp As String




woTemp = dataTemp.ORDERNAME
qtyWaferTemp = 0
qtyDieTemp = 0

woDeptTemp = dataTemp.PARA8

woDeptTemp = Right(woDeptTemp, Len(woDeptTemp) - InStr(woDeptTemp, "_"))


On Error GoTo DealError
         
Cnn.BeginTrans
         
cmdStr = "insert into ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",to_date('" & dataTemp.ERPCREATEDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANENDDATE & "','yyyy-mm-dd')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & dataTemp.PARA8 & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"

 
 cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",convert(datetime,'" & dataTemp.ERPCREATEDATE & "'),convert(datetime,'" & dataTemp.PLANSTARTDATE & "'),convert(datetime,'" & dataTemp.PLANENDDATE & "')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & woDeptTemp & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"
         
         
'cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
'         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
'         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
'         " Values (" & dataTemp.id & ",'" & dataTemp.OrderName & "','" & dataTemp.orderType & "' ,'CREATED','" & dataTemp.ERPUser & "','" & dataTemp.Product & "'," & dataTemp.Qty & ",convert(datetime,'" & dataTemp.ERPCreateDate & "'),to_date('" & dataTemp.ERPCreateDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanStartDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanEndDate & "','yyyy-mm-dd')," & _
'         " '" & dataTemp.Customer & "','" & dataTemp.SalesOrder & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FabFacility & "','" & dataTemp.ImagerRev & "','" & dataTemp.Designid & "','" & dataTemp.MLevel235 & "','" & dataTemp.Mlevel260 & "','" & dataTemp.NGFlag & "','" & dataTemp.Para1 & "'," & _
'         "  '" & dataTemp.Para2 & "','" & dataTemp.Para3 & "','" & dataTemp.Para4 & "','" & dataTemp.Para5 & "','" & dataTemp.Para6 & "','" & dataTemp.RequestDate & "','" & dataTemp.Protective_Film_Apld & "','" & dataTemp.Lot_Stauts & "'," & _
'         " '" & dataTemp.MPN & "')"
 
 
 
'
' AddSql (cmdStr)
AddSqlERPInt (cmdStr)

Call addLogTxt(woTemp, " 插入表:ib_workorder ")
  
 AddSql2 (cmdStrSql)

Call addLogTxt(woTemp, " 插入SqlServer表:tblTSVworkorder ")

With FrmToERPApplyWO2.Fps(0)

For i = 1 To .MaxRows

    .Row = i
    .Col = 8
    If .text = 1 Then

    detailTemp.ORDERNAME = FrmToERPApplyWO2.Text2.text
    .Row = i
    .Col = 2
    detailTemp.waferid = .text
    
    .Col = 4
    detailTemp.DIEQTY = .text
    
    .Col = 5
    detailTemp.FGDIEQTY = .text
    
    .Col = 6
    
    
     detailTemp.WAFERLOT = .text
     
    
'    If upLoadWoFile = True Then
'     detailTemp.WaferLot = .Text
'    Else
'       detailTemp.WaferLot = UCase(Trim(FrmToERPApplyWO2.TxtSourceBatchId.Text))
'
'    End If
    
    
   ' detailTemp.WaferLot = UCase(Trim(FrmToERPApplyWO2.TxtSourceBatchId.Text))
    
   ' detailTemp.WaferLot = .Text
        
    
    If InStr(1, UpLotId, detailTemp.WAFERLOT) = 0 Then
        UpLotId = UpLotId & "," & detailTemp.WAFERLOT
    
    End If
    
    
    .Col = 7
    detailTemp.MARKINGCODE = .text
    
    
   cmdStr2 = "insert into ib_waferlist(ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"

   cmdStrSql = "insert into [erpdata].[dbo].[tblTSVwaferlist](ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"



'    AddSql (cmdStr2)
    AddSqlERPInt (cmdStr2)
     
    
    AddSql2 (cmdStrSql)
    
    qtyWaferTemp = qtyWaferTemp + 1
 
     Call addLogTxt(woTemp, " 插入表:ib_waferlist,tblTSVwaferlist " & detailTemp.waferid)
    
    End If

Next i


End With


'校验数据
Dim ArrayLot() As String
Dim j As Integer
Dim detailCount As Integer
Dim detailQty As Long

Dim headerCount As Integer
Dim headerQty As Long



    If InStr(1, UpLotId, ",") > 0 Then
        ArrayLot = Split(UpLotId, ",")
        
        For i = 1 To UBound(ArrayLot)
            '算出明细表中的笔数，及数量;算出Head表中的Wafer数，及数量；
            '如果都相等，说明这些单据已完成，压上状态，并Update数量
            '如果<Header表，说明未完成，不压状态，但Update数量
            If FrmToERPApplyWO2.CmbCustomer.text = "AA" Then
            
            Set updateRS = GetDetailData(ArrayLot(i))
            detailCount = CInt(updateRS.Fields("num").Value)
            detailQty = CLng(updateRS.Fields("sumQty").Value)
            
            Set updateRSHeader = GetHeaderData(ArrayLot(i))
            headerCount = CInt(updateRSHeader.Fields("current_wafer_qty").Value)
            headerQty = CLng(updateRSHeader.Fields("die_qty").Value)
        
            If detailCount = headerCount And detailQty = headerQty Then
                '关闭Header状态
               Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
                
            ElseIf detailCount < headerCount And detailQty < headerQty Then
                 '更新数量
                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
                 
            ElseIf detailCount > headerCount Or detailQty > headerQty Then
                 '异常
'                 GoTo DealError
            
            End If
            
            Else
                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
            
            End If
        
           
        Next
    End If


Dim bomStrTemp As String

qtyDieTemp = dataTemp.QTY

'bomStrTemp = "INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',a.qty*" & qtyWaferTemp & " + (a.SHRateQty*a.qty*" & qtyWaferTemp & ")/100 ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a WHERE a.qty>0"



Call addLogTxt(woTemp, " 准备 插入SqlServer表:tblllplan " & "料号：" & dataTemp.product)

'2015-02-11 jiayun modify
'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',CAST( ( CAST(a.qty AS DECIMAL(18,8))*erpdata.dbo.Get_TSV_BomQtyNew2(b.序号," & qtyWaferTemp & "," & qtyDieTemp & ") + (a.SHRateQty*CAST(a.qty AS DECIMAL(18,8))*erpdata.dbo.Get_TSV_BomQtyNew2(b.序号," & qtyWaferTemp & "," & qtyDieTemp & "))/100) AS  DECIMAL(18,3)  ) ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a ,[erpdata].[dbo].[TSVtblMRuleData] b  WHERE a.qty>0  and a.employ='" & gUserName & "' AND a.mid=b.材料规范编号 AND a.wlid=b.物料编号 "



'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',CAST( ( CAST(a.qty AS DECIMAL(18,8))*erpdata.dbo.Get_TSV_BomQtyNew2(b.序号," & qtyWaferTemp & "," & qtyDieTemp & ") + (a.SHRateQty*CAST(a.qty AS DECIMAL(18,8))*erpdata.dbo.Get_TSV_BomQtyNew2(b.序号," & qtyWaferTemp & "," & qtyDieTemp & "))/100) AS  DECIMAL(18,3)  ) ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a ,[erpdata].[dbo].[TSVtblMRuleData] b  WHERE a.qty>0  and a.employ='" & gUserName & "' AND a.mid=b.材料规范编号 AND a.wlid=b.物料编号 "



'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & _
'" SELECT distinct '" + woTemp + "',X.物料编号,'1','主选材料', " & _
'" CAST( (  CAST(X.用量 AS DECIMAL(18,8)) *erpdata.dbo.Get_TSV_BomQtyNew2(Y.序号," & qtyWaferTemp & "," & qtyDieTemp & ") " & _
'"  + (X.损耗* CAST(X.用量 AS DECIMAL(18,8)) *erpdata.dbo.Get_TSV_BomQtyNew2(Y.序号," & qtyWaferTemp & "," & qtyDieTemp & "))/100) AS  DECIMAL(18,3)) ,1 " & _
'" from (SELECT b.材料规范编号, b.物料编号,b.每只用量 as 用量,b.损耗 " & _
'" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
'" Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & dataTemp.product & "' AND b.料号 NOT LIKE '18%' AND b.料号 NOT LIKE '19%' AND b.料号1 NOT LIKE '18%' AND b.料号1 NOT LIKE '19%' Union " & _
'" SELECT  b.材料规范编号, b.物料编号,b.每只用量 as 用量,b.损耗 " & _
'" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
'" Where a.材料规范编号 = b.材料规范编号 AND a.材料规范编号=( select childBomNameID  from [erpdata].[dbo].[TSVtblBomSetup] c where c.ProductName='" & dataTemp.product & "' and c.Flag='Y'))  X ,[erpdata].[dbo].[TSVtblMRuleData] Y " & _
'" Where X.材料规范编号 = Y.材料规范编号 And X.物料编号 = Y.物料编号 "


'2015-12-07 jiayun add

bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & _
" SELECT distinct  '" + woTemp + "',X.物料编号,'1','主选材料', " & _
" CAST( (CAST(X.用量 AS DECIMAL(18,8)) * " & qtyWaferTemp & " ) AS  DECIMAL(18,3))  ,1 " & _
" from ( SELECT b.材料规范编号, b.物料编号,sum(b.每只用量) as 用量 " & _
" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
" Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & dataTemp.product & "' " & _
" group by b.材料规范编号, b.物料编号 )  X "





Call addLogTxt(woTemp, " 插入SqlServer表:tblllplan OK")

AddSql2 (bomStrTemp)


Call addLogTxt(woTemp, " 插入SqlServer表:tblllplan ")

Cnn.CommitTrans

'Call UpdateDataToJD(dataTemp.OrderName, dataTemp.Product)

'Call addLogTxt(woTemp, " 工单抛金碟 ")

MsgBox "工单：" & FrmToERPApplyWO2.Text2.text & "建立成功 !", vbInformation, "提示"

Call addLogTxt(woTemp, " 成功保存工单！ ")

Exit Sub

DealError:

Call addLogTxt(woTemp, " 保存工单失败！ ")

Cnn.RollbackTrans


End Sub

Public Function Check37PO(waferid As String) As Boolean
Dim strSql As String

Check37PO = False

strSql = "select a.po_num from customeroitbl_test a, mappingdatatest b where to_char(a.id) = b.filename and b.substrateid = '" & waferid & "' and a.po_num is not null and a.test_mtrl_desc is not null"

If Get_OracleCnt(strSql) > 0 Then
    Check37PO = True
End If

End Function

Public Function Check37DATECODE(waferid As String) As Boolean
Dim strSql As String

Check37DATECODE = False

strSql = "select * from weight37 where waferid = '" & waferid & "' "

If Get_OracleCnt(strSql) > 0 Then
    Check37DATECODE = True
End If

End Function

Public Function CheckMapping(LOTID As String) As Boolean
Dim strSql As String

CheckMapping = False

strSql = " select lotid,sum(failbincount) from mappingdatatest where lotid = '" & LOTID & "' and substrateid not like '%+%'  group by lotid having sum(failbincount) = 0 "

If Get_OracleCnt(strSql) > 0 Then
    CheckMapping = True
End If

End Function


Public Sub AddBillHeaderReWork(dataTemp As BillHeader)
'增加Header
Dim cmdStr As String
Dim cmdStrSql As String

Dim cmdStr2 As String
Dim UpLotId As String

Dim i As Integer
Dim detailTemp As BillDetail
UpLotId = ""

'新增加到Bom领料表中
Dim woTemp As String
Dim qtyWaferTemp As Long
Dim qtyDieTemp As Long
Dim woDeptTemp As String


woTemp = dataTemp.ORDERNAME
qtyWaferTemp = 0
qtyDieTemp = 0


woDeptTemp = dataTemp.PARA8

woDeptTemp = Right(woDeptTemp, Len(woDeptTemp) - InStr(woDeptTemp, "_"))

On Error GoTo DealError
         
Cnn.BeginTrans
         
cmdStr = "insert into ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",to_date('" & dataTemp.ERPCREATEDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANENDDATE & "','yyyy-mm-dd')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & dataTemp.PARA8 & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"

 
 cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",convert(datetime,'" & dataTemp.ERPCREATEDATE & "'),convert(datetime,'" & dataTemp.PLANSTARTDATE & "'),convert(datetime,'" & dataTemp.PLANENDDATE & "')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & woDeptTemp & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"
         
         
'cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
'         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
'         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
'         " Values (" & dataTemp.id & ",'" & dataTemp.OrderName & "','" & dataTemp.orderType & "' ,'CREATED','" & dataTemp.ERPUser & "','" & dataTemp.Product & "'," & dataTemp.Qty & ",convert(datetime,'" & dataTemp.ERPCreateDate & "'),to_date('" & dataTemp.ERPCreateDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanStartDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanEndDate & "','yyyy-mm-dd')," & _
'         " '" & dataTemp.Customer & "','" & dataTemp.SalesOrder & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FabFacility & "','" & dataTemp.ImagerRev & "','" & dataTemp.Designid & "','" & dataTemp.MLevel235 & "','" & dataTemp.Mlevel260 & "','" & dataTemp.NGFlag & "','" & dataTemp.Para1 & "'," & _
'         "  '" & dataTemp.Para2 & "','" & dataTemp.Para3 & "','" & dataTemp.Para4 & "','" & dataTemp.Para5 & "','" & dataTemp.Para6 & "','" & dataTemp.RequestDate & "','" & dataTemp.Protective_Film_Apld & "','" & dataTemp.Lot_Stauts & "'," & _
'         " '" & dataTemp.MPN & "')"
 
 
 
'
' AddSql (cmdStr)
AddSqlERPInt (cmdStr)
  
 AddSql2 (cmdStrSql)


With FrmApplyWO2.Fps(0)

For i = 1 To .MaxRows

    .Row = i
    .Col = 8
    If .text = 1 Then

    detailTemp.ORDERNAME = UCase(Trim(FrmApplyWO2.Text2.text))
    .Row = i
    .Col = 2
    detailTemp.waferid = .text
    
    .Col = 4
    detailTemp.DIEQTY = .text
    
    .Col = 5
    detailTemp.FGDIEQTY = .text
    
    .Col = 6
    detailTemp.WAFERLOT = .text
    
    If InStr(1, UpLotId, detailTemp.WAFERLOT) = 0 Then
        UpLotId = UpLotId & "," & detailTemp.WAFERLOT
    
    End If
    
    
    .Col = 7
    detailTemp.MARKINGCODE = .text
    
    
   cmdStr2 = "insert into ib_waferlist(ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"

   cmdStrSql = "insert into [erpdata].[dbo].[tblTSVwaferlist](ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"



'    AddSql (cmdStr2)
    AddSqlERPInt (cmdStr2)
     
    
    AddSql2 (cmdStrSql)
    
    qtyWaferTemp = qtyWaferTemp + 1
 
    
    End If

Next i


End With


'校验数据
Dim ArrayLot() As String
Dim j As Integer
Dim detailCount As Integer
Dim detailQty As Long

Dim headerCount As Integer
Dim headerQty As Long



    If InStr(1, UpLotId, ",") > 0 Then
        ArrayLot = Split(UpLotId, ",")
        
        For i = 1 To UBound(ArrayLot)
            '算出明细表中的笔数，及数量;算出Head表中的Wafer数，及数量；
            '如果都相等，说明这些单据已完成，压上状态，并Update数量
            '如果<Header表，说明未完成，不压状态，但Update数量
            If FrmApplyWO2.CmbCustomer.text = "AA" Then
            
            Set updateRS = GetDetailData(ArrayLot(i))
            detailCount = CInt(updateRS.Fields("num").Value)
            detailQty = CLng(updateRS.Fields("sumQty").Value)
            
            Set updateRSHeader = GetHeaderData(ArrayLot(i))
            headerCount = CInt(updateRSHeader.Fields("current_wafer_qty").Value)
            headerQty = CLng(updateRSHeader.Fields("die_qty").Value)
        
            If detailCount = headerCount And detailQty = headerQty Then
                '关闭Header状态
               Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
                
            ElseIf detailCount < headerCount And detailQty < headerQty Then
                 '更新数量
                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
                 
            ElseIf detailCount > headerCount Or detailQty > headerQty Then
                 '异常
'                 GoTo DealError
            
            End If
            
            Else
                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
            
            End If
        
           
        Next
    End If


Dim bomStrTemp As String

qtyDieTemp = dataTemp.QTY

'bomStrTemp = "INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',a.qty*" & qtyWaferTemp & " + (a.SHRateQty*a.qty*" & qtyWaferTemp & ")/100 ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a WHERE a.qty>0"



'2015-02-11 jiayun modify
'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',CAST( ( CAST(a.qty AS DECIMAL(18,8))*" & qtyDieTemp & " + (a.SHRateQty*CAST(a.qty AS DECIMAL(18,8))*" & qtyDieTemp & ")/100) AS  DECIMAL(18,3)  ) ,1  FROM  [erpdata].[dbo].[TSVtblBillBom2InitData] a where a.employ='" & gUserName & "' "

'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',CAST( ( CAST(a.qty AS DECIMAL(18,8))*" & qtyDieTemp & " + (a.SHRateQty*CAST(a.qty AS DECIMAL(18,8))*" & qtyDieTemp & ")/100) AS  DECIMAL(18,3)  ) ,1  FROM  [erpdata].[dbo].[TSVtblBillBom2InitData] a where a.employ='" & gUserName & "' "



'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & _
'" SELECT distinct '" + woTemp + "',a.物料编号,'1','主选材料'," & _
'" CAST( 1.0 * " & qtyDieTemp & " + (0*1.0 * " & qtyDieTemp & " )/100) AS  DECIMAL(18,3)  ) ,1 " & _
'" FROM  [erpdata].[dbo].[tblSmainM2] a where a.料号='" & dataTemp.product & "' "

'2015-12-09 jiayun add

bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & _
" SELECT distinct '" + woTemp + "',a.物料编号,'1','主选材料'," & _
" CAST( 1.0 * " & qtyDieTemp & " + (0*1.0 * " & qtyDieTemp & " )/100) AS  DECIMAL(18,3)  ) ,1 " & _
" FROM  [erpdata].[dbo].[tblSmainM2] a where a.料号='" & dataTemp.product & "' "




AddSql2 (bomStrTemp)


Cnn.CommitTrans

'2013-05-08 jiayun add

'Call UpdateDataToJD(dataTemp.OrderName, dataTemp.Product)

MsgBox "工单：" & UCase(Trim(FrmApplyWO2.Text2.text)) & "建立成功 !", vbInformation, "提示"


Exit Sub

DealError:

Cnn.RollbackTrans


End Sub

Public Sub AddBillHeaderReWorkNotToERP(dataTemp As BillHeader)
'增加Header
Dim cmdStr As String
Dim cmdStrSql As String

Dim cmdStr2 As String
Dim UpLotId As String

Dim i As Integer
Dim detailTemp As BillDetail
UpLotId = ""

'新增加到Bom领料表中
Dim woTemp As String
Dim qtyWaferTemp As Long
Dim qtyDieTemp As Long

Dim woDeptTemp As String


woTemp = dataTemp.ORDERNAME
qtyWaferTemp = 0
qtyDieTemp = 0

woDeptTemp = dataTemp.PARA8

woDeptTemp = Right(woDeptTemp, Len(woDeptTemp) - InStr(woDeptTemp, "_"))


On Error GoTo DealError
         
Cnn.BeginTrans
         
cmdStr = "insert into ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA9,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",to_date('" & dataTemp.ERPCREATEDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANENDDATE & "','yyyy-mm-dd')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & dataTemp.PARA8 & "','" & dataTemp.PARA9 & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"

 
 cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA9,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",convert(datetime,'" & dataTemp.ERPCREATEDATE & "'),convert(datetime,'" & dataTemp.PLANSTARTDATE & "'),convert(datetime,'" & dataTemp.PLANENDDATE & "')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & woDeptTemp & "','" & dataTemp.PARA9 & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"
         
         
'cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
'         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
'         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
'         " Values (" & dataTemp.id & ",'" & dataTemp.OrderName & "','" & dataTemp.orderType & "' ,'CREATED','" & dataTemp.ERPUser & "','" & dataTemp.Product & "'," & dataTemp.Qty & ",convert(datetime,'" & dataTemp.ERPCreateDate & "'),to_date('" & dataTemp.ERPCreateDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanStartDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanEndDate & "','yyyy-mm-dd')," & _
'         " '" & dataTemp.Customer & "','" & dataTemp.SalesOrder & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FabFacility & "','" & dataTemp.ImagerRev & "','" & dataTemp.Designid & "','" & dataTemp.MLevel235 & "','" & dataTemp.Mlevel260 & "','" & dataTemp.NGFlag & "','" & dataTemp.Para1 & "'," & _
'         "  '" & dataTemp.Para2 & "','" & dataTemp.Para3 & "','" & dataTemp.Para4 & "','" & dataTemp.Para5 & "','" & dataTemp.Para6 & "','" & dataTemp.RequestDate & "','" & dataTemp.Protective_Film_Apld & "','" & dataTemp.Lot_Stauts & "'," & _
'         " '" & dataTemp.MPN & "')"
 
 
 
'
' AddSql (cmdStr)
AddSqlERPInt (cmdStr)
  
 AddSql2 (cmdStrSql)


With FrmNotToERPApplyWO2.Fps(0)

For i = 1 To .MaxRows

    .Row = i
    .Col = 8
    If .text = 1 Then

    detailTemp.ORDERNAME = UCase(Trim(FrmNotToERPApplyWO2.Text2.text))
    .Row = i
    .Col = 2
    detailTemp.waferid = .text
    
    .Col = 4
    detailTemp.DIEQTY = .text
    
    .Col = 5
    detailTemp.FGDIEQTY = .text
    
    .Col = 6
    detailTemp.WAFERLOT = .text
    
    If InStr(1, UpLotId, detailTemp.WAFERLOT) = 0 Then
        UpLotId = UpLotId & "," & detailTemp.WAFERLOT
    
    End If
    
    
    .Col = 7
    detailTemp.MARKINGCODE = .text
    
    
   cmdStr2 = "insert into ib_waferlist(ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"

   cmdStrSql = "insert into [erpdata].[dbo].[tblTSVwaferlist](ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"



'    AddSql (cmdStr2)
    AddSqlERPInt (cmdStr2)
     
    
    AddSql2 (cmdStrSql)
    
    qtyWaferTemp = qtyWaferTemp + 1
 
    
    End If

Next i


End With


'校验数据
Dim ArrayLot() As String
Dim j As Integer
Dim detailCount As Integer
Dim detailQty As Long

Dim headerCount As Integer
Dim headerQty As Long



    If InStr(1, UpLotId, ",") > 0 Then
        ArrayLot = Split(UpLotId, ",")
        
        For i = 1 To UBound(ArrayLot)
            '算出明细表中的笔数，及数量;算出Head表中的Wafer数，及数量；
            '如果都相等，说明这些单据已完成，压上状态，并Update数量
            '如果<Header表，说明未完成，不压状态，但Update数量
            If FrmNotToERPApplyWO2.CmbCustomer.text = "AA" Then
            
            Set updateRS = GetDetailData(ArrayLot(i))
            detailCount = CInt(updateRS.Fields("num").Value)
            detailQty = CLng(updateRS.Fields("sumQty").Value)
            
            Set updateRSHeader = GetHeaderData(ArrayLot(i))
            headerCount = CInt(updateRSHeader.Fields("current_wafer_qty").Value)
            headerQty = CLng(updateRSHeader.Fields("die_qty").Value)
        
            If detailCount = headerCount And detailQty = headerQty Then
                '关闭Header状态
               Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
                
            ElseIf detailCount < headerCount And detailQty < headerQty Then
                 '更新数量
                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
                 
            ElseIf detailCount > headerCount Or detailQty > headerQty Then
                 '异常
'                 GoTo DealError
            
            End If
            
            Else
                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
            
            End If
        
           
        Next
    End If


Dim bomStrTemp As String

qtyDieTemp = dataTemp.QTY

'bomStrTemp = "INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',a.qty*" & qtyWaferTemp & " + (a.SHRateQty*a.qty*" & qtyWaferTemp & ")/100 ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a WHERE a.qty>0"




'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',CAST( ( CAST(a.qty AS DECIMAL(18,8))*" & qtyDieTemp & " + (a.SHRateQty*CAST(a.qty AS DECIMAL(18,8))*" & qtyDieTemp & ")/100) AS  DECIMAL(18,3)  ) ,1  FROM  [erpdata].[dbo].[TSVtblBillBom2InitData] a where a.employ='" & gUserName & "'  "


'AddSql2 (bomStrTemp)


Cnn.CommitTrans

'2013-05-08 jiayun add

'Call UpdateDataToJD(dataTemp.OrderName, dataTemp.Product)

MsgBox "工单：" & UCase(Trim(FrmNotToERPApplyWO2.Text2.text)) & "建立成功 !", vbInformation, "提示"


Exit Sub

DealError:

Cnn.RollbackTrans


End Sub

Public Sub AddBillHeaderReWorkSplit(dataTemp As BillHeader)
'增加Header
Dim cmdStr As String
Dim cmdStrSql As String

Dim cmdStr2 As String
Dim UpLotId As String

Dim i As Integer
Dim detailTemp As BillDetail
UpLotId = ""

'新增加到Bom领料表中
Dim woTemp As String
Dim qtyWaferTemp As Long
Dim qtyDieTemp As Long

Dim woDeptTemp As String


woDeptTemp = dataTemp.PARA8

woDeptTemp = Right(woDeptTemp, Len(woDeptTemp) - InStr(woDeptTemp, "_"))


woTemp = dataTemp.ORDERNAME
qtyWaferTemp = 0
qtyDieTemp = 0

'Call addLogTxt(UCase(Trim(Text2.Text)), " 点击保存按钮 ")


On Error GoTo DealError
         
Cnn.BeginTrans
         
cmdStr = "insert into ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",to_date('" & dataTemp.ERPCREATEDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANENDDATE & "','yyyy-mm-dd')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & dataTemp.PARA8 & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"

 
 cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",convert(datetime,'" & dataTemp.ERPCREATEDATE & "'),convert(datetime,'" & dataTemp.PLANSTARTDATE & "'),convert(datetime,'" & dataTemp.PLANENDDATE & "')," & _
         " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
         "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & woDeptTemp & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
         " '" & dataTemp.MPN & "')"
         
         
'cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
'         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
'         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
'         " Values (" & dataTemp.id & ",'" & dataTemp.OrderName & "','" & dataTemp.orderType & "' ,'CREATED','" & dataTemp.ERPUser & "','" & dataTemp.Product & "'," & dataTemp.Qty & ",convert(datetime,'" & dataTemp.ERPCreateDate & "'),to_date('" & dataTemp.ERPCreateDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanStartDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanEndDate & "','yyyy-mm-dd')," & _
'         " '" & dataTemp.Customer & "','" & dataTemp.SalesOrder & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FabFacility & "','" & dataTemp.ImagerRev & "','" & dataTemp.Designid & "','" & dataTemp.MLevel235 & "','" & dataTemp.Mlevel260 & "','" & dataTemp.NGFlag & "','" & dataTemp.Para1 & "'," & _
'         "  '" & dataTemp.Para2 & "','" & dataTemp.Para3 & "','" & dataTemp.Para4 & "','" & dataTemp.Para5 & "','" & dataTemp.Para6 & "','" & dataTemp.RequestDate & "','" & dataTemp.Protective_Film_Apld & "','" & dataTemp.Lot_Stauts & "'," & _
'         " '" & dataTemp.MPN & "')"
 
 
 
'
' AddSql (cmdStr)
AddSqlERPInt (cmdStr)

Call addLogTxt(woTemp, " 插入表:ib_workorder ")
  
 AddSql2 (cmdStrSql)

Call addLogTxt(woTemp, " 插入SqlServer表:tblTSVworkorder ")

With FrmToERPApplyWO2.Fps(0)

For i = 1 To .MaxRows

    .Row = i
    .Col = 8
    If .text = 1 Then

    detailTemp.ORDERNAME = UCase(Trim(FrmToERPApplyWO2.Text2.text))
    .Row = i
    .Col = 2
    detailTemp.waferid = .text
    
    .Col = 4
    detailTemp.DIEQTY = .text
    
    .Col = 5
    detailTemp.FGDIEQTY = .text
    
    .Col = 6
    detailTemp.WAFERLOT = .text
    
    If InStr(1, UpLotId, detailTemp.WAFERLOT) = 0 Then
        UpLotId = UpLotId & "," & detailTemp.WAFERLOT
    
    End If
    
    
    .Col = 7
    detailTemp.MARKINGCODE = .text
    
    
   cmdStr2 = "insert into ib_waferlist(ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"

   cmdStrSql = "insert into [erpdata].[dbo].[tblTSVwaferlist](ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"



'    AddSql (cmdStr2)
    AddSqlERPInt (cmdStr2)
     
    
    AddSql2 (cmdStrSql)
    
    qtyWaferTemp = qtyWaferTemp + 1
    
      Call addLogTxt(woTemp, " 插入表:ib_waferlist,tblTSVwaferlist " & detailTemp.waferid)
 
    
    End If

Next i


End With


'校验数据
Dim ArrayLot() As String
Dim j As Integer
Dim detailCount As Integer
Dim detailQty As Long

Dim headerCount As Integer
Dim headerQty As Long



    If InStr(1, UpLotId, ",") > 0 Then
        ArrayLot = Split(UpLotId, ",")
        
        For i = 1 To UBound(ArrayLot)
            '算出明细表中的笔数，及数量;算出Head表中的Wafer数，及数量；
            '如果都相等，说明这些单据已完成，压上状态，并Update数量
            '如果<Header表，说明未完成，不压状态，但Update数量
            If FrmToERPApplyWO2.CmbCustomer.text = "AA" Then
            
            Set updateRS = GetDetailData(ArrayLot(i))
            detailCount = CInt(updateRS.Fields("num").Value)
            detailQty = CLng(updateRS.Fields("sumQty").Value)
            
            Set updateRSHeader = GetHeaderData(ArrayLot(i))
            headerCount = CInt(updateRSHeader.Fields("current_wafer_qty").Value)
            headerQty = CLng(updateRSHeader.Fields("die_qty").Value)
        
            If detailCount = headerCount And detailQty = headerQty Then
                '关闭Header状态
               Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
                
            ElseIf detailCount < headerCount And detailQty < headerQty Then
                 '更新数量
                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
                 
            ElseIf detailCount > headerCount Or detailQty > headerQty Then
                 '异常
'                 GoTo DealError
            
            End If
            
            Else
                Call updateHeaderDate(CStr(ArrayLot(i)), "Y", detailQty)
            
            End If
        
           
        Next
    End If


Dim bomStrTemp As String

qtyDieTemp = dataTemp.QTY

'bomStrTemp = "INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',a.qty*" & qtyWaferTemp & " + (a.SHRateQty*a.qty*" & qtyWaferTemp & ")/100 ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a WHERE a.qty>0"

Call addLogTxt(woTemp, " 准备 插入SqlServer表:tblllplan " & "料号：" & dataTemp.product)

'2015-02-11 jiayun modify
'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',CAST( ( CAST(a.qty AS DECIMAL(18,8))*" & qtyDieTemp & " + (a.SHRateQty*CAST(a.qty AS DECIMAL(18,8))*" & qtyDieTemp & ")/100) AS  DECIMAL(18,3)  ) ,1  FROM  [erpdata].[dbo].[TSVtblBillBom2InitData] a  where a.employ='" & gUserName & "' "


'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',CAST( ( CAST(a.qty AS DECIMAL(18,8))*" & qtyDieTemp & " + (a.SHRateQty*CAST(a.qty AS DECIMAL(18,8))*" & qtyDieTemp & ")/100) AS  DECIMAL(18,3)  ) ,1  FROM  [erpdata].[dbo].[TSVtblBillBom2InitData] a  where a.employ='" & gUserName & "' "


'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & _
'" SELECT distinct '" + woTemp + "',a.物料编号,'1','主选材料'," & _
'" CAST((1.0 * " & qtyDieTemp & " + (0*1.0 * " & qtyDieTemp & " )/100) AS  DECIMAL(18,3)  ) ,1 " & _
'" FROM  [erpdata].[dbo].[tblSmainM2] a where a.料号='" & dataTemp.product & "' "


'2015-12-07 jiayun add
bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & _
" SELECT distinct '" + woTemp + "',a.物料编号,'1','主选材料'," & _
" CAST((1.0 * " & qtyDieTemp & " + (0*1.0 * " & qtyDieTemp & " )/100) AS  DECIMAL(18,3)  ) ,1 " & _
" FROM  [erpdata].[dbo].[tblSmainM2] a where a.料号='" & dataTemp.product & "' "




AddSql2 (bomStrTemp)


Cnn.CommitTrans

Call addLogTxt(woTemp, " 插入SqlServer表:tblllplan  OK ")


'2013-05-08 jiayun add

'Call UpdateDataToJD(dataTemp.OrderName, dataTemp.Product)

'Call addLogTxt(woTemp, " 工单抛金碟 ")


MsgBox "工单：" & UCase(Trim(FrmToERPApplyWO2.Text2.text)) & "建立成功 !", vbInformation, "提示"


Call addLogTxt(woTemp, " 成功保存工单！ ")

Exit Sub

DealError:
Call addLogTxt(woTemp, " 保存工单失败！ ")

Cnn.RollbackTrans


End Sub



'2013-08-30 jiayun add log 日志
Public Sub addLogTxt(woTemp As String, msgTxt As String)
'判断txt文件是否存在，如不存在，则建立
Dim fileNameTemp As String
Dim dirNameTemp As String
Dim fileTemp As String

dirNameTemp = "C:\TSVWoLog\"
fileNameTemp = woTemp & ".txt"
fileTemp = dirNameTemp & fileNameTemp

Open fileTemp For Append As #1   '文件存在就追加，不存在就自动创建
Print #1, CStr(Now) + msgTxt
Close #1


End Sub

'2014-07-22 jiayun add HY WaferId标签
Public Sub addLabelTxt(filename As String, msgTxt As String, dirtemp As String)
'判断txt文件是否存在，如不存在，则建立
Dim fileNameTemp As String
Dim dirNameTemp As String
Dim fileTemp As String

dirNameTemp = dirtemp
fileNameTemp = Replace(filename, "'", "") & ".txt"
fileTemp = dirNameTemp & fileNameTemp

Open fileTemp For Append As #1   '文件存在就追加，不存在就自动创建
Print #1, msgTxt
Close #1


End Sub

Public Sub AddBillHeaderWo(dataTemp As BillHeader)

    '增加Header
    Dim cmdStr       As String

    Dim cmdStrSql    As String

    Dim cmdStr2      As String

    Dim strSql       As String

    Dim Rst          As New ADODB.Recordset

    Dim i            As Integer

    Dim detailTemp   As BillDetail

    '新增加到Bom领料表中
    Dim woTemp       As String

    Dim qtyWaferTemp As Long

    Dim qtyDieTemp   As Long

    Dim woDeptTemp   As String

    Dim mCodetemp    As String
    
    Dim pLot         As String
    
    pLot = ""

    UpLotId = ""
   
    woTemp = dataTemp.ORDERNAME

    qtyWaferTemp = 0
    qtyDieTemp = 0

    woDeptTemp = dataTemp.PARA8

    woDeptTemp = Right(woDeptTemp, Len(woDeptTemp) - InStr(woDeptTemp, "_"))

    If Len(woDeptTemp) < 3 Then
        woDeptTemp = "ERROR"

    End If

    On Error GoTo DealError
         
    Cnn.BeginTrans
    INIadoCon.BeginTrans
    
    cmdStr = "insert into ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
       " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
       "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA10,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
       " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",sysdate,to_date('" & dataTemp.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANENDDATE & "','yyyy-mm-dd')," & _
       " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
       "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & dataTemp.PARA8 & "','" & dataTemp.PARA10 & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
       " '" & dataTemp.MPN & "')"
 
    cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
       " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
       "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PARA8,PARA10,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
       " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",GetDate(),convert(datetime,'" & dataTemp.PLANSTARTDATE & "'),convert(datetime,'" & dataTemp.PLANENDDATE & "')," & _
       " '" & dataTemp.CUSTOMER & "','" & dataTemp.SALESORDER & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FABFACILITY & "','" & dataTemp.IMAGERREV & "','" & dataTemp.DESIGNID & "','" & dataTemp.MLEVEL235 & "','" & dataTemp.MLEVEL260 & "','" & dataTemp.NGFLAG & "','" & dataTemp.PARA1 & "'," & _
       "  '" & dataTemp.PARA2 & "','" & dataTemp.PARA3 & "','" & dataTemp.PARA4 & "','" & dataTemp.PARA5 & "','" & dataTemp.PARA6 & "','" & dataTemp.RequestDate & "','" & woDeptTemp & "', '" & dataTemp.PARA10 & "','" & dataTemp.PROTECTIVE_FILM_APLD & "','" & dataTemp.Lot_Stauts & "'," & _
       " '" & dataTemp.MPN & "')"
         
    AddSql (cmdStr)

    Call addLogTxt(woTemp, " 插入表:ib_workorder ")
  
    AddSql2 (cmdStrSql)
 
    Call addLogTxt(woTemp, " 插入SqlServer表:tblTSVworkorder ")

    With FrmApplyWO.Fps(0)

        For i = 1 To .MaxRows

            .Row = i
            .Col = 8

            If .text = 1 Then

                detailTemp.ORDERNAME = UCase(Trim(FrmApplyWO.Text2.text))
                .Row = i
                .Col = 2
                detailTemp.waferid = .text
    
                .Col = 4
                detailTemp.DIEQTY = .text
    
                .Col = 5
                detailTemp.FGDIEQTY = .text

                .Col = 6
                detailTemp.WAFERLOT = .text

                If detailTemp.WAFERLOT <> pLot Then
                    Call UpdateLotState(dataTemp.CUSTOMER, detailTemp.WAFERLOT, dataTemp.SALESORDER)
                    pLot = detailTemp.WAFERLOT
                End If
    
                .Col = 7
                If dataTemp.CUSTOMER = "81" Then
                    If InStr(dataTemp.product, "X81001B") Then
                        detailTemp.MARKINGCODE = "HS" & Mid(Year(DATE), 3, 1) & "A" & Mid(Year(DATE), 4, 1) & "S" & DatePart("ww", Now)
                    Else
                        detailTemp.MARKINGCODE = "EHD510"

                    End If

                Else
                    detailTemp.MARKINGCODE = .text

                End If

                cmdStr2 = "insert into ib_waferlist(ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"

                cmdStrSql = "insert into [erpdata].[dbo].[tblTSVwaferlist](ORDERNAME ,WAFERID,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE) values('" & detailTemp.ORDERNAME & "'," & " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & "," & detailTemp.FGDIEQTY & ",'" & detailTemp.WAFERLOT & "',100,'" & detailTemp.MARKINGCODE & "')"

                AddSql (cmdStr2)
     
                AddSql2 (cmdStrSql)

                qtyWaferTemp = qtyWaferTemp + 1
 
                Call addLogTxt(woTemp, " 插入表:ib_waferlist,tblTSVwaferlist " & detailTemp.waferid)
    
            End If

        Next i

    End With

    Dim bomStrTemp As String

    qtyDieTemp = dataTemp.QTY

    Call addLogTxt(woTemp, " 准备 插入SqlServer表:tblllplan " & "料号：" & dataTemp.product)

    If detailTemp.WAFERLOT = "95FPC" And dataTemp.CUSTOMER = "95" Then
        bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & " SELECT distinct  '" + woTemp + "',X.物料编号,'1','主选材料', " & " CAST( (CAST(X.用量 AS DECIMAL(18,8)) * " & qtyDieTemp & " ) AS  DECIMAL(18,3))  ,1 " & " from ( SELECT b.材料规范编号, b.物料编号,sum(b.每只用量) as 用量 " & " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & " Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & dataTemp.product & "' " & " group by b.材料规范编号, b.物料编号 )  X "
    Else
        bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) " & " SELECT distinct  '" + woTemp + "',X.物料编号,'1','主选材料', " & " CAST( (CAST(X.用量 AS DECIMAL(18,8)) * " & qtyWaferTemp & " ) AS  DECIMAL(18,3))  ,1 " & " from ( SELECT b.材料规范编号, b.物料编号,sum(b.每只用量) as 用量 " & " FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & " Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & dataTemp.product & "' " & " group by b.材料规范编号, b.物料编号 )  X "

    End If

    AddSql2 (bomStrTemp)

    Call addLogTxt(woTemp, " 插入SqlServer表:tblllplan OK")
    
    Cnn.CommitTrans
    INIadoCon.CommitTrans
    
    Call addLogTxt(woTemp, " 成功保存工单！ ")
    Exit Sub
    
DealError:
    Call addLogTxt(woTemp, " 保存工单失败！ ")
    Cnn.RollbackTrans
    INIadoCon.RollbackTrans

End Sub

Public Sub UpdateLotState(strCusCode As String, strLotID As String, strWO As String)

    If strCusCode = "AA" Or strCusCode = "AA(ON)" Then
            
        Call ONLotIDClose(strLotID)
    Else
        Call updateHeaderDateForGC(strLotID, "Y", 0, strWO)
            
    End If

End Sub

Public Sub AddSpecialGR(dataTemp As SpGR)
'增加Header
Dim cmdStr As String
Dim cmdStrSql As String

Dim cmdStr2 As String

On Error GoTo DealError
         
Cnn.BeginTrans
         
cmdStr = "insert into SPEGRDETAILHISTORY (PO_NUM,PO_ITEM ,PREVIOUS_BATCH_ID ,PREVIOUS_MTRL_NUM ,BATCH_ID , MTRL_NUM , MTRL_DESC ,MTRL_NUM_MTRLGRP," & _
         " OUTPUT_QTY ,CONSUMED_QTY , REJECT_QTY ,CURRENT_WAFER_QTY ,COUNTRY_OF_ASSEMBLY ,OFFSHORE_ASM_COMPANY,DATE_CODE, ASSEMBLY_FACILITY ," & _
         " COUNTRY_OF_TEST ,OFFSHORE_TEST_COMPANY ,TST_PROGRAM_REV,CREATED_DATE ,CREATED_TIME , FLAG ,CREATE_BY, CREATE_DATE )" & _
         " Values ('" & dataTemp.PoNum & "','" & dataTemp.PoItem & "','" & dataTemp.PoLotID & "' ,'" & dataTemp.PreviousMtrl & "','" & dataTemp.BatchID & "','" & dataTemp.MtrlNum & "','" & dataTemp.MtrlDesc & "','" & dataTemp.MtrlNumMtr & "'," & _
         " 0," & dataTemp.ConsumedQty & "," & dataTemp.RejectQty & "," & dataTemp.CurrentWaferQty & ",'CN','Q_TECH','" & dataTemp.DATECODE & "','OFFSITE ASSEMBLY'," & _
         "  'CN','Q_TECH','" & dataTemp.TstProgram & "','" & dataTemp.CreatedDate & "','" & dataTemp.CreatedTime & "','Y','Auto',sysdate)"
       

 
 cmdStrSql = "insert into [erpdata].[dbo].[SpecialGRdetailHistory] (PO_NUM,PO_ITEM ,PREVIOUS_BATCH_ID ,PREVIOUS_MTRL_NUM ,BATCH_ID , MTRL_NUM , MTRL_DESC ,MTRL_NUM_MTRLGRP," & _
          " OUTPUT_QTY ,CONSUMED_QTY , REJECT_QTY ,CURRENT_WAFER_QTY ,COUNTRY_OF_ASSEMBLY ,OFFSHORE_ASM_COMPANY,DATE_CODE, ASSEMBLY_FACILITY ," & _
          " COUNTRY_OF_TEST ,OFFSHORE_TEST_COMPANY ,TST_PROGRAM_REV,CREATED_DATE ,CREATED_TIME  )" & _
          " Values ('" & dataTemp.PoNum & "','" & dataTemp.PoItem & "','" & dataTemp.PoLotID & "' ,'" & dataTemp.PreviousMtrl & "','" & dataTemp.BatchID & "','" & dataTemp.MtrlNum & "','" & dataTemp.MtrlDesc & "','" & dataTemp.MtrlNumMtr & "'," & _
          " 0," & dataTemp.ConsumedQty & "," & dataTemp.RejectQty & "," & dataTemp.CurrentWaferQty & ",'CN','Q_TECH','" & dataTemp.DATECODE & "','OFFSITE ASSEMBLY'," & _
          "  'CN','Q_TECH','" & dataTemp.TstProgram & "','" & dataTemp.CreatedDate & "','" & dataTemp.CreatedTime & "')"
         
         
AddSql (cmdStr)

AddSql2 (cmdStrSql)
 

Cnn.CommitTrans
MsgBox ("添加成功！")


Exit Sub

DealError:

MsgBox ("添加失败！")

Cnn.RollbackTrans


End Sub


Public Sub ModifySpecialGR(lotIdTemp As String, newQtyPiece As Double, newQtyDie As Integer)
'增加Header
Dim cmdStr As String
Dim cmdStrSql As String

Dim cmdStr2 As String

On Error GoTo DealError
         
Cnn.BeginTrans
         
cmdStr = "Update SPEGRDETAILHISTORY set current_wafer_qty=" & newQtyPiece & ",reject_qty=" & newQtyDie & ",consumed_qty=" & newQtyDie & " where previous_batch_id='" & lotIdTemp & "' and flag='Y' "
cmdStrSql = "Update [erpdata].[dbo].[SpecialGRdetailHistory] set Current_Wafer_Qty=" & newQtyPiece & ",Reject_Qty=" & newQtyDie & ",Consumed_Qty=" & newQtyDie & " where Previous_Batch_ID='" & lotIdTemp & "'"


         
AddSql (cmdStr)

AddSql2 (cmdStrSql)
 

Cnn.CommitTrans
MsgBox ("修改成功！")


Exit Sub

DealError:

MsgBox ("修改失败！")

Cnn.RollbackTrans


End Sub



Public Sub DelSpecialGR(lotIdTemp As String)
'增加Header
Dim cmdStr As String
Dim cmdStrSql As String

Dim cmdStr2 As String

On Error GoTo DealError
         
Cnn.BeginTrans
         
cmdStr = "delete from SPEGRDETAILHISTORY  where previous_batch_id='" & lotIdTemp & "' and flag='Y' "
cmdStrSql = "delete from [erpdata].[dbo].[SpecialGRdetailHistory]  where Previous_Batch_ID='" & lotIdTemp & "'"


         
AddSql (cmdStr)

AddSql2 (cmdStrSql)
 

Cnn.CommitTrans
MsgBox ("删除成功！")


Exit Sub

DealError:

MsgBox ("修改失败！")

Cnn.RollbackTrans


End Sub



Public Function GetTSVWOType(woTypeTemp As String) As String
'查询工单类型，插入到新ERP中
Dim cmdStr As String
Dim RSResult As String

cmdStr = Mid$(woTypeTemp, 2, 1)

If cmdStr = "P" Or cmdStr = "T" Then

RSResult = "量产工单"

ElseIf cmdStr = "E" Or cmdStr = "S" Then
RSResult = "样品工单"

ElseIf cmdStr = "R" Then
RSResult = "重工工单"

Else
RSResult = ""

End If

GetTSVWOType = RSResult
End Function



Public Sub AddWLOBillHeaderWo(dataTemp As WLOBillHeader)
'增加Header
Dim cmdStr As String
Dim cmdStrSql As String

Dim cmdStr2 As String
Dim UpLotId As String
Dim woNoTemp As String


Dim i As Integer
Dim detailTemp As WLOBillDetail
UpLotId = ""

'新增加到Bom领料表中
Dim woTemp As String
Dim qtyWaferTemp As Long
Dim qtyDieTemp As Long


woTemp = dataTemp.ORDERNAME
qtyWaferTemp = 0
qtyDieTemp = 0




On Error GoTo DealError
         
Cnn.BeginTrans
         
cmdStr = "insert into wlo_ib_workorder (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE,para1,customer)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",to_date('" & dataTemp.ERPCREATEDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANSTARTDATE & "','yyyy-mm-dd'),to_date('" & dataTemp.PLANENDDATE & "','yyyy-mm-dd')," & dataTemp.PieceQty & ",'" & dataTemp.CUSTOMER & "')"

 
 cmdStrSql = "insert into [erpdata].[dbo].[tblWLOworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,para1,customer)" & _
         " Values (" & dataTemp.id & ",'" & dataTemp.ORDERNAME & "','" & dataTemp.ORDERTYPE & "' ,'CREATED','" & dataTemp.ERPUSER & "','" & dataTemp.product & "'," & dataTemp.QTY & ",convert(datetime,'" & dataTemp.ERPCREATEDATE & "'),convert(datetime,'" & dataTemp.PLANSTARTDATE & "'),convert(datetime,'" & dataTemp.PLANENDDATE & "')," & dataTemp.PieceQty & ",'" & dataTemp.CUSTOMER & "')"
         
         
'cmdStrSql = "insert into [erpdata].[dbo].[tblTSVworkorder] (SEQ_IBWO ,ORDERNAME , ORDERTYPE ,EVENTTYPE ,ERPUSER ,PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ," & _
'         " CUSTOMER ,SALESORDER,CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1, " & _
'         "PARA2,PARA3,PARA4,PARA5, PARA6,PARA7,PROTECTIVE_FILM_APLD ,LOT_STATUS ,MPN)" & _
'         " Values (" & dataTemp.id & ",'" & dataTemp.OrderName & "','" & dataTemp.orderType & "' ,'CREATED','" & dataTemp.ERPUser & "','" & dataTemp.Product & "'," & dataTemp.Qty & ",convert(datetime,'" & dataTemp.ERPCreateDate & "'),to_date('" & dataTemp.ERPCreateDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanStartDate & "','yyyy-mm-dd'),to_date('" & dataTemp.PlanEndDate & "','yyyy-mm-dd')," & _
'         " '" & dataTemp.Customer & "','" & dataTemp.SalesOrder & "','" & dataTemp.CustomerERPN & "','" & dataTemp.FabFacility & "','" & dataTemp.ImagerRev & "','" & dataTemp.Designid & "','" & dataTemp.MLevel235 & "','" & dataTemp.Mlevel260 & "','" & dataTemp.NGFlag & "','" & dataTemp.Para1 & "'," & _
'         "  '" & dataTemp.Para2 & "','" & dataTemp.Para3 & "','" & dataTemp.Para4 & "','" & dataTemp.Para5 & "','" & dataTemp.Para6 & "','" & dataTemp.RequestDate & "','" & dataTemp.Protective_Film_Apld & "','" & dataTemp.Lot_Stauts & "'," & _
'         " '" & dataTemp.MPN & "')"
 
 
 
'
' AddSql (cmdStr)
AddSqlERPInt (cmdStr)
  
 AddSql2 (cmdStrSql)

''2013-05-08 jiayun add 工单数量抛金碟
'
''-----begin------
'
' Set adoCmd = New ADODB.Command
' Set adoCmd.ActiveConnection = INIadoCon
'     adoCmd.CommandText = "uspPMC_XDInterface"
'     adoCmd.Parameters.Refresh
'     adoCmd.CommandType = adCmdStoredProc
'
'  Set adoprm1 = New ADODB.Parameter   '工单号
'  adoprm1.Type = adChar
'  adoprm1.Size = 20
'  adoprm1.Direction = adParamInput
'  adoprm1.Value = dataTemp.OrderName
'  adoCmd.Parameters.Append adoprm1
'
'  Set adoprm2 = New ADODB.Parameter   '料号
'  adoprm2.Type = adChar
'  adoprm2.Size = 20
'  adoprm2.Direction = adParamInput
'  adoprm2.Value = dataTemp.Product
'  adoCmd.Parameters.Append adoprm2
'
'  Set adoprm3 = New ADODB.Parameter   '数量
'  adoprm3.Type = adInteger
'  adoprm3.Direction = adParamInput
'  adoprm3.Value = dataTemp.Qty
'  adoCmd.Parameters.Append adoprm3
'
'  adoCmd.Execute
'
''----end-------


With FrmWLOApplyWO.Fps(0)

For i = 1 To .MaxRows

    detailTemp.ORDERNAME = UCase(Trim(FrmWLOApplyWO.Text2.text))
    .Row = i
    .Col = 2
    detailTemp.waferid = .text
    
    .Col = 3
    detailTemp.DIEQTY = .text
    

    
    
   cmdStr2 = "insert into WLO_IB_WAFERLIST(ORDERNAME ,WAFERID,DIEQTY,WAFERSEQUENCE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & ",100)"

   cmdStrSql = "insert into [erpdata].[dbo].[tblWLOwaferlist](ORDERNAME ,WAFERID,DIEQTY,WAFERSEQUENCE) values('" & detailTemp.ORDERNAME & "'," & _
             " '" & detailTemp.waferid & "'," & detailTemp.DIEQTY & ",100)"



'    AddSql (cmdStr2)
    AddSqlERPInt (cmdStr2)
     
    
    AddSql2 (cmdStrSql)
    
    qtyWaferTemp = qtyWaferTemp + 1
 

Next i


End With

Dim bomStrTemp As String
'2012-12-14 jiayun add

'qtyDieTemp 总Die数
'qtyWaferTemp 总片数


qtyDieTemp = dataTemp.QTY

'bomStrTemp = "INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',a.qty*" & qtyWaferTemp & " + (a.SHRateQty*a.qty*" & qtyWaferTemp & ")/100 ,1  FROM  [erpdata].[dbo].[TSVtblBillBomInitData] a WHERE a.qty>0"


'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1',b.材料类型,CAST( (  CAST(a.qty AS DECIMAL(18,8)) *" & qtyWaferTemp & " + (a.SHRateQty* CAST(a.qty AS DECIMAL(18,8)) *" & qtyWaferTemp & ")/100) AS  DECIMAL(18,3)  ) ,2  FROM  [erpdata].[dbo].[WLOtblBillBomInitData] a ,[erpdata].[dbo].[TSVtblMRuleData] b  WHERE a.qty>0 AND a.mid=b.材料规范编号 AND a.wlid=b.物料编号 "

'2013-05-31 modify sql

'bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1',b.材料类型,CAST( (  CAST(a.qty AS DECIMAL(18,8)) *" & qtyWaferTemp & " + (a.SHRateQty* CAST(a.qty AS DECIMAL(18,8)) *" & qtyWaferTemp & ")/100) AS  DECIMAL(18,3)  ) ,2  FROM  [erpdata].[dbo].[WLOtblBillBomInitData] a ,[erpdata].[dbo].[TSVtblMRuleData] b  WHERE a.qty>0 AND a.mid=b.材料规范编号 AND a.wlid=b.物料编号 " & _
'             " Union All  SELECT distinct 'WLOTEST02',a.wlid,'1','主选材料',CAST( (  CAST(a.qty AS DECIMAL(18,8)) *20 + (a.SHRateQty* CAST(a.qty AS DECIMAL(18,8)) *20)/100) AS  DECIMAL(18,3)  ) ,2 " & _
'             " FROM  [erpdata].[dbo].[WLOtblBillBomInitData] a  Where a.Qty > 0 And a.Mid Is Null "



'2013-06-04 jiayun

bomStrTemp = " INSERT INTO  [erpbase].[dbo].[tblllplan] (工单号,物料编号, 序组, 材料,用量,产线标记) SELECT distinct '" + woTemp + "',a.wlid,'1',b.材料类型,erpdata.dbo.Get_WLO_BomPTQty(a.wlid ,CAST( (  CAST(a.qty AS DECIMAL(18,8)) *" & qtyWaferTemp & " + (a.SHRateQty* CAST(a.qty AS DECIMAL(18,8)) *" & qtyWaferTemp & ")/100) AS  DECIMAL(18,3)  )) ,2  FROM  [erpdata].[dbo].[WLOtblBillBomInitData] a ,[erpdata].[dbo].[TSVtblMRuleData] b  WHERE a.qty>0 AND a.mid=b.材料规范编号 AND a.wlid=b.物料编号 " & _
             " Union All  SELECT distinct '" + woTemp + "',a.wlid,'1','主选材料',erpdata.dbo.Get_WLO_BomPTQty(a.wlid,CAST( (  CAST(a.qty AS DECIMAL(18,8)) *20 + (a.SHRateQty* CAST(a.qty AS DECIMAL(18,8)) *20)/100) AS  DECIMAL(18,3)  )) ,2 " & _
             " FROM  [erpdata].[dbo].[WLOtblBillBomInitData] a  Where a.Qty > 0 And a.Mid Is Null "




AddSql2 (bomStrTemp)




Cnn.CommitTrans


'2013-07-05 jiayun add 把数据抛到金碟

Call UpdateWLODataToJD(dataTemp.ORDERNAME, dataTemp.product)


MsgBox "工单：" & UCase(Trim(FrmWLOApplyWO.Text2.text)) & "建立成功 !", vbInformation, "提示"


Exit Sub

DealError:

Cnn.RollbackTrans


End Sub


Public Sub DelWLOBillHeaderWo(woTemp As String)

Dim cmdStr As String
Dim cmdStrSql As String
Dim cmdStrSql3 As String


Dim cmdStr2 As String
Dim UpLotId As String
Dim woNoTemp As String


On Error GoTo DealError
         
Cnn.BeginTrans


cmdStr = " delete from  wlo_ib_workorder where ORDERNAME= '" & woTemp & "' "
cmdStrSql = " delete from  [erpdata].[dbo].[tblWLOworkorder] where ORDERNAME= '" & woTemp & "' "
         
AddSqlERPInt (cmdStr)
  
AddSql2 (cmdStrSql)


'删除明细
            
cmdStr2 = " delete from  WLO_IB_WAFERLIST where ORDERNAME= '" & woTemp & "' "
             
cmdStrSql = " delete from  [erpdata].[dbo].[tblWLOwaferlist]  where ORDERNAME= '" & woTemp & "' "

AddSqlERPInt (cmdStr2)
     
AddSql2 (cmdStrSql)


'2014-03-03 jiayun add 删除新ERP bom 的数据

cmdStrSql3 = " delete from  [erpbase].[dbo].[tblllplan] where 产线标记=2 and 工单号= '" & woTemp & "' "

AddSql2 (cmdStrSql3)



'2013-07-05 jiayun add 删除金碟中的数据

cmdStrSql3 = " delete from  from  AIS20141114094336..cbInQty where FBillNo='" & woTemp & "' "

AddSql2 (cmdStrSql3)


Cnn.CommitTrans


MsgBox "工单：" & woTemp & " 已成功删除 !", vbInformation, "提示"


Exit Sub

DealError:

Cnn.RollbackTrans


End Sub



Private Sub UpdateDataToJD(woTemp As String, productTemp As String)

Dim lotIdTemp As String
Dim qtyTemp As Long
Dim erpdate As String

'2013-07-05 jiayun 存储过程加一个生产线别参数
productTemp = Trim(Replace(Replace(productTemp, Chr(13), ""), Chr(10), ""))
Set billLotTemp = GetBillLot(woTemp)
If (billLotTemp.RecordCount > 0) Then

    Do While Not billLotTemp.EOF
        lotIdTemp = billLotTemp.Fields("waferlot").Value
        qtyTemp = CLng(billLotTemp.Fields("qty").Value)
        erpdate = Format(CDate(billLotTemp.Fields("erpcreationdate").Value), "YYYY-MM-DD")
                
          '-----begin------
          
            Set adoCmd = New ADODB.Command
            Set adoCmd.ActiveConnection = INIadoCon2
            adoCmd.CommandText = "uspPMC_XDInterface"
            adoCmd.Parameters.Refresh
            adoCmd.CommandType = adCmdStoredProc
        

          Set adoprm1 = New ADODB.Parameter   '工单号
          adoprm1.type = adChar
          adoprm1.Size = 20
          adoprm1.Direction = adParamInput
          adoprm1.Value = lotIdTemp
          adoCmd.Parameters.Append adoprm1
          
          
          Set adoprm2 = New ADODB.Parameter   '料号
          adoprm2.type = adChar
          adoprm2.Size = 20
          adoprm2.Direction = adParamInput
          adoprm2.Value = productTemp
          adoCmd.Parameters.Append adoprm2
          
          
          Set adoPrm3 = New ADODB.Parameter   '数量
          adoPrm3.type = adInteger
          adoPrm3.Direction = adParamInput
          adoPrm3.Value = qtyTemp
          adoCmd.Parameters.Append adoPrm3
          

          Set adoPrm4 = New ADODB.Parameter   '日期
          adoPrm4.type = adChar
          adoPrm4.Size = 20
          adoPrm4.Direction = adParamInput
          adoPrm4.Value = erpdate
          adoCmd.Parameters.Append adoPrm4
          
         Set adoPrm5 = New ADODB.Parameter   '线别
          adoPrm5.type = adInteger
          adoPrm5.Direction = adParamInput
          adoPrm5.Value = 1
          adoCmd.Parameters.Append adoPrm5
          
          
          
        
          adoCmd.Execute

        
        billLotTemp.MoveNext
   
    Loop
    
End If



End Sub


Public Sub DelERPQboxData(qboxtemp As String)

            Set adoCmd = New ADODB.Command
            Set adoCmd.ActiveConnection = INIadoCon2
            adoCmd.CommandText = "usp_TSV_DelQboxToERP"
            adoCmd.Parameters.Refresh
            adoCmd.CommandType = adCmdStoredProc
        
          Set adoprm1 = New ADODB.Parameter   '箱号
          adoprm1.type = adChar
          adoprm1.Size = 20
          adoprm1.Direction = adParamInput
          adoprm1.Value = qboxtemp
          adoCmd.Parameters.Append adoprm1
          
          adoCmd.Execute

        
End Sub




Private Sub UpdateWLODataToJD(woTemp As String, productTemp As String)

Dim lotIdTemp As String
Dim qtyTemp As Long
Dim erpdate As String

'2013-07-05 jiayun 存储过程加一个生产线别参数
productTemp = Trim(Replace(Replace(productTemp, Chr(13), ""), Chr(10), ""))
Set billLotTemp = GetWLOBillLot(woTemp)
If (billLotTemp.RecordCount > 0) Then

    Do While Not billLotTemp.EOF
        lotIdTemp = billLotTemp.Fields("waferlot").Value
        qtyTemp = CLng(billLotTemp.Fields("qty").Value)
        erpdate = Format(CDate(billLotTemp.Fields("erpcreationdate").Value), "YYYY-MM-DD")
                
          '-----begin------
          
            Set adoCmd = New ADODB.Command
            Set adoCmd.ActiveConnection = INIadoCon2
            adoCmd.CommandText = "uspPMC_XDInterface"
            adoCmd.Parameters.Refresh
            adoCmd.CommandType = adCmdStoredProc
        

          Set adoprm1 = New ADODB.Parameter   '工单号
          adoprm1.type = adChar
          adoprm1.Size = 20
          adoprm1.Direction = adParamInput
          adoprm1.Value = lotIdTemp
          adoCmd.Parameters.Append adoprm1
          
          
          Set adoprm2 = New ADODB.Parameter   '料号
          adoprm2.type = adChar
          adoprm2.Size = 20
          adoprm2.Direction = adParamInput
          adoprm2.Value = productTemp
          adoCmd.Parameters.Append adoprm2
          
          
          Set adoPrm3 = New ADODB.Parameter   '数量
          adoPrm3.type = adInteger
          adoPrm3.Direction = adParamInput
          adoPrm3.Value = qtyTemp
          adoCmd.Parameters.Append adoPrm3
          

          Set adoPrm4 = New ADODB.Parameter   '日期
          adoPrm4.type = adChar
          adoPrm4.Size = 20
          adoPrm4.Direction = adParamInput
          adoPrm4.Value = erpdate
          adoCmd.Parameters.Append adoPrm4
          
         Set adoPrm5 = New ADODB.Parameter   '线别
          adoPrm5.type = adInteger
          adoPrm5.Direction = adParamInput
          adoPrm5.Value = 2
          adoCmd.Parameters.Append adoPrm5
          
          
          
        
          adoCmd.Execute

        
        billLotTemp.MoveNext
   
    Loop
    
End If



End Sub



Public Function GetBillLot(woTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
      
cmdStr = "select waferlot,sum(dieqty) qty  ,to_char(sysdate,'YYYY-MM-DD') erpcreationdate from ib_waferlist where ordername='" + woTemp + "' and wafersequence<10000 group by waferlot order by waferlot"

         
Set RSResult = getStr(cmdStr)
Set GetBillLot = RSResult
End Function


Public Function GetWLO_ViewData(woTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
      
cmdStr = "select ordername,qty,fitemid,deptid,sysdate erpcreatedate from  Vw_WLO_Wo where ORDERNAME='" + woTemp + "' "
         
Set RSResult = getSqlStr(cmdStr)
Set GetWLO_ViewData = RSResult
End Function

Public Function GetTSV_ViewData(woTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
      
'cmdStr = "select mfg.mfgordername,b.productname, sysdate erpcreationdate from mfgorder mfg ,productbase b where mfg.mfgordername='" + woTemp + "'  and b.productbaseid=mfg.productbaseid"

cmdStr = "select a.ordername as mfgordername,a.product as  productname , a.erpcreatedate  as erpcreationdate from ib_wohistory a where a.ordername in ('" + woTemp + "') "
          
Set RSResult = getStr(cmdStr)
Set GetTSV_ViewData = RSResult
End Function



Public Function GetWLOBillLot(woTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
      

cmdStr = "select ordername as waferlot,sum(dieQty) qty, to_char(sysdate,'YYYY-MM-DD')  erpcreationdate from  erpintegration2.WLO_IB_WAFERLIST where ordername='" + woTemp + "' group by ordername "

         
Set RSResult = getStr(cmdStr)
Set GetWLOBillLot = RSResult
End Function



Public Function GetBillLot2() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
      
'cmdStr = "select waferlot,productname,qty from  jiayun_0509 order by  waferlot,productname "

cmdStr = " select   wafernumber as waferlot, productname,qty, erpcreationdate ,deptid  from  jiayun_003temp4  order by wafernumber, productname,erpcreationdate "


         
Set RSResult = getStr(cmdStr)
Set GetBillLot2 = RSResult
End Function




Public Sub DelMesWO(woTemp As String)
'删除Mes数据

'先删中间表数据
Dim strTemp As String
Dim strTemp2 As String

Dim strMesTemp As String
Dim strMesTemp2 As String
Dim strMesTemp3 As String

ConOracle


strTemp = " delete  from   ib_wohistory  where ordername='" & woTemp & "' "

strTemp2 = " delete  from ib_waferlist  where ordername='" & woTemp & "' "

On Error GoTo MidTableError



CnnERPInt.BeginTrans

AddSqlERPInt (strTemp)
AddSqlERPInt (strTemp2)

CnnERPInt.CommitTrans


Cnn.BeginTrans

'再删Mes表数据


strMesTemp = " delete from container conn where conn.mfgorderid = (select mfg.mfgorderid from mfgorder mfg where mfg.mfgordername = '" + woTemp + "' ) "

strMesTemp2 = " delete from mfgorder mfg where mfg.mfgordername='" + woTemp + "' "

strMesTemp3 = " delete from A_Lotwafers al where al.workordername='" + woTemp + "' "

AddSql (strMesTemp)
AddSql (strMesTemp2)
AddSql (strMesTemp3)

Cnn.CommitTrans

MsgBox "工单：" & woTemp & "删除成功 !", vbInformation, "提示"



Exit Sub
MidTableError:

CnnERPInt.RollbackTrans



End Sub


Public Sub DoCloseWo(woTemp As String)
'关闭SqlServer 工单
Dim strTemp As String

Dim FCostOBJID As String

Dim FDeptID As Integer
Dim qtyTemp As Long
Dim erpdate As String

Dim strTemp2 As String


strTemp = " Update [erpbase].[dbo].[tblllplan] Set 完工标记 = 2 WHERE (产线标记=1  or 产线标记=2) AND 完工标记=0 AND 工单号='" + woTemp + "' "

strTemp2 = " insert into [erpbase].[dbo].[tblTSVCloseWoHis](WO,CloseDate,username) values ('" + woTemp + "',getdate(),'" + woTemp + "') "

 
 
 
'2013-07-15 add
If FrmCloseWO.Combo1.text = "WLO" Then


Set billLotTemp = GetWLO_ViewData(woTemp)
If (billLotTemp.RecordCount > 0) Then

    Do While Not billLotTemp.EOF
    
      FCostOBJID = billLotTemp.Fields("fitemid").Value
      FDeptID = CInt(billLotTemp.Fields("deptid").Value)
      qtyTemp = CLng(billLotTemp.Fields("qty").Value)
      erpdate = Format(CDate(billLotTemp.Fields("erpcreatedate").Value), "YYYY-MM-DD")
      
                
          '-----begin------
          
            Set adoCmd = New ADODB.Command
            Set adoCmd.ActiveConnection = INIadoCon2
            adoCmd.CommandText = "uspRework_Interface"
            adoCmd.Parameters.Refresh
            adoCmd.CommandType = adCmdStoredProc
        

          Set adoprm1 = New ADODB.Parameter   '工单号
          adoprm1.type = adChar
          adoprm1.Size = 20
          adoprm1.Direction = adParamInput
          adoprm1.Value = woTemp
          adoCmd.Parameters.Append adoprm1
          
          
          Set adoprm2 = New ADODB.Parameter   '成本料号
          adoprm2.type = adChar
          adoprm2.Size = 20
          adoprm2.Direction = adParamInput
          adoprm2.Value = FCostOBJID
          adoCmd.Parameters.Append adoprm2
          
        Set adoPrm3 = New ADODB.Parameter   '部门
          adoPrm3.type = adInteger
          adoPrm3.Direction = adParamInput
          adoPrm3.Value = FDeptID
          adoCmd.Parameters.Append adoPrm3
          
          
          Set adoPrm4 = New ADODB.Parameter   '数量
          adoPrm4.type = adInteger
          adoPrm4.Direction = adParamInput
          adoPrm4.Value = qtyTemp
          adoCmd.Parameters.Append adoPrm4
          

          Set adoPrm5 = New ADODB.Parameter   '日期
          adoPrm5.type = adChar
          adoPrm5.Size = 20
          adoPrm5.Direction = adParamInput
          adoPrm5.Value = erpdate
          adoCmd.Parameters.Append adoPrm5
          
         Set adoPrm6 = New ADODB.Parameter   '是否可修复
          adoPrm6.type = adInteger
          adoPrm6.Direction = adParamInput
          adoPrm6.Value = 1
          adoCmd.Parameters.Append adoPrm6
          
          
        
          adoCmd.Execute

        
        billLotTemp.MoveNext
   
    Loop
    
End If





End If


 
 '2013-07-24 add tsv
 
 Dim qtyAllTemp As Long
 Dim toINVqtyTemp As Long
 Dim productNameTemp As String
 
 If FrmCloseWO.Combo1.text = "TSV" Then

'查询这笔工单PMC开工单的数量
Set TSVQty1 = GetTSVXiaDanQty(woTemp)
qtyAllTemp = CLng(TSVQty1.Fields("qtyall").Value)

'查询抛新ERP数量

Set TSVQty2 = GetTSVToINVQty(woTemp)
toINVqtyTemp = CLng(TSVQty2.Fields("qtyinv").Value)

If qtyAllTemp - toINVqtyTemp > 0 Then

    Set billLotTemp = GetTSV_ViewData(woTemp)
    If (billLotTemp.RecordCount > 0) Then
    
        Do While Not billLotTemp.EOF
          qtyTemp = qtyAllTemp - toINVqtyTemp
          productNameTemp = billLotTemp.Fields("productname").Value
          erpdate = Format(CDate(billLotTemp.Fields("erpcreationdate").Value), "YYYY-MM-DD")
          FCostOBJID = GetFItemID(productNameTemp)
          FDeptID = 350
          
          billLotTemp.MoveNext
       
        Loop

          
                    
              '-----begin------
              
                Set adoCmd = New ADODB.Command
                Set adoCmd.ActiveConnection = INIadoCon2
                adoCmd.CommandText = "uspRework_Interface"
                adoCmd.Parameters.Refresh
                adoCmd.CommandType = adCmdStoredProc
            
    
              Set adoprm1 = New ADODB.Parameter   '工单号
              adoprm1.type = adChar
              adoprm1.Size = 20
              adoprm1.Direction = adParamInput
              adoprm1.Value = woTemp
              adoCmd.Parameters.Append adoprm1
              
              
              Set adoprm2 = New ADODB.Parameter   '成本料号
              adoprm2.type = adChar
              adoprm2.Size = 20
              adoprm2.Direction = adParamInput
              adoprm2.Value = FCostOBJID
              adoCmd.Parameters.Append adoprm2
              
            Set adoPrm3 = New ADODB.Parameter   '部门
              adoPrm3.type = adInteger
              adoPrm3.Direction = adParamInput
              adoPrm3.Value = FDeptID
              adoCmd.Parameters.Append adoPrm3
              
              
              Set adoPrm4 = New ADODB.Parameter   '数量
              adoPrm4.type = adInteger
              adoPrm4.Direction = adParamInput
              adoPrm4.Value = qtyTemp
              adoCmd.Parameters.Append adoPrm4
              
    
              Set adoPrm5 = New ADODB.Parameter   '日期
              adoPrm5.type = adChar
              adoPrm5.Size = 20
              adoPrm5.Direction = adParamInput
              adoPrm5.Value = erpdate
              adoCmd.Parameters.Append adoPrm5
              
             Set adoPrm6 = New ADODB.Parameter   '是否可修复
              adoPrm6.type = adInteger
              adoPrm6.Direction = adParamInput
              adoPrm6.Value = 1
              adoCmd.Parameters.Append adoPrm6
              
              
            
              adoCmd.Execute
    
            
       
        
    End If
    

End If


End If



 

AddSql2 (strTemp)

MsgBox "工单：" & woTemp & "关闭成功 !", vbInformation, "提示"

End Sub


Public Sub DoModifyWoQty(woTemp As String, mtID As String, beforQty As Double, afterQty As Double, Userid As String)
'关闭SqlServer 工单
Dim strTemp As String

Dim FCostOBJID As String

Dim FDeptID As Integer
Dim qtyTemp As Long
Dim erpdate As String

Dim strTemp2 As String


strTemp = " Update [erpbase].[dbo].[tblllplan] Set 用量 = " & afterQty & "  WHERE 产线标记=1  AND 工单号='" & woTemp & "'  and 物料编号 = '" & mtID & "'  "

strTemp2 = " insert into [erpbase].[dbo].[tblTSVModifyWoBom](WO,PTID,BeforQty,AfterQty,username,ModifyDate) values ('" & woTemp & "','" & mtID & "'," & beforQty & "," & afterQty & ",'" & Userid & "',getdate()) "

 
AddSql2 (strTemp)
 

AddSql2 (strTemp)

'MsgBox "工单：" & woTemp & "Bom用量修改成功 !", vbInformation, "提示"

End Sub




Public Sub DoCloseWoNew(woTemp As String, Userid As String)
'关闭SqlServer 工单
Dim strTemp As String

Dim FCostOBJID As String

Dim FDeptID As Integer
Dim qtyTemp As Long
Dim erpdate As String

Dim strTemp2 As String


strTemp = " Update [erpbase].[dbo].[tblllplan] Set 完工标记 = 2 WHERE (产线标记=1  or 产线标记=2) AND 完工标记=0 AND 工单号='" + woTemp + "' "

strTemp2 = " insert into [erpbase].[dbo].[tblTSVCloseWoHis](WO,CloseDate,username) values ('" + woTemp + "',getdate(),'" + Userid + "') "

 
 
 
'2013-07-15 add
If FrmCloseWO.Combo1.text = "WLO" Then


Set billLotTemp = GetWLO_ViewData(woTemp)
If (billLotTemp.RecordCount > 0) Then

    Do While Not billLotTemp.EOF
    
      FCostOBJID = billLotTemp.Fields("fitemid").Value
      FDeptID = CInt(billLotTemp.Fields("deptid").Value)
      qtyTemp = CLng(billLotTemp.Fields("qty").Value)
      erpdate = Format(CDate(billLotTemp.Fields("erpcreatedate").Value), "YYYY-MM-DD")
      
                
          '-----begin------
          
            Set adoCmd = New ADODB.Command
            Set adoCmd.ActiveConnection = INIadoCon2
            adoCmd.CommandText = "uspRework_Interface"
            adoCmd.Parameters.Refresh
            adoCmd.CommandType = adCmdStoredProc
        

          Set adoprm1 = New ADODB.Parameter   '工单号
          adoprm1.type = adChar
          adoprm1.Size = 20
          adoprm1.Direction = adParamInput
          adoprm1.Value = woTemp
          adoCmd.Parameters.Append adoprm1
          
          
          Set adoprm2 = New ADODB.Parameter   '成本料号
          adoprm2.type = adChar
          adoprm2.Size = 20
          adoprm2.Direction = adParamInput
          adoprm2.Value = FCostOBJID
          adoCmd.Parameters.Append adoprm2
          
        Set adoPrm3 = New ADODB.Parameter   '部门
          adoPrm3.type = adInteger
          adoPrm3.Direction = adParamInput
          adoPrm3.Value = FDeptID
          adoCmd.Parameters.Append adoPrm3
          
          
          Set adoPrm4 = New ADODB.Parameter   '数量
          adoPrm4.type = adInteger
          adoPrm4.Direction = adParamInput
          adoPrm4.Value = qtyTemp
          adoCmd.Parameters.Append adoPrm4
          

          Set adoPrm5 = New ADODB.Parameter   '日期
          adoPrm5.type = adChar
          adoPrm5.Size = 20
          adoPrm5.Direction = adParamInput
          adoPrm5.Value = erpdate
          adoCmd.Parameters.Append adoPrm5
          
         Set adoPrm6 = New ADODB.Parameter   '是否可修复
          adoPrm6.type = adInteger
          adoPrm6.Direction = adParamInput
          adoPrm6.Value = 1
          adoCmd.Parameters.Append adoPrm6
          
          
        
          adoCmd.Execute

        
        billLotTemp.MoveNext
   
    Loop
    
End If





End If


 
 '2013-07-24 add tsv
 
 Dim qtyAllTemp As Long
 Dim toINVqtyTemp As Long
 Dim productNameTemp As String
 
 If FrmCloseWO.Combo1.text = "TSV" Then

'查询这笔工单PMC开工单的数量
Set TSVQty1 = GetTSVXiaDanQty(woTemp)
qtyAllTemp = CLng(TSVQty1.Fields("qtyall").Value)

'查询抛新ERP数量

Set TSVQty2 = GetTSVToINVQty(woTemp)
toINVqtyTemp = CLng(TSVQty2.Fields("qtyinv").Value)

If qtyAllTemp - toINVqtyTemp > 0 Then

    Set billLotTemp = GetTSV_ViewData(woTemp)
    If (billLotTemp.RecordCount > 0) Then
    
        Do While Not billLotTemp.EOF
          qtyTemp = qtyAllTemp - toINVqtyTemp
          productNameTemp = billLotTemp.Fields("productname").Value
          erpdate = Format(CDate(billLotTemp.Fields("erpcreationdate").Value), "YYYY-MM-DD")
          FCostOBJID = GetFItemID(productNameTemp)
          FDeptID = 350
          
          billLotTemp.MoveNext
       
        Loop

          
                    
              '-----begin------
              
                Set adoCmd = New ADODB.Command
                Set adoCmd.ActiveConnection = INIadoCon2
                adoCmd.CommandText = "uspRework_Interface"
                adoCmd.Parameters.Refresh
                adoCmd.CommandType = adCmdStoredProc
            
    
              Set adoprm1 = New ADODB.Parameter   '工单号
              adoprm1.type = adChar
              adoprm1.Size = 20
              adoprm1.Direction = adParamInput
              adoprm1.Value = woTemp
              adoCmd.Parameters.Append adoprm1
              
              
              Set adoprm2 = New ADODB.Parameter   '成本料号
              adoprm2.type = adChar
              adoprm2.Size = 20
              adoprm2.Direction = adParamInput
              adoprm2.Value = FCostOBJID
              adoCmd.Parameters.Append adoprm2
              
            Set adoPrm3 = New ADODB.Parameter   '部门
              adoPrm3.type = adInteger
              adoPrm3.Direction = adParamInput
              adoPrm3.Value = FDeptID
              adoCmd.Parameters.Append adoPrm3
              
              
              Set adoPrm4 = New ADODB.Parameter   '数量
              adoPrm4.type = adInteger
              adoPrm4.Direction = adParamInput
              adoPrm4.Value = qtyTemp
              adoCmd.Parameters.Append adoPrm4
              
    
              Set adoPrm5 = New ADODB.Parameter   '日期
              adoPrm5.type = adChar
              adoPrm5.Size = 20
              adoPrm5.Direction = adParamInput
              adoPrm5.Value = erpdate
              adoCmd.Parameters.Append adoPrm5
              
             Set adoPrm6 = New ADODB.Parameter   '是否可修复
              adoPrm6.type = adInteger
              adoPrm6.Direction = adParamInput
              adoPrm6.Value = 1
              adoCmd.Parameters.Append adoPrm6
              
              
            
              adoCmd.Execute
    
            
       
        
    End If
    

End If


End If

AddSql2 (strTemp)
AddSql2 (strTemp2)

'MsgBox "工单：" & woTemp & "关闭成功 !", vbInformation, "提示"

End Sub



Public Function GetFItemID(productNameTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

'cmdStr = "select e.FItemID from dbo.tblSmainM2 c,  AIS20130630133218.dbo.cbCostObj  e where  c.物料编号 = e.FNumber and c.料号='" + productNameTemp + "'"
''RSResult = GetStr2(cmdStr)

cmdStr = "select e.FItemID from dbo.tblSmainM2 c,  AIS20141114094336.dbo.cbCostObj  e where  c.物料编号 = e.FNumber and c.料号='" + productNameTemp + "'"

RSResult = GetSqlServerStr(cmdStr)

GetFItemID = RSResult
End Function




Public Function GetProduct() As ADODB.Recordset
'生产料号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
" where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and (b.productname like '18%' or b.productname like '19%' ) order by b.productname"

Set RSResult = getStr(cmdStr)
Set GetProduct = RSResult
End Function

Public Function GetProductOrder(cus As String) As ADODB.Recordset
'生产料号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

If cus = "" Then

    cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
    " where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and (b.productname like '18%' or b.productname like '19%' ) order by b.productname"
Else

'    cmdStr = "select distinct b.productname PID, b.productname from product a, PRODUCTBASE b, tbltsvnpiproduct c" & _
'        " where a.productbaseid = b.productbaseid and b.objectcategory = 'PN' and a.objecttype = 'PN' and (b.productname like '18%' or b.productname like '19%')" & _
'        " and c.qtechptno2 = b.productname and c.customershortname = '" & cus & "' order by b.productname "

     cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
    " where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and (b.productname like '18%' or b.productname like '19%' ) order by b.productname"
End If




Set RSResult = getStr(cmdStr)
Set GetProductOrder = RSResult
End Function



Public Function GetDummyProduct(cus As String) As ADODB.Recordset
'生产料号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
'" where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and b.productname like '18%'  and b.productname like '%WCF' order by b.productname"

' jiayun 2013-12-25

If cus = "" Then
cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
" where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and (b.productname like '18%' or b.productname like '19%' )   order by b.productname"



Else
cmdStr = "select distinct b.productname PID, b.productname from product a, PRODUCTBASE b, tbltsvnpiproduct c" & _
        " where a.productbaseid = b.productbaseid and b.objectcategory = 'PN' and a.objecttype = 'PN' and (b.productname like '18%' or b.productname like '19%')" & _
        " and c.qtechptno2 = b.productname and c.customershortname = '" & cus & "' order by b.productname "
End If


cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
" where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and (b.productname like '18%' or b.productname like '19%' )   order by b.productname"

Set RSResult = getStr(cmdStr)
Set GetDummyProduct = RSResult
End Function



Public Function GetBillType() As ADODB.Recordset
'工单类型
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " SELECT   rtrim(ltrim([名称])) as 名称 ,rtrim(ltrim([说明2])) as 说明2  FROM [erpdata].[dbo].[tblbase] where 说明='工单类型' "

Set RSResult = getSqlStr(cmdStr)
Set GetBillType = RSResult
End Function

Public Function GetWLOBomProduct() As ADODB.Recordset
'WLO Bom 料号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "SELECT  rtrim(ltrim(物料编号)) as 物料编号1 , rtrim(ltrim(物料编号)) as 物料编号2 From [erpdata].[dbo].[TSVtblSetMRule] Where 产线标记 = 3  order by 物料编号"
  
Set RSResult = getSqlStr(cmdStr)
Set GetWLOBomProduct = RSResult
End Function





Public Function GetCloseWo(lineTypeTemp As Integer) As ADODB.Recordset
'查询出sqlserver中的工单号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " SELECT DISTINCT 工单号 PID, 工单号 productname FROM [erpbase].[dbo].[tblllplan] WHERE 产线标记=" & lineTypeTemp & " AND 完工标记=0  ORDER BY 工单号 "
 


Set RSResult = getSqlServerStr2(cmdStr)
Set GetCloseWo = RSResult
End Function





Public Function GetInitPtNo() As ADODB.Recordset
'初始化工单
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select distinct  workordername PID , workordername from a_lotwafers  where txntimestamp>sysdate-2 order by workordername "

Set RSResult = getStr(cmdStr)
Set GetInitPtNo = RSResult
End Function

Public Function GetInitPName() As ADODB.Recordset
'初始化品名
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select distinct  ALTERNATENAME  PID ,ALTERNATENAME from  product  where ALTERNATENAME is not null and ALTERNATENAME not like '%CV'  order by 1 "
  
Set RSResult = getStr(cmdStr)
Set GetInitPName = RSResult
End Function

Public Function GetInitTestName() As ADODB.Recordset
'初始化测试版本号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select distinct  QtechTestVersion PID ,QtechTestVersion  from OITestVersionSet where flag='Y' order by QtechTestVersion"
  
Set RSResult = getStr(cmdStr)
Set GetInitTestName = RSResult
End Function

Public Function GetProductAA() As ADODB.Recordset
'生产料号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
" where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and (b.productname like '18%' or b.productname like '19%') and b.productname not like '18GEC1%' order by b.productname"

Set RSResult = getStr(cmdStr)
Set GetProductAA = RSResult
End Function


Public Function GetProductBB() As ADODB.Recordset
'生产料号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
" where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and (b.productname  like '18%' or b.productname like '19%') order by b.productname"

Set RSResult = getStr(cmdStr)
Set GetProductBB = RSResult
End Function



Public Function GetCOGLVDataList() As ADODB.Recordset
'COG 铝箔袋信息
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select distinct  a.qboxnumber,a.containername,b.barcodeqbox from  GR_COG_IniData a ,TSV_QBOXTBL_GC_COGSEQ b" & _
" Where b.maincontainername = a.containername and b.typename='Lvbodai' order by a.qboxnumber,a.containername,b.barcodeqbox "


Set RSResult = getStr(cmdStr)
Set GetCOGLVDataList = RSResult
End Function



Public Function GetCustomerPONum(customerTemp As String) As ADODB.Recordset
'根据客户代码，查询PO号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select distinct a.po_num PID,  a.po_num productname   from CustomerOItbl_test a where a.flag='Y' and a.qtech_created_date>to_date('2016-01-01','YYYY-MM-DD') " & _
" and a.customershortname='" & customerTemp & "' and po_num is not null order by po_num "


Set RSResult = getStr(cmdStr)
Set GetCustomerPONum = RSResult
End Function



Public Function GetProductBD() As ADODB.Recordset
'生产料号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
" where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and ( b.productname  like '18%' or b.productname like '19%') order by b.productname"

Set RSResult = getStr(cmdStr)
Set GetProductBD = RSResult
End Function



Public Function GetFps37POCompleteICI() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


'
'cmdStr = " SELECT  'COMPLETE' as Event ,c.PO_NUM,c.PO_ITEM,c.MPN,'' as OrderClose,c.CURRENT_WAFER_QTY,COUNT(b.流程卡编号) as Quantity,'' as ScrapQuantity, " & _
'" a.工单号 FROM  [ERPdata].[dbo].[tblStockMove] a, [ERPdata].[dbo].[tblStockMovesub] b ,[ERPBASE].[dbo].[tblCustomerOI] c " & _
'" where a.单据编号='" + billNoTemp + "' and a.客户代码='37' " & _
'" and b.单据编号=a.单据编号 and b.工单号=a.工单号 " & _
'" and c.CUSTOMERSHORTNAME='37' and c.SOURCE_BATCH_ID=a.工单号 and c.PO_NUM <>'' " & _
'" group by   c.PO_NUM,c.PO_ITEM,c.MPN,c.CURRENT_WAFER_QTY, a.工单号 "


'cmdStr = " select X.Event,X.po_num,X.po_item,X.mpn,X.OrderClose,X.die_qty,sum(X.moveinqty) as qty,X.ScrapQuantity,X.source_batch_id from ( " & _
'" select distinct a.containername,'COMPLETE' as Event , wo.po_num, wo.po_item, wo.mpn, '' as OrderClose, wo.die_qty, a.moveinqty, '' as ScrapQuantity, wo.source_batch_id " & _
'"  from a_wiplothistory  a, a_wiplotdetailshistory b, container conn, mfgorder mfg, ib_wohistory ibwo, a_lotwafers  waf, customeroitbl_test wo " & _
'" where a.specname = '5272' and a.creationtimestamp > to_date('2016-04-20', 'YYYY-MM-DD') " & _
'"   and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
'"   and waf.waferscribenumber = substr(conn.containername, 1, instr(conn.containername, '-') - 1)" & _
'"   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername" & _
'"    and wo.customershortname = '37' " & _
'"   and wo.source_batch_id = waf.wafernumber and  a.creationtimestamp>sysdate-7 ) X group by X.Event,X.po_num,X.po_item,X.mpn,X.OrderClose,X.die_qty,X.ScrapQuantity,X.source_batch_id "


   
' cmdStr = " select X.Event,X.po_num,X.po_item,X.mpn,X.OrderClose,X.die_qty,sum(X.moveinqty) as qty,X.ScrapQuantity,X.source_batch_id from ( " & _
'" select distinct a.containername,'COMPLETE' as Event , wo.po_num, wo.po_item, wo.mpn, '' as OrderClose, wo.die_qty, a.moveoutqty as moveinqty, '' as ScrapQuantity, wo.source_batch_id " & _
'"  from a_wiplothistory  a, a_wiplotdetailshistory b, container conn, mfgorder mfg, ib_wohistory ibwo, a_lotwafers  waf, customeroitbl_test wo " & _
'" where a.specname = '5270'  and a.productname ='18X37025B000BR'  and a.creationtimestamp > to_date('2016-04-20', 'YYYY-MM-DD') " & _
'"   and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
'"   and waf.waferscribenumber = substr(conn.containername, 1, instr(conn.containername, '-') - 1)" & _
'"   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername" & _
'"    and wo.customershortname = '37' " & _
'"   and wo.source_batch_id = waf.wafernumber and  a.creationtimestamp>sysdate-7 ) X group by X.Event,X.po_num,X.po_item,X.mpn,X.OrderClose,X.die_qty,X.ScrapQuantity,X.source_batch_id order by X.po_num "
'



cmdStr = " select X.Event, X.po_num, X.po_item, X.mpn, X.OrderClose, x.die_qty*(select  max(rownum) from customeroitbl_test  ct where ct.po_num = X.po_num group by ct.po_num) as dieqty," & _
" sum(X.qty) as qty, X.ScrapQuantity, X.source_batch_id from (select 'COMPLETE' as Event, qty, b.po_num, b.po_item, b.mpn, '' as OrderClose, b.die_qty, '' as ScrapQuantity," & _
" b.source_batch_id from Cus37_5272Qty a, customeroitbl_test b, mappingdatatest c where b.id = c.filename and (c.substrateid = substr(containername, 1, instr(containername, '-A') - 1) or " & _
" c.substrateid = substr(containername, 1, instr(containername, '-A') - 1))) x group by X.Event, X.po_num, X.po_item, X.mpn, X.OrderClose, X.die_qty, X.ScrapQuantity, X.source_batch_id " & _
" order by X.po_num "
     


Set RSResult = getStr(cmdStr)

Set GetFps37POCompleteICI = RSResult
End Function



Public Function GetFps37POShipICI() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


'cmdStr = " select X.Event,X.po_num,X.po_item,X.mpn,X.OrderClose,X.die_qty,sum(X.moveinqty) as qty,X.ScrapQuantity,X.source_batch_id from ( " & _
'" select distinct a.containername,'COMPLETE' as Event , wo.po_num, wo.po_item, wo.mpn, '' as OrderClose, wo.die_qty, a.moveinqty, '' as ScrapQuantity, wo.source_batch_id " & _
'"  from a_wiplothistory  a, a_wiplotdetailshistory b, container conn, mfgorder mfg, ib_wohistory ibwo, a_lotwafers  waf, mappingdatatest map, customeroitbl_test wo " & _
'" where a.specname = '5272' and a.creationtimestamp > to_date('2016-04-20', 'YYYY-MM-DD') " & _
'"   and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
'"   and waf.waferscribenumber = substr(conn.containername, 1, instr(conn.containername, '-') - 1)" & _
'"   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername" & _
'"   and map.substrateid = waf.waferscribenumber and map.customershortname = 'IT'  and wo.customershortname = 'IT' " & _
'"   and wo.source_batch_id = map.lotid ) X group by X.Event,X.po_num,X.po_item,X.mpn,X.OrderClose,X.die_qty,X.ScrapQuantity,X.source_batch_id "

'cmdStr = " select  X.Event,X.po_num,X.po_item,X.mpn,X.MPlant,X.SPlant,X.MPN_DESC,X.source_batch_id,'' as OrderClose,  sum(X.moveinqty) as qty, X.sdate, " & _
'"    'CN' AS COrigin,X.datecode, X.source_batch_id as 工单号,X.containername " & _
'"  from (select distinct a.containername, 'SHIP' as Event, wo.po_num,wo.po_item,wo.mpn,'' as OrderClose, wo.die_qty, a.moveinqty, " & _
'"  '' as ScrapQuantity, wo.source_batch_id, substr(ltrim(wo.offshore_asm_company),1,4) as MPlant, substr(ltrim(wo.offshore_test_company),1,4) as SPlant, " & _
'"  wo.mpn_desc, to_char(a.moveintimestamp,'YYYYMMDD') as sdate, to_char(ibwo.erpcreatedate,'YYWW') as datecode " & _
' "  from a_wiplothistory a , a_wiplotdetailshistory b, container conn, mfgorder mfg, ib_wohistory ibwo, a_lotwafers waf, customeroitbl_test wo " & _
'"   where a.specname = '5272' and a.creationtimestamp > to_date('2016-04-20', 'YYYY-MM-DD') and b.wiplothistoryid = a.wiplothistoryid " & _
'" and conn.containername = a.containername and waf.waferscribenumber = substr(conn.containername, 1,instr(conn.containername, '-') - 1) " & _
'"  and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername  " & _
'"    and wo.customershortname = '37' " & _
'"           and wo.source_batch_id = waf.wafernumber and a.containername like '%-A%' and  a.creationtimestamp>sysdate-7 ) X " & _
'"  group by X.Event,  X.po_num,X.po_item, X.mpn, X.OrderClose,X.die_qty,X.ScrapQuantity,X.source_batch_id, X.MPlant, X.SPlant , X.mpn_desc, X.sdate, X.DateCode ,X.containername "
'



'cmdStr = " select  X.Event,X.po_num,X.po_item,X.mpn,X.MPlant,X.SPlant,X.MPN_DESC,X.source_batch_id,'' as OrderClose,  sum(X.moveinqty) as qty, X.sdate, " & _
'"    'CN' AS COrigin,X.datecode, X.source_batch_id as 工单号,X.containername " & _
'"  from (select distinct a.containername, 'SHIP' as Event, wo.po_num,wo.po_item,wo.mpn,'' as OrderClose, wo.die_qty, a.moveinqty, " & _
'"  '' as ScrapQuantity, wo.source_batch_id, substr(ltrim(wo.offshore_asm_company),1,4) as MPlant, substr(ltrim(wo.offshore_test_company),1,4) as SPlant, " & _
'"  wo.mpn_desc, to_char(a.moveintimestamp,'YYYYMMDD') as sdate, to_char(ibwo.erpcreatedate,'YYWW') as datecode " & _
' "  from a_wiplothistory a , a_wiplotdetailshistory b, container conn, mfgorder mfg, ib_wohistory ibwo, a_lotwafers waf, customeroitbl_test wo " & _
'"   where a.specname = '5272' and a.creationtimestamp > to_date('2016-04-20', 'YYYY-MM-DD') and b.wiplothistoryid = a.wiplothistoryid " & _
'" and conn.containername = a.containername and waf.waferscribenumber = substr(conn.containername, 1,instr(conn.containername, '-') - 1) " & _
'"  and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername  " & _
'"  and wo.customershortname = '37' " & _
'"  and wo.source_batch_id = waf.wafernumber and a.containername like '%-A%' and  a.creationtimestamp>sysdate-7   and a.containername<>'10001-A-01' " & _
'"union   select distinct a.containername,'SHIP' as Event, wo.po_num,wo.po_item, wo.mpn,'' as OrderClose,wo.die_qty,waf.ndpw, '' as ScrapQuantity, " & _
'" wo.source_batch_id, substr(ltrim(wo.offshore_asm_company), 1, 4) as MPlant,substr(ltrim(wo.offshore_test_company), 1, 4) as SPlant,wo.mpn_desc,to_char(a.moveintimestamp, 'YYYYMMDD') as sdate, to_char(ibwo.erpcreatedate, 'YYWW') as datecode " & _
'"  from a_wiplothistory a,a_wiplotdetailshistory b,container conn,mfgorder mfg,ib_wohistory  ibwo,a_lotwafers waf, customeroitbl_test wo " & _
'" where a.specname = '5272' and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
'"   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername and wo.source_batch_id = waf.wafernumber and wo.customershortname = '37' " & _
'"   and a.containername not like '%-F%' and waf.containerid=conn.containerid and a.creationtimestamp > sysdate - 7" & _
'" ) X " & _
'"  group by X.Event,  X.po_num,X.po_item, X.mpn, X.OrderClose,X.die_qty,X.ScrapQuantity,X.source_batch_id, X.MPlant, X.SPlant , X.mpn_desc, X.sdate, X.DateCode ,X.containername  order by  X.containername "



'cmdStr = " select X.Event,X.po_num,X.po_item,X.mpn,X.MPlant,X.SPlant,X.mpn_desc,X.source_batch_id,X.OrderClose,sum(X.qty) as qty,x.sdate,'CN' AS COrigin, " & _
'" X.date_code, X.source_batch_id as 工单号,'' as containername from ( " & _
'" select 'SHIP' as Event  ,qty ,b.po_num,b.po_item, b.mpn,'' as OrderClose ,b.die_qty ,'' as ScrapQuantity,b. source_batch_id ," & _
'" substr(ltrim(b.offshore_asm_company),1,4) as MPlant,substr(ltrim(b.offshore_test_company),1,4) as SPlant,mpn_desc,to_char(sysdate,'YYYYMMDD') as sdate, " & _
'" b.Date_Code from  cus37_5272qtyNotMerg a , customeroitbl_test b " & _
'" where b.mtrl_num=substr( containername,1,instr(containername,'-A')-1) " & _
'" ) x group by X.Event,X.po_num,X.po_item,X.mpn,X.MPlant,X.SPlant,X.mpn_desc,X.OrderClose,X.die_qty,X.ScrapQuantity,X.source_batch_id,x.sdate,X.date_code"


 cmdStr = "select distinct X.Event,X.po_num,X.po_item,X.mpn,'3179' as MPlant,X.SPlant,X.mpn_desc,X.source_batch_id,X.OrderClose,X.qty ,x.sdate,'CN' AS COrigin, " & _
" X.date_code, X.source_batch_id as 工单号, X.containername from  cus37_5272qtydetails X "


Set RSResult = getStr(cmdStr)

Set GetFps37POShipICI = RSResult
End Function



Public Function GetQty37NGQty(lotIdTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long


'cmdStr = "select COUNT(waferid) as ID from  [erpdata].[dbo].[TSVtblBAOFEI] a where a.LOTID='" & lotidTemp & "' "
'

cmdStr = " select nvl( sum(c.qty),0) as ID from  a_lotwafers  a ,container b ,historymainline c " & _
" where a.wafernumber='" & lotIdTemp & "'  and b.containerid=a.containerid  and b.containername  like '%-F%' " & _
" and c.containername=b.containername and c.specname='5272' and c.cdoname=  'MoveInLot' "
 
RSResult = GetSeq(cmdStr)
GetQty37NGQty = RSResult
End Function


Public Function GetQty37OutQtyICI(lotIdTemp As String) As Long
Dim cmdStr As String
Dim RSResult As Long

'查询5272站所有数据

 
'cmdStr = "select COUNT(waferid) as ID from  [erpdata].[dbo].[TSVtblBAOFEI] a where a.LOTID='" & lotIDTemp & "' "

'cmdStr = " SELECT COUNT(流程卡编号) as ID FROM [erpdata].[dbo].[tblStockMovesub]  WHERE 工单号='" & lotidTemp & "'  and 单据编号 like 'F%' "


cmdStr = "  select nvl( sum(A.NDPW),0) as ID from a_lotwafers a, container b, historymainline c " & _
" where a.wafernumber = '" & lotIdTemp & "' and b.containerid = a.containerid " & _
" and c.containername = b.containername and c.specname = '5272' and c.cdoname = 'MoveInLot'  and b.containername  like '%-A%' "
   
 

RSResult = GetSeq(cmdStr)
GetQty37OutQtyICI = RSResult
End Function

Public Function GetWeiPiItem(mainitemIdTemp As String) As ADODB.Recordset
'查询小类型
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


'cmdStr = "select id,smallname from  spc_smalltype where bigtypdId=" & mainitemIdTemp & " and flag='Y' order by smallname "
'

cmdStr = "  select a.containername as id , a.containername||' '||a.moveinqty as name from a_wiplot a ,a_lotwafers b ,ib_wohistory c where a.specname='5272'" & _
        " and b.containerid=a.containerid  and c.ordername=b.workordername  and  not exists( SELECT * FROM  woLotWeiPi d where d.containername=a.containername)   and c.customerpn='" & mainitemIdTemp & "' "
                 

         
Set RSResult = getStr(cmdStr)
Set GetWeiPiItem = RSResult
End Function



Public Function GetDetailData(waferlotTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select count(*) num,sum(fgdieqty) sumQty from IB_WAFERLIST where waferlot='" + waferlotTemp + "'"
  
Set RSResult = getStr(cmdStr)

Set GetDetailData = RSResult
End Function





Public Function GetWO_GC_Die(waferlotTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select dieqty from  tblCustomerDieQty    where CustomerName='GC' and  CustomerPT ='" + waferlotTemp + "'  "

Set RSResult = getStr(cmdStr)

Set GetWO_GC_Die = RSResult
End Function


Public Function GetWO_GC_Ver(ptTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select Gcversion from  tblCustomerDieQty    where CustomerName='GC' and  CustomerPT ='" + ptTemp + "' and Gcversion is not null "



Set RSResult = getStr(cmdStr)

Set GetWO_GC_Ver = RSResult
End Function


Public Function GetHeaderData(waferlotTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select current_wafer_qty,die_qty from  CustomerOItbl_test where source_batch_id='" + waferlotTemp + "' and flag='Y' "
 
Set RSResult = getStr(cmdStr)

Set GetHeaderData = RSResult
End Function

Public Function GetTSVXiaDanQty(waferlotTemp As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = " select sum(dieqty) as  qtyall from (select distinct waferid,dieqty from ib_waferlist where ordername='" + waferlotTemp + "') "


Set RSResult = getStr(cmdStr)

Set GetTSVXiaDanQty = RSResult
End Function

Public Function GetTSVToINVQty(waferlotTemp As String) As ADODB.Recordset

'Dim cmdStr As String
'Dim RSResult As New ADODB.Recordset
'
''cmdStr = " select nvl(sum(ndpw),0 )  as  qtyinv  from  TSV_QBOXNUMBER_DETAILS where workordername='" + waferlotTemp + "' and specname in ('5270','5275','5271') "
'
'
'cmdStr = "  SELECT  SUM(e.入库数) as  qtyinv FROM erpdata..tblPackToHouseRec e WHERE e.大工单 = '" + waferlotTemp + "' "
'
'Set RSResult = getSqlServerStr2(cmdStr)
'
'Set GetTSVToINVQty = RSResult

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = " select nvl(sum(ndpw),0 )  as  qtyinv  from  TSV_QBOXNUMBER_DETAILS where workordername='" + waferlotTemp + "' and specname in ('5270','5275','5271') "


cmdStr = "  SELECT  ISNULL(SUM(e.入库数),0) as  qtyinv FROM erpdata..tblPackToHouseRec e WHERE e.大工单 = '" + waferlotTemp + "' "

Set RSResult = getSqlServerStr2(cmdStr)

Set GetTSVToINVQty = RSResult

End Function



Public Sub updateHeaderDate(waferlotTemp As String, statusTemp As String, qtyTemp As Long)

Dim sqlTemp As String

'2014-03-10 jiayun modify

'判断这个lotid ,mapping中的片数,与开工单的片数，是不是一致

Dim slectResult As Boolean

slectResult = False

slectResult = JudgeUpdateWOStatus(waferlotTemp)

If Not slectResult Then

sqlTemp = "Update CustomerOItbl_test set flag='" & statusTemp & "',DownQty=" & qtyTemp & " where source_batch_id='" & waferlotTemp & "' and flag='Y'"

AddSql (sqlTemp)
 
 End If


End Sub
Public Sub updateKR001Map(waferIdTemp As String, goodDieQtyTemp As Long, ngDieQtyTemp As Long)
Dim cmdStr As String
Dim cmdStr2 As String

On Error GoTo DealError

'导入Mapping
Cnn.BeginTrans
     
cmdStr = "update mappingdatatest  a " & _
" set a.passbincount=" & goodDieQtyTemp & ",a.failbincount=" & ngDieQtyTemp & " ,a.QTECH_LASTUPDATE_BY = '" & gUserRealName & "',a.QTECH_LASTUPDATE_DATE = sysdate " & _
" where customershortname='KR001'  and a.substrateid= '" & waferIdTemp & "' "

cmdStr2 = "update [ERPBASE].[dbo].[tblmappingDataMG]   " & _
" set passbincount=" & goodDieQtyTemp & ",failbincount=" & ngDieQtyTemp & "  " & _
" where customershortname='KR001'  and substrateid= '" & waferIdTemp & "' "

AddSql (cmdStr)
 
AddSql2 (cmdStr2)

If AddSql(cmdStr) = 0 Then
    MsgBox "没有更新到数据", vbCritical, "警告"
  
    Exit Sub
End If

Cnn.CommitTrans

SumCount = SumCount + 1
   
 Exit Sub
 
DealError:

Cnn.RollbackTrans

End Sub

Public Function GetKRGrossDie(waferid As String) As Integer
GetKRGrossDie = 3380
Dim cmdOra As String

cmdOra = "select nvl((a.passbincount + a.failbincount), 3380)  from mappingdatatest a WHERE a.substrateid = ( '" & waferid & "')"

If GetSeqTW(cmdOra) <> 0 Then
    GetKRGrossDie = GetSeqTW(cmdOra)
End If


End Function
Public Sub updateKR002Map(waferIdTemp As String, goodDieQtyTemp As Long, ngDieQtyTemp As Long)
Dim cmdStr As String
Dim cmdStr2 As String   '2- 代表 SqlServer2008

On Error GoTo DealError

'导入Mapping
Cnn.BeginTrans
                 
cmdStr = "update mappingdatatest  a " & _
" set a.passbincount=" & goodDieQtyTemp & ",a.failbincount=" & ngDieQtyTemp & " ,a.QTECH_LASTUPDATE_BY = '" & gUserRealName & "',a.QTECH_LASTUPDATE_DATE = sysdate " & _
" where customershortname='KR002'  and a.substrateid= '" & waferIdTemp & "' "

cmdStr2 = "update [ERPBASE].[dbo].[tblmappingData]    " & _
" set passbincount=" & goodDieQtyTemp & ",failbincount=" & ngDieQtyTemp & "  " & _
" where customershortname='KR002'  and substrateid= '" & waferIdTemp & "' "

If AddSql(cmdStr) = 0 Then
    MsgBox "没有更新到数据", vbCritical, "警告"

    Exit Sub
End If
 
AddSql2 (cmdStr2)
 
 Cnn.CommitTrans
 
SumCount = SumCount + 1
   
 Exit Sub
 
DealError:

 SumCount = SumCount - 1

Cnn.RollbackTrans

End Sub


Public Sub ONLotIDClose(waferlotTemp As String)

Dim sqlTemp As String

sqlTemp = "insert into  On_WO_HisTory (lotid,waferid,flag,createdate) values ('" & waferlotTemp & "',' ','Y',sysdate) "

AddSql (sqlTemp)

End Sub




Public Function JudgeUpdateWOStatus(lotIdTemp As String) As Boolean
'耞琌睰筁
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "  select * from mappingdatatest a where a. lotid ='" + lotIdTemp + "' and a.substrateid not in (select b.waferid from  ib_waferlist b where b.waferlot='" + lotIdTemp + "' )"
      
slectResult = QueryStr(cmdStr)
JudgeUpdateWOStatus = slectResult
End Function


'2012-12-07 jiayun add 添加GC WO_NO
Public Sub updateHeaderDateForGC(waferlotTemp As String, statusTemp As String, qtyTemp As Long, woNoTemp As String)

Dim sqlTemp As String

sqlTemp = "Update CustomerOItbl_test set flag='" & statusTemp & "',DownQty=" & qtyTemp & " where source_batch_id='" & waferlotTemp & "' and flag='Y' and mtrl_num='" & woNoTemp & "' "

 AddSql (sqlTemp)


End Sub


'2014-01-15 jiayun add NpiProduct
Public Sub AddNpiProduct(nPIProductTemp As NpiProduct)

Dim sqlTemp As String
Dim sqlTemp1 As String
Dim sqlid As String
Dim id As Long
Dim Rs3 As New ADODB.Recordset

sqlid = "  SELECT  NpiProduct_SEQ.Nextval FROM DUAL "

 If Rs3.State = adStateOpen Then Rs3.Close
    Rs3.Open sqlid, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
   id = Val(Rs3.Fields(0).Value)

sqlTemp = " insert into TBLTsvNpiProduct(ID,CUSTOMERSHORTNAME ,QTECHPTNO,QtechPTNo2,CUSTOMERPTNO1,CUSTOMERPTNO2 , " & _
          " CUSTOMERDIEQTY,QTECHDIEQTY,XIANGSU ,USEDAREA ,STRUCKSTR1," & _
          " STRUCKSTR2 ,STRUCKSTR3, FLAG ,CREATED_BY ,CREATED_DATE," & _
          " ST_DATE,TT_DATE ,PT_DATE,CUSTOMERPTNO3,CUSTOMERPTNO4,CUSTOMERPTNO5,CUSTOMERPTNO6,PKG_TYPE,Residual,MARKING_CODE, P_E ,MAPPING,MARKETLASTUPDATE_BY) values ( " & _
          " '" & id & "','" & nPIProductTemp.CUSTOMERSHORTNAME & "','" & nPIProductTemp.qtechPTNo & "','" & nPIProductTemp.QtechPTNo2 & "','" & nPIProductTemp.CustomerPTNo1 & "','" & nPIProductTemp.CustomerPTNo2 & "', " & _
          "  '" & nPIProductTemp.CustomerDieQty & "','" & nPIProductTemp.QtechDieQty & "','" & nPIProductTemp.XiangSu & "','" & nPIProductTemp.UsedArea & "','" & nPIProductTemp.StruckStr1 & "', " & _
          "  '" & nPIProductTemp.StruckStr2 & "','" & nPIProductTemp.StruckStr3 & "','Y','" & nPIProductTemp.CreateBy & "',sysdate, " & _
          "  '" & nPIProductTemp.STDate & "','" & nPIProductTemp.TTDate & "','" & nPIProductTemp.PTDate & "' ,'" & nPIProductTemp.CustomerPTNo3 & "', " & _
          " '" & nPIProductTemp.CustomerPTNo4 & "','" & nPIProductTemp.CustomerPTNo5 & "','" & nPIProductTemp.CustomerPTNo6 & "','" & nPIProductTemp.PKG & "','" & nPIProductTemp.residual & "','" & nPIProductTemp.MARKINGCODE & "','" & nPIProductTemp.ProducEng & "','" & nPIProductTemp.MAPPING & "','" & nPIProductTemp.WaferPN & "')"
  
 
sqlTemp1 = " insert into erptemp..TBLTsvNpiProduct(ID,CUSTOMERSHORTNAME ,QTECHPTNO,QtechPTNo2,CUSTOMERPTNO1,CUSTOMERPTNO2 , " & _
          " CUSTOMERDIEQTY,QTECHDIEQTY,XIANGSU ,USEDAREA ,STRUCKSTR1," & _
          " STRUCKSTR2 ,STRUCKSTR3, FLAG ,CREATED_BY ,CREATED_DATE," & _
          " ST_DATE,TT_DATE ,PT_DATE,CUSTOMERPTNO3,CUSTOMERPTNO4,CUSTOMERPTNO5,CUSTOMERPTNO6,PKG_TYPE,Residual,MARKING_CODE, P_E,MAPPING,MARKETLASTUPDATE_BY) values ( " & _
          " '" & id & "','" & nPIProductTemp.CUSTOMERSHORTNAME & "','" & nPIProductTemp.qtechPTNo & "','" & nPIProductTemp.QtechPTNo2 & "','" & nPIProductTemp.CustomerPTNo1 & "','" & nPIProductTemp.CustomerPTNo2 & "', " & _
          "  '" & nPIProductTemp.CustomerDieQty & "','" & nPIProductTemp.QtechDieQty & "','" & nPIProductTemp.XiangSu & "','" & nPIProductTemp.UsedArea & "','" & nPIProductTemp.StruckStr1 & "', " & _
          "  '" & nPIProductTemp.StruckStr2 & "','" & nPIProductTemp.StruckStr3 & "','Y','" & nPIProductTemp.CreateBy & "',CONVERT(varchar(100),GETDATE(),20), " & _
          "  '" & nPIProductTemp.STDate & "','" & nPIProductTemp.TTDate & "','" & nPIProductTemp.PTDate & "' ,'" & nPIProductTemp.CustomerPTNo3 & "', " & _
          " '" & nPIProductTemp.CustomerPTNo4 & "','" & nPIProductTemp.CustomerPTNo5 & "','" & nPIProductTemp.CustomerPTNo6 & "','" & nPIProductTemp.PKG & "','" & nPIProductTemp.residual & "','" & nPIProductTemp.MARKINGCODE & "','" & nPIProductTemp.ProducEng & "','" & nPIProductTemp.MAPPING & "','" & nPIProductTemp.WaferPN & "')"
   
  
AddSql (sqlTemp)
  
AddSql2 (sqlTemp1)

End Sub

Public Sub AddCode37(strINfo As CODE37)

Dim sqlTemp As String

sqlTemp = " insert into code37(CUSTOMER ,DEVICE,BLINE,CODE,STATUS,TIMESTAMP ,SEQ ) values ( " & _
          "  '" & strINfo.strCus & "', '" & strINfo.strDev & "', '" & strINfo.strBline & "', '" & strINfo.strCode & "', '" & strINfo.strStatus & "', sysdate, NpiProduct_SEQ.Nextval)"
          
AddSql (sqlTemp)

End Sub

Public Sub AddPOPrice(nPoPriceTemp As POPrice)
Dim strSql  As String
Dim strSql2 As String

strSql = " insert into TSV_MD_POPrice( ID ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE ," & "  PO_TYPE ,PT , QTY ,PRICE ,UNIT ," & "  FileName , Flag, QTECH_CREATED_BY, QTECH_CREATED_DATE,PeaceQty,CUSTAA,BJ,DIE_PRICE,RATE  ) values ( " & "  '" & nPoPriceTemp.id & "','" & nPoPriceTemp.CUSTOMERSHORTNAME & "','" & nPoPriceTemp.customerName & "','" & nPoPriceTemp.PONo & "',to_date('" & nPoPriceTemp.PODATE & "','YYYY-MM-DD HH24:MI:SS'), " & "  '" & nPoPriceTemp.POType & "','" & nPoPriceTemp.pt & "','" & nPoPriceTemp.QTY & "','" & nPoPriceTemp.Price & "','" & nPoPriceTemp.unit & "', " & "  '" & nPoPriceTemp.File & "','Y','" & nPoPriceTemp.CreateBy & "',sysdate," & nPoPriceTemp.peaseQty & ",'" & nPoPriceTemp.custAA & "' ,'" & nPoPriceTemp.bj & "', '" & nPoPriceTemp.SingDie & "','" & nPoPriceTemp.SingWafer & "' ) "
strSql2 = "insert into erptemp..tblBB_CSRPO values('" & nPoPriceTemp.CUSTOMERSHORTNAME & "','" & nPoPriceTemp.PONo & "','10','" & nPoPriceTemp.bj & "',  " & " '" & nPoPriceTemp.pt & "','" & nPoPriceTemp.peaseQty & "','" & nPoPriceTemp.QTY & "','" & nPoPriceTemp.Price & "','" & nPoPriceTemp.SingDie & "' ,'" & nPoPriceTemp.unit & "','',CONVERT(varchar(100), getdate(),20),'" & nPoPriceTemp.SingWafer & "')  "
AddSql (strSql)
AddSql2 (strSql2)

End Sub

Public Sub UpdatePOPriceStatus(id As Long, userNameTemp As String)

Dim sqlTemp As String

sqlTemp = " update TSV_MD_POPrice set Flag='N',QTECH_LASTUPDATE_BY='" & userNameTemp & "',QTECH_LASTUPDATE_DATE = sysdate where  id=" & id & " and flag='Y' "
 
AddSql (sqlTemp)


End Sub



'2015-08-18 jiayun add
Public Sub ModifyPOPrice(nPoPriceTemp As POPrice)

Dim sqlTemp As String
Dim sqlTemp1 As String
Dim sqlTemp2 As String
Dim userName As String


Dim id As Long

id = nPoPriceTemp.id
userName = nPoPriceTemp.CreateBy
  
  '--插入修改历史表
  
 sqlTemp = " insert into TSV_MD_POPriceModiyHis " & _
           " select   ID,CUSTOMERSHORTNAME ,CUSTOMERNAME,PO_NUM,PO_DATE, " & _
           " PO_TYPE, PT ,QTY ,PRICE ,UNIT ,FILENAME , " & _
           "  FLAG, QTECH_CREATED_BY , QTECH_CREATED_DATE,QTECH_LASTUPDATE_BY,QTECH_LASTUPDATE_DATE ,'" & userName & "',sysdate,BJ " & _
           " from  TSV_MD_POPrice  a where id=" & id & " and flag='Y' "
           
 AddSql (sqlTemp)
 
 
   '--再修改表
   
   
    sqlTemp = " Update TSV_MD_POPrice set PO_NUM='" & nPoPriceTemp.PONo & "', " & _
              " PO_DATE=to_date('" & nPoPriceTemp.PODATE & "','YYYY-MM-DD HH24:MI:SS')," & _
              " PO_TYPE='" & nPoPriceTemp.POType & "',PT='" & nPoPriceTemp.pt & "'," & _
              " QTY='" & nPoPriceTemp.QTY & "',BJ='" & nPoPriceTemp.bj & "',PRICE='" & nPoPriceTemp.Price & "', " & _
              " UNIT='" & nPoPriceTemp.unit & "',FILENAME='" & nPoPriceTemp.File & "'," & _
              " QTECH_LASTUPDATE_BY='" & nPoPriceTemp.CreateBy & "',QTECH_LASTUPDATE_DATE = sysdate , peaceqty = '" & nPoPriceTemp.peaseQty & "' " & _
              " ,DIE_PRICE = '" & nPoPriceTemp.DIE_PRICE & "'  where  id=" & id & " and flag='Y' "

              
 AddSql (sqlTemp)
 
 
 sqlTemp2 = " update erptemp..tblBB_CSRPO set customer = '" & nPoPriceTemp.CUSTOMERSHORTNAME & "', po_num = '" & nPoPriceTemp.PONo & "', " & _
            " quote = '" & nPoPriceTemp.bj & "', fab_device = '" & nPoPriceTemp.pt & "', qty = '" & nPoPriceTemp.peaseQty & "', " & _
            " qty_die = '" & nPoPriceTemp.QTY & "', wafer_price = '" & nPoPriceTemp.Price & "',die_price = '" & nPoPriceTemp.DIE_PRICE & "', " & _
            " currency = '" & nPoPriceTemp.unit & "' where po_num = '" & nPoPriceTemp.PONo & "' and fab_device = '" & nPoPriceTemp.pt & "' "
 
 
  
 AddSql2 (sqlTemp2)
 

End Sub





Public Sub AddBaoFei(nPIProductTemp As Baofei)

'2015-01-11 jiayun add  ,把数据抛一份到 SqlServer


Dim sqlTemp As String

Dim cmdStrSql As String



 sqlTemp = " insert into TBLTSVBaoFei( id ,PutInDept,WaferID ,LotID ,GDDie,NGDie ,CUSTOMERPTNO,QTECHPTNO,ProductName,ERR_DATE,PutIn_DATE,FLAG,CREATED_BY,CREATED_DATE,BillStatus,typename) " & _
           " values (" & nPIProductTemp.id & ",'" & nPIProductTemp.putInDept & "','" & nPIProductTemp.waferid & "','" & nPIProductTemp.LOTID & "'," & nPIProductTemp.gDDie & "," & nPIProductTemp.ngdie & ",'" & nPIProductTemp.customerPTNo & "','" & nPIProductTemp.qtechPTNo & "','" & nPIProductTemp.productName & "','" & nPIProductTemp.err_date & "','" & nPIProductTemp.putIn_date & "','Y','" & nPIProductTemp.CreateBy & "',sysdate,'1','" & nPIProductTemp.TYPENAME & "')"
 
 
  cmdStrSql = " insert into [erpdata].[dbo].[TSVtblBAOFEI]( id ,PutInDept,WaferID ,LotID ,GDDie,NGDie ,CUSTOMERPTNO,QTECHPTNO,ProductName,ERR_DATE,PutIn_DATE,FLAG,CREATED_BY,CREATED_DATE,BillStatus,typename) " & _
           " values (" & nPIProductTemp.id & ",'" & nPIProductTemp.putInDept & "','" & nPIProductTemp.waferid & "','" & nPIProductTemp.LOTID & "'," & nPIProductTemp.gDDie & "," & nPIProductTemp.ngdie & ",'" & nPIProductTemp.customerPTNo & "','" & nPIProductTemp.qtechPTNo & "','" & nPIProductTemp.productName & "','" & nPIProductTemp.err_date & "','" & nPIProductTemp.putIn_date & "','Y','" & nPIProductTemp.CreateBy & "',getdate(),'1','" & nPIProductTemp.TYPENAME & "')"
 
 
 
  
 AddSql (sqlTemp)

 AddSql2 (cmdStrSql)

End Sub

'2014-01-15 jiayun add NpiProduct
Public Sub ModifyNpiProduct(nPIProductTemp As NpiProduct, idTemp As Long)

    Dim sqlTemp As String
    Dim sqlTemp1 As String
    
    If gUserName = "16642" Or gUserName = "15236" Or gUserName = "12725" Or gUserName = "16452" Or gUserName = "14117" Or gUserName = "12089" Or gUserName = "15507" Or gUserName = "16368" Or gUserName = "08240" Then
        MsgBox "当前账号只有修改客户机种的权限", vbInformation, "提示"
    
        sqlTemp = " Update TBLTsvNpiProduct set CUSTOMERPTNO1='" & nPIProductTemp.CustomerPTNo1 & "' Where id = " & idTemp & ""
        
        sqlTemp1 = " Update erptemp..TBLTsvNpiProduct set CUSTOMERPTNO1='" & nPIProductTemp.CustomerPTNo1 & "' Where id = " & idTemp & ""
    Else
         sqlTemp = " Update TBLTsvNpiProduct " & _
           " set CUSTOMERSHORTNAME='" & nPIProductTemp.CUSTOMERSHORTNAME & "',QTECHPTNO='" & nPIProductTemp.qtechPTNo & "',QTECHPTNO2='" & nPIProductTemp.QtechPTNo2 & "',CUSTOMERPTNO1='" & nPIProductTemp.CustomerPTNo1 & "', " & _
           " CUSTOMERPTNO2='" & nPIProductTemp.CustomerPTNo2 & "',CUSTOMERDIEQTY='" & nPIProductTemp.CustomerDieQty & "',QTECHDIEQTY='" & nPIProductTemp.QtechDieQty & "', " & _
           " XIANGSU='" & nPIProductTemp.XiangSu & "',USEDAREA='" & nPIProductTemp.UsedArea & "',STRUCKSTR1='" & nPIProductTemp.StruckStr1 & "', " & _
           " STRUCKSTR2='" & nPIProductTemp.StruckStr2 & "',STRUCKSTR3='" & nPIProductTemp.StruckStr3 & "',ST_DATE='" & nPIProductTemp.STDate & "', " & _
           " TT_DATE='" & nPIProductTemp.TTDate & "',PT_DATE='" & nPIProductTemp.PTDate & "',lastupdate_by='" & nPIProductTemp.CreateBy & "',lastupdate_date=sysdate,CUSTOMERPTNO3='" & nPIProductTemp.CustomerPTNo3 & "',CUSTOMERPTNO4='" & nPIProductTemp.CustomerPTNo4 & "',CUSTOMERPTNO5='" & nPIProductTemp.CustomerPTNo5 & "',CUSTOMERPTNO6='" & nPIProductTemp.CustomerPTNo6 & "'," & _
           " PKG_TYPE = '" & nPIProductTemp.PKG & "',Residual = '" & nPIProductTemp.residual & "',MARKING_CODE =  '" & nPIProductTemp.MARKINGCODE & "', P_E = '" & nPIProductTemp.ProducEng & "', MAPPING='" & nPIProductTemp.MAPPING & "' , MARKETLASTUPDATE_BY = '" & nPIProductTemp.WaferPN & "'  Where id = " & idTemp & ""
           
           
          sqlTemp1 = " Update erptemp..TBLTsvNpiProduct " & _
           " set CUSTOMERSHORTNAME='" & nPIProductTemp.CUSTOMERSHORTNAME & "',QTECHPTNO='" & nPIProductTemp.qtechPTNo & "',QTECHPTNO2='" & nPIProductTemp.QtechPTNo2 & "',CUSTOMERPTNO1='" & nPIProductTemp.CustomerPTNo1 & "', " & _
           " CUSTOMERPTNO2='" & nPIProductTemp.CustomerPTNo2 & "',CUSTOMERDIEQTY='" & nPIProductTemp.CustomerDieQty & "',QTECHDIEQTY='" & nPIProductTemp.QtechDieQty & "', " & _
           " XIANGSU='" & nPIProductTemp.XiangSu & "',USEDAREA='" & nPIProductTemp.UsedArea & "',STRUCKSTR1='" & nPIProductTemp.StruckStr1 & "', " & _
           " STRUCKSTR2='" & nPIProductTemp.StruckStr2 & "',STRUCKSTR3='" & nPIProductTemp.StruckStr3 & "',ST_DATE='" & nPIProductTemp.STDate & "', " & _
           " TT_DATE='" & nPIProductTemp.TTDate & "',PT_DATE='" & nPIProductTemp.PTDate & "',lastupdate_by='" & nPIProductTemp.CreateBy & "',lastupdate_date=GetDate(),CUSTOMERPTNO3='" & nPIProductTemp.CustomerPTNo3 & "',CUSTOMERPTNO4='" & nPIProductTemp.CustomerPTNo4 & "',CUSTOMERPTNO5='" & nPIProductTemp.CustomerPTNo5 & "',CUSTOMERPTNO6='" & nPIProductTemp.CustomerPTNo6 & "'," & _
           " PKG_TYPE = '" & nPIProductTemp.PKG & "',Residual = '" & nPIProductTemp.residual & "',MARKING_CODE =  '" & nPIProductTemp.MARKINGCODE & "', P_E = '" & nPIProductTemp.ProducEng & "', MAPPING='" & nPIProductTemp.MAPPING & "', MARKETLASTUPDATE_BY = '" & nPIProductTemp.WaferPN & "'    Where id = " & idTemp & ""
   
    End If

    AddSql (sqlTemp)
    AddSql2 (sqlTemp1)

End Sub

Public Sub ModCode37(strINfo As CODE37, idTemp As Long)

    Dim sqlTemp As String
   
    sqlTemp = " Update CODE37 set BLINE='" & strINfo.strBline & "',CODE='" & strINfo.strCode & "',STATUS='" & strINfo.strStatus & "', TIMESTAMP = sysdate Where seq= " & idTemp & ""

    AddSql (sqlTemp)

End Sub


Public Sub DelDataNpiProduct(idTemp As Long)

Dim sqlTemp As String
Dim sqlTemp1 As String
 

AddSql ("insert into TBLTsvNpiProduct_BAK select * from TBLTsvNpiProduct where  id = " & idTemp & " ")
 
sqlTemp = " delete from TBLTsvNpiProduct where  id = " & idTemp & ""
sqlTemp1 = " delete from erptemp..TBLTsvNpiProduct where  id = " & idTemp & ""

 AddSql (sqlTemp)
 AddSql2 (sqlTemp1)
 
End Sub

Public Sub DelCode37(idTemp As Long)

    Dim sqlTemp As String
  
    sqlTemp = " delete from code37 where  seq = " & idTemp & ""

    AddSql (sqlTemp)

End Sub



'2014-01-15 jiayun add NpiProductPrice
Public Sub ModifyNpiProductPrice(nPIProductTemp As NpiProduct, idTemp As Long)

Dim sqlTemp As String

   
sqlTemp = " Update TBLTsvNpiProduct " & _
        " set FzFreeUSD='" & nPIProductTemp.FzFreeUSD & "',TestFreeUSD='" & nPIProductTemp.TestFreeUSD & "',FzFreeRMB='" & nPIProductTemp.FzFreeRMB & "',TestFreeRMB='" & nPIProductTemp.TestFreeRMB & "', " & _
        " NreFree='" & nPIProductTemp.NreFree & "',NreMethod='" & nPIProductTemp.NreMethod & "',UpdatePrice2='" & nPIProductTemp.UpdatePrice2 & "', " & _
        " UpdatePrice1='" & nPIProductTemp.UpdatePrice1 & "',MarketLASTUPDATE_BY='" & nPIProductTemp.CreateBy & "',MarketLASTUPDATE_DATE=sysdate " & _
        " Where id = " & idTemp & ""
 AddSql (sqlTemp)


End Sub



Public Function GetReportData_Where(strTemp As String) As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 cmdStr = "select  LotID,LotID|| substr('0'||wafer_id,-2,2) as WaferID,substr(Productid,2,4) as MarkingCode,passbincount,qtech_created_date" & _
          " from  mappingDataTest  where customershortname='GC' and LotID like '" & strTemp & "%' order by 2 "
         

Set RSResult = getStr(cmdStr)
Set GetReportData_Where = RSResult
End Function

Public Function GetpfData() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 cmdStr = "select fieldName,fieldValue,resultValue,other from  tblsetpf where flag='Y'"
         

Set RSResult = getStr(cmdStr)
Set GetpfData = RSResult
End Function


Public Function GetWoCustName() As ADODB.Recordset


Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 cmdStr = " select distinct customer as id , customer as name from ib_wohistory A where A.ERPCREATEDATE>sysdate-365  order by 1"
 
Set RSResult = getStr(cmdStr)
Set GetWoCustName = RSResult
End Function




Public Function GetptData() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
 
'cmdStr = " select CUSTOMERSHORTNAME,CUSTOMERPTNo, productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by CUSTOMERSHORTNAME,CUSTOMERPTNo,productName,pfstaus,traystaus "
'
         
'cmdStr = " select CUSTOMERSHORTNAME,CUSTOMERPTNo, productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by CUSTOMERSHORTNAME,CUSTOMERPTNo,productName,pfstaus,traystaus "
'
cmdStr = " select productName, pfStaus, trayStaus, testNo from TBLSETPT where flag = 'Y' order by productName, pfstaus, traystaus "
              
         
Set RSResult = getStr(cmdStr)
Set GetptData = RSResult

End Function

Public Function GetNPIData() As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
cmdStr = "select  id  , CUSTOMERSHORTNAME , QtechPTNo   ,QtechPTNo2, CUSTOMERPTNo1   , CUSTOMERPTNo2  ,CUSTOMERPTNo3   , CUSTOMERPTNo4 , CUSTOMERPTNo5 , CUSTOMERPTNo6  , " & _
         " CUSTOMERDieQty , QtechDieQty, XiangSu, UsedArea, StruckStr1, StruckStr2, StruckStr3,ST_DATE,TT_DATE,PT_DATE ,PKG_TYPE,MARKING_CODE , P_E,MAPPING,RESIDUAL,MARKETLASTUPDATE_BY " & _
         " From TBLTsvNpiProduct where flag='Y' order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
               
Set RSResult = getStr(cmdStr)
Set GetNPIData = RSResult

End Function



Public Function GetGCTrayRptWla() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   
cmdStr = "select id,CustomerPT,productname,lotid,qty,to_char(createddate,'YYYY-MM-DD') cdate  from  TSV_GCTRAY_SetWLA where   flag='Y' order by id desc  "
               
Set RSResult = getStr(cmdStr)
Set GetGCTrayRptWla = RSResult

End Function




Public Function GetEBRData() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   
cmdStr = "select id, BATCHID ,EBRNumber,PT,Contact from  CUSTOMEREBRtbl where flag='Y' order by id desc "
               
Set RSResult = getStr(cmdStr)
Set GetEBRData = RSResult

End Function



Public Function GetPMCWOHeader(sqlTemp As String) As ADODB.Recordset
'查询Wo header

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   
cmdStr = sqlTemp
               
Set RSResult = getStr(cmdStr)
Set GetPMCWOHeader = RSResult

End Function

Public Function GetPMCWOLine(sqlTemp As String) As ADODB.Recordset
'查询Wo header

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   
cmdStr = sqlTemp
               
Set RSResult = getStr(cmdStr)
Set GetPMCWOLine = RSResult

End Function


Public Function GetBaoFeiData() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   
cmdStr = "select  id, typename,PutInDept,created_by, WaferID ,LotID ,GDDie,NGDie ,CUSTOMERPTNO, QTECHPTNO,ProductName,ERR_DATE  ,PutIn_DATE  " & _
         " from TBLTSVBaoFei where flag='Y'   order by id desc, LotID,WaferID"
               
Set RSResult = getStr(cmdStr)
Set GetBaoFeiData = RSResult

End Function


Public Sub SignBaoFeiBill(EmpIDTemp As Long, signNo As String, typeIDtemp As Integer, typeTemp As String)
'虫沮
Dim cmdStr As String
Dim cmdStr2 As String

Dim mailContent  As String
Dim signType As String
                  
 If typeIDtemp = 2 Then
 
cmdStr = "   Update TBLTSVBaoFei set BillStatus='2',Sign_By='" & signNo & "',Sign_date=sysdate,SignStatus='Y' Where id = " & EmpIDTemp & " "
mailContent = "厂内有" & typeTemp & "记录已通过质量1审核，请及时进系统审核， 谢谢！   请知晓"
signType = "BaoFeiSys_Sign2"
                       
 ElseIf typeIDtemp = 3 Then
 
 cmdStr = "   Update TBLTSVBaoFei set BillStatus='3',Tell_By='" & signNo & "',Tell_date=sysdate,TellStatus='Y' Where id = " & EmpIDTemp & " "
  
 mailContent = "厂内有" & typeTemp & "记录已通过质量2审核（即已通知客户），请及时进系统审核， 谢谢！   请知晓"
  signType = "BaoFeiSys_Sign3"
                       
  ElseIf typeIDtemp = 4 Then
 
 cmdStr = "   Update TBLTSVBaoFei set BillStatus='4',Market_By='" & signNo & "',Market_date=sysdate,MarketStatus='Y' Where id = " & EmpIDTemp & " "
 
 
 '2015-01-11 jiayun modify 市场部审核后，更新Sqlserver单据状态
 
 
  cmdStr2 = "   Update [erpdata].[dbo].[TSVtblBAOFEI] set BillStatus='4',Market_By='" & signNo & "',Market_date=getdate(),MarketStatus='Y' Where id = " & EmpIDTemp & " "
                                       
  AddSql2 (cmdStr2)
                       
 End If
                       
                       
 AddSql (cmdStr)
 
 Call MailDetailBaoFei(typeTemp & " 数据审核", signType, mailContent)
 

End Sub

Public Sub RejectBaoFeiBill(EmpIDTemp As Long, signNo As String, typeIDtemp As Integer, typeTemp As String, txtCommand As String)
'虫沮
Dim cmdStr As String
Dim mailContent  As String
Dim signType As String
Dim ToMailTemp As String

ToMailTemp = GetBaofeiRejectMailName(EmpIDTemp)

'select b.mailname from  TBLTSVBaoFei a , AutoMailList b where a.id=21 and b.systemname='BaoFeiSys_Reject'  and b.deptname=a.created_by

                  
 If typeIDtemp = 2 Then
 
cmdStr = "   Update TBLTSVBaoFei set BillStatus='5',Sign_By='" & signNo & "',Sign_date=sysdate,SignStatus='Y',  rejectReason ='" & txtCommand & "'  Where id = " & EmpIDTemp & " "
mailContent = "厂内有" & typeTemp & "记录没有通过质量1审核，请进系统查看， 谢谢！   请知晓"
signType = "BaoFeiSys_Reject"
                       
 ElseIf typeIDtemp = 3 Then
 
 cmdStr = "   Update TBLTSVBaoFei set BillStatus='5',Tell_By='" & signNo & "',Tell_date=sysdate,TellStatus='Y',  rejectReason ='" & txtCommand & "' Where id = " & EmpIDTemp & " "
  
 mailContent = "厂内有" & typeTemp & "记录没有通过质量2审核，请进系统查看， 谢谢！   请知晓"
  signType = "BaoFeiSys_Reject"
                       
  ElseIf typeIDtemp = 4 Then
 
 cmdStr = "   Update TBLTSVBaoFei set BillStatus='5',Market_By='" & signNo & "',Market_date=sysdate,MarketStatus='Y' ,  rejectReason ='" & txtCommand & "' Where id = " & EmpIDTemp & " "
  
                                       
                       
 End If
                       
                       
 AddSql (cmdStr)
 
 Call MailDetailBaoFeiRejectMail(typeTemp & " 数据审核", signType, mailContent, ToMailTemp)
 

End Sub




Public Function GetBaoFeiDataSign(typeId As Integer, typeTemp As String) As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   
 If typeId = 2 Then
 
cmdStr = "select  id ,typename,PutInDept, WaferID ,LotID ,GDDie,NGDie ,CUSTOMERPTNO, QTECHPTNO,ProductName,ERR_DATE  ,PutIn_DATE ,1 " & _
         " from TBLTSVBaoFei where flag='Y' and BillStatus='1'  and typename='" & typeTemp & "' order by PutIn_DATE desc, LotID,WaferID"
               
  ElseIf typeId = 3 Then
  
  cmdStr = "select  id ,typename,PutInDept, WaferID ,LotID ,GDDie,NGDie ,CUSTOMERPTNO, QTECHPTNO,ProductName,ERR_DATE  ,PutIn_DATE ,1 " & _
         " from TBLTSVBaoFei where flag='Y' and BillStatus='2' and typename='" & typeTemp & "' order by PutIn_DATE desc, LotID,WaferID"
         
 ElseIf typeId = 4 Then
  
  cmdStr = "select  id ,typename,PutInDept, WaferID ,LotID ,GDDie,NGDie ,CUSTOMERPTNO, QTECHPTNO,ProductName,ERR_DATE  ,PutIn_DATE ,1 " & _
         " from TBLTSVBaoFei where flag='Y' and BillStatus='3' and typename='" & typeTemp & "'  order by PutIn_DATE desc, LotID,WaferID"
End If

Set RSResult = getStr(cmdStr)
Set GetBaoFeiDataSign = RSResult

End Function



Public Function GetBaoFeiDataQuery(beginDateTemp As Date, endDateTemp As Date, lotIdTemp As String) As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   

If lotIdTemp = "" Then
  
  cmdStr = "select  id ,typename,PutInDept, created_by,WaferID ,LotID ,GDDie,NGDie ,CUSTOMERPTNO, QTECHPTNO,ProductName,ERR_DATE  ,PutIn_DATE , sign_by,sign_date,tell_by,tell_date,market_by,market_date ,Get_BaoFeiBillStatusNew(billstatus) as billstatus , rejectreason " & _
         " from TBLTSVBaoFei where created_date >=to_date('" & beginDateTemp & "','YYYY-MM-DD') and  created_date<=to_date('" & endDateTemp & "','YYYY-MM-DD')+1  order by PutIn_DATE desc, LotID,WaferID"
Else

  cmdStr = "select  id ,typename,PutInDept, created_by,WaferID ,LotID ,GDDie,NGDie ,CUSTOMERPTNO, QTECHPTNO,ProductName,ERR_DATE  ,PutIn_DATE , sign_by,sign_date,tell_by,tell_date,market_by,market_date ,Get_BaoFeiBillStatusNew(billstatus) as billstatus ,rejectreason  " & _
         " from TBLTSVBaoFei where  LotID='" & lotIdTemp & "' order by PutIn_DATE desc, LotID,WaferID"
         
End If

Set RSResult = getStr(cmdStr)
Set GetBaoFeiDataQuery = RSResult

End Function


Public Function GetMSLevel() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   



  cmdStr = "select id,ms_level,typename||numberofhours as name from  CUSTOMERMSLevelTBL where flag='Y' order by id"
         


Set RSResult = getStr(cmdStr)
Set GetMSLevel = RSResult

End Function

Public Function GetMPNAtri() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   



  cmdStr = " select id, LOC,PART,MarkingCodeFirst,LEAD_FREE,ECAT,MSL,TEMP,HALIDE_FREE,PBF_DIE_ATTACH,MPQ_QTY,PACKAGING_TYPE,PKG_GRP_CD,UL_LISTED_FLAG  from  CUSTOMERMPNAttributes where flag='Y' order by id "
         


Set RSResult = getStr(cmdStr)
Set GetMPNAtri = RSResult

End Function


Public Function GetONMarkingCode() As ADODB.Recordset


Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   



 ' cmdStr = " select  ID,MPNPART,WSGPART,MarkCodeFirst from  CUSTOMERMarkingCode  where flag='Y' order by id "
  
  'cmdStr = " select  ID,MPNPART,WSGPART from  CUSTOMERMarkingCode  where flag='Y' order by id "
         
  cmdStr = " select  ID,MPNPART,WSGPART,markcodefirst,ZX_REQUESTNUMBER,ZX_SHIPPER,ZX_SELFPICKUP,ZX_TRACKINGNO,ZX_SPECIALMARK from  CUSTOMERMarkingCode  where flag='Y' order by id desc "

Set RSResult = getStr(cmdStr)
Set GetONMarkingCode = RSResult

End Function


Public Function GetHTPTCross() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   



  cmdStr = " select id, NewOnOPNPart,FG_Legacy ,OPN_Site ,WDQ ,WDQ_Legacy,WDQ_Site,NewOnWSGPart ,WSG_Legacy ,WSG_Site ," & _
           " GRV ,GRV_Legacy,GRV_Site ,SWF ,swf_legacy,SWF_Site  from  CUSTOMERHTPartCross where flag='Y' order by id "
         


Set RSResult = getStr(cmdStr)
Set GetHTPTCross = RSResult

End Function


Public Function GetSpecialGRDataList() As ADODB.Recordset
'
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
          
         
 cmdStr = " select PO_NUM, PO_ITEM ,PREVIOUS_BATCH_ID ,PREVIOUS_MTRL_NUM , BATCH_ID , MTRL_NUM , MTRL_DESC  ," & _
 " MTRL_NUM_MTRLGRP , OUTPUT_QTY , CONSUMED_QTY, REJECT_QTY  , CURRENT_WAFER_QTY ,FILM_FRAME_QTY ," & _
 " OPTICAL_QUALITY , COUNTRY_OF_ASSEMBLY , OFFSHORE_ASM_COMPANY , ASM_CONTAINMENT_TYPE, DATE_CODE ," & _
 " ASM_CONV_ID , ASM_EXCR_ID , ASSEMBLY_FACILITY, COUNTRY_OF_TEST , OFFSHORE_TEST_COMPANY , TST_CONTAINMENT_TYPE ," & _
 " TST_PROGRAM_REV  ,  to_char(to_date(CREATED_DATE,'YYYY-MM-DD'),'MM/DD/YYYY') CREATED_DATE , CREATED_TIME ,DEL_NOTE    , AWB , WEIGHT , PACKAGEQTY  From SPEGRDETAILHISTORY  where  flag='Y' and to_char(create_date,'YYYY-MM-DD')=to_char(sysdate,'YYYY-MM-DD')"
         
                      
Set RSResult = getStr(cmdStr)
Set GetSpecialGRDataList = RSResult

End Function




Public Function GetNPIDataPrice() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   
cmdStr = "select  id  , CUSTOMERSHORTNAME , QtechPTNo   ,QtechPTNo2, CUSTOMERPTNo1   , CUSTOMERPTNo2  , " & _
         " CUSTOMERDieQty , XiangSu, fzFreeUSD, testFreeUSD, fzFreeRMB, testFreeRMB, nreFree, nreMethod, updatePrice2, updatePrice1  " & _
         " From TBLTsvNpiProduct where flag='Y' order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
               
Set RSResult = getStr(cmdStr)
Set GetNPIDataPrice = RSResult

End Function


Public Function GetPOPrice() As ADODB.Recordset
'查询表中数据

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
cmdStr = " select   ID ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE , PO_TYPE ,PT ,PeaceQty, QTY ,PRICE ,UNIT , " & _
"  FILENAME   from  TSV_MD_POPrice where flag='Y'  order by id desc "
               
Set RSResult = getStr(cmdStr)
Set GetPOPrice = RSResult

End Function


Public Function GetPOPriceModify(customerTemp As String, flagTemp As String) As ADODB.Recordset
'查询表中数据

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
cmdStr = " select   ID ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE , PO_TYPE ,PT ,PeaceQty, QTY ,BJ ,PRICE ,DIE_PRICE, UNIT , " & _
"  FILENAME  ,'' from  TSV_MD_POPrice where flag='Y' and  CUSTOMERSHORTNAME ='" & customerTemp & "' and check_flag = '" & flagTemp & "' order by id desc "
               
Set RSResult = getStr(cmdStr)
Set GetPOPriceModify = RSResult

End Function

Public Function GetPOPriceModify2(customerTemp As String, flagTemp As String) As ADODB.Recordset
'查询表中数据

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
cmdStr = " select   ID ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE , PO_TYPE ,PT ,peaceqty, QTY ,BJ ,PRICE ,DIE_PRICE ,UNIT , " & _
"  FILENAME  ,'' from  TSV_MD_POPrice where flag='Y'  and  CUSTOMERSHORTNAME ='" & customerTemp & "' and check_flag = '" & flagTemp & "'  order by id desc "
               
Set RSResult = getStr(cmdStr)
Set GetPOPriceModify2 = RSResult

End Function


Public Function Get37CTData(time1 As String, time2 As String) As ADODB.Recordset
'查询表中数据

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    

cmdStr = "select customername,mpn_desc,qtechptno,source_batch_id, qtech_created_date,erpcreatedate," & _
"max(mfgdate) outdate,shipdate,max(ct1) as ct1,max(ct2) as ct2,Get_LotCTTime(source_batch_id) from " & _
"( select distinct 'Semtech' as customername,b.mpn_desc,c.qtechptno,b.source_batch_id,b.qtech_created_date,d.erpcreatedate,a.mfgdate, " & _
" '' as shipdate,trunc(a.mfgdate-d.erpcreatedate,2) as ct1,trunc(a.mfgdate-b.qtech_created_date,2) as ct2, '' as holdtime " & _
"  from historymainline a, customeroitbl_test b,TBLTsvNpiProduct c , ib_wohistory d ,ib_waferlist e " & _
" where a.mfgdate >=to_date('" & time1 & "','YYYY-MM-DD') and a.mfgdate<to_date('" & time2 & "','YYYY-MM-DD')+1 " & _
"   and a.productname = '18X37025B000BR' and a.specname = '5272' and a.callbycdoname = 'WaferWIPMain' and a.cdoname = 'MoveInLot' " & _
"   and a.containername like '%-A%'  and b.customershortname='37' and b.mtrl_num=substr( a.containername,1,instr(a.containername,'-A')-1) " & _
"   and c.customershortname='37' and c.customerptno1=b.mpn_desc and e.waferid=b.mtrl_num and d.ordername=e.ordername" & _
"   and d.customer='37') x group by customername,mpn_desc,qtechptno,source_batch_id, qtech_created_date,erpcreatedate,holdtime"
   
   
   
   
               
Set RSResult = getStr(cmdStr)
Set Get37CTData = RSResult

End Function





Public Function GetPOPriceModify3(customerTemp As String) As ADODB.Recordset
'查询表中数据

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
cmdStr = " select   ID ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE , PO_TYPE ,PT , QTY ,PRICE ,UNIT , " & _
"  FILENAME  ,''  from  TSV_MD_POPrice where flag='Y'  and  PO_NUM ='" & customerTemp & "'order by id desc "
               
Set RSResult = getStr(cmdStr)
Set GetPOPriceModify3 = RSResult

End Function








Public Function GetNPIDataID(idTemp As Long) As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   
cmdStr = "select  id  , CUSTOMERSHORTNAME , QtechPTNo   ,QtechPTNo2, CUSTOMERPTNo1   , CUSTOMERPTNo2  ,CUSTOMERPTNo3,CUSTOMERPTNo4,CUSTOMERPTNo5,CUSTOMERPTNo6, " & _
         " CUSTOMERDieQty , QtechDieQty, XiangSu, UsedArea, StruckStr1, StruckStr2, StruckStr3,ST_DATE,TT_DATE,PT_DATE,pkg_type,residual, MARKING_CODE, MAPPING,MARKETLASTUPDATE_BY  " & _
         " From TBLTsvNpiProduct where id=" & idTemp & "  order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
               
Set RSResult = getStr(cmdStr)
Set GetNPIDataID = RSResult

End Function


Public Function GetNPIDataIDPrice(idTemp As Long) As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 'cmdStr = " select productName,pfStaus,trayStaus,testNo   from TBLSETPT  where flag='Y'  order by productName,pfstaus,traystaus "
   
cmdStr = "select  id  , fzFreeUSD,testFreeUSD,fzFreeRMB,testFreeRMB,nreFree,nreMethod,updatePrice2,updatePrice1 " & _
         " From TBLTsvNpiProduct where flag='Y' and id=" & idTemp & "   "
               
Set RSResult = getStr(cmdStr)
Set GetNPIDataIDPrice = RSResult

End Function



Public Function GetptOtherCustPT() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 cmdStr = "  select customercode,customerpt,qtechpt  from  TBLSETQtechPT where  flag='Y' order by  customercode,customerpt,qtechpt "

Set RSResult = getStr(cmdStr)
Set GetptOtherCustPT = RSResult
End Function




Public Function GetTrayData() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 cmdStr = "select FIELDNAME,FIELDVALUE,TrayType,OTHER from  TBLSETTray where flag='Y' order by  TrayType,FIELDNAME,FIELDVALUE "
         

Set RSResult = getStr(cmdStr)
Set GetTrayData = RSResult
End Function


Public Function GettestNo() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 cmdStr = "select productname,productnamenew,testedition  from  tblTestNo  where flag='Y' order by productname ,productnamenew "
         
Set RSResult = getStr(cmdStr)
Set GettestNo = RSResult
End Function

Public Function GetGCDieQty() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 cmdStr = "select CustomerName,CustomerPT,DieQty from  tblCustomerDieQty where flag='Y' order by  CustomerName,CustomerPT "
         

Set RSResult = getStr(cmdStr)
Set GetGCDieQty = RSResult
End Function




Public Function GettestNo2() As ADODB.Recordset
'查询GC MarkingCode

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 cmdStr = "select productname, productnamenew,testedition  ,FIELDNAME1,FIELDVALUE1,Remark1,FIELDNAME2,FIELDVALUE2,Remark2,FIELDNAME3,FIELDVALUE3,Remark3  from  tblTestNo2  where flag='Y' order by productname,productnamenew "
         

Set RSResult = getStr(cmdStr)
Set GettestNo2 = RSResult
End Function


Public Function GetAAMPNData(sqlTemp As String) As ADODB.Recordset


Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 cmdStr = sqlTemp
         

Set RSResult = getStr(cmdStr)
Set GetAAMPNData = RSResult
End Function

Public Function GetMDPODetail(sqlTemp As String) As ADODB.Recordset


Dim cmdStr As String
Dim RSResult As New ADODB.Recordset
    
 cmdStr = sqlTemp
         

Set RSResult = getStr(cmdStr)
Set GetMDPODetail = RSResult
End Function





Public Function GetMainItem() As ADODB.Recordset
'琩高干摸
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select tt.nameTemp id ,tt.valueTemp  typename from (select  a.testedition nameTemp  ,a.testedition valueTemp   from  tblTestNo a " & _
         " Union select b.testedition nameTemp ,b.testedition valueTemp  from  tblTestNo2 b ) TT order by tt.nameTemp"
         
         
Set RSResult = getStr(cmdStr)
Set GetMainItem = RSResult
End Function

Public Function GetMainItemProduct(productNameTemp As String) As ADODB.Recordset
'根据成品料号找测试版本号
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "    select testedition id ,testedition typename  from  tblTestNo where productnamenew='" & productNameTemp & "' and flag='Y' "
              
Set RSResult = getStr(cmdStr)
Set GetMainItemProduct = RSResult
End Function




Public Function MenuGrant(ByVal Frm As Form) As Boolean
Dim strSql          As String
Dim rs              As New ADODB.Recordset
Dim objMenu         As Object
Dim i               As Long

    MenuGrant = False

'    On Error GoTo ErrHandle
        strSql = "select MenuVale from tblSysMenuInfo where SysName='" & C_SysName & "'"
        
        If Cnn.State = 0 Then
        ConOracle
        End If
        
        rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not rs.EOF
            For Each objMenu In Frm.Controls
                If TYPENAME(objMenu) = "Menu" Then
                    If UCase(objMenu.name) = UCase(rs!MenuVale) Then
                        objMenu.Enabled = False
                        Exit For
                    End If
                End If
            Next
            rs.MoveNext
        Loop
        rs.Close
        
        strSql = "select a.MenuVale from tblSysMenuInfo a,tblGrantSysMenu b where a.SysName='" & C_SysName & "' " & _
                 " and a.id=b.id and b.username='" & gUserName & "'"
                 
        If Cnn.State = 0 Then
        ConOracle
        End If

        rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not rs.EOF
            If Len(rs!MenuVale) > 0 Then
                For Each objMenu In Frm.Controls
                    If TYPENAME(objMenu) = "Menu" Then
                        If UCase(objMenu.name) = UCase(rs!MenuVale) Then
                            objMenu.Enabled = True
                            Exit For
                        End If
                    End If
                Next
            End If
            rs.MoveNext
        Loop
        rs.Close

    
    MenuGrant = True
Exit Function
ErrHandle:
   MsgBox "MenuGrant菜单权限失败", vbInformation, "友情提示"
   
   
End Function

Public Function MailDetailBaoFei(ByVal Subject As String, ByVal MailType As String, ByVal mailContent As String) As Boolean
Dim JM              As New jmail.Message
Dim strBodyinfo     As String
Dim i               As Integer

Dim strSql      As String
Dim j As Integer
Dim rs          As New ADODB.Recordset

On Error GoTo ErrHandler
    MailDetailBaoFei = False
    

    
    '调用jmail
    JM.Charset = "gb2312"
    JM.Silent = False
    JM.Priority = 1
    JM.MailServerUserName = "Qtech_Report@qtechglobal.com" '帐号
    JM.MailServerPassWord = "qtech8" '密码
    JM.From = "Qtech_Report@qtechglobal.com"    '名字
    JM.FromName = "Qtech_Report"  '发件人姓名
    
'    '多人邮件区分
'    Recipients = Split(Recipient, ",")
'    For i = 0 To UBound(Recipients)
'        If InStr(1, Recipients(i), "@") <= 1 Or InStr(1, Recipients(i), "@") >= Len(Recipients(i)) Then
'            Exit Function
'        Else
'            JM.AddRecipient Recipients(i)
'        End If
'    Next i
'
'    '操送人
'    RecipientCCs = Split(RecipientCC, ",")
'    For i = 0 To UBound(RecipientCCs)
'        If InStr(1, RecipientCCs(i), "@") <= 1 Or InStr(1, RecipientCCs(i), "@") >= Len(RecipientCCs(i)) Then
'            Exit Function
'        Else
'            JM.AddRecipientCC RecipientCCs(i)
'        End If
'    Next i
    
    '2011-11-15  jiayunzhang modify 收件人，抄送人
    '收件人
'    strSql = "select mailname from  AutoMailList where SystemName='" & MailType & "' and SendType='To' and flag='Y' order by DeptName,MailName "
'    If Rs.State = adStateOpen Then Rs.Close
'    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
'    If Not Rs.EOF Then
'        For j = 1 To Rs.RecordCount
'            JM.AddRecipient Rs.fields(0).Value
'            Rs.MoveNext
'        Next
'    Else
'        JM.AddRecipient Recipient
'    End If
'    Rs.Close
'    Set Rs = Nothing
'
    
        
    strSql = "select mailname from  AutoMailList where SystemName='" & MailType & "' and SendType='To' and flag='Y' order by DeptName,MailName "
            
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Function
     
    For j = 1 To rs.RecordCount
        JM.AddRecipient rs.Fields(0).Value
        rs.MoveNext
    Next

    rs.Close
    Set rs = Nothing
    
     strSql = "select mailname from  AutoMailList where SystemName='" & MailType & "' and SendType='Cc' and flag='Y' order by DeptName,MailName "
            
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Function
     
    For j = 1 To rs.RecordCount
        JM.AddRecipientCC rs.Fields(0).Value
        rs.MoveNext
    Next

    rs.Close
    Set rs = Nothing
    
    
   
    
    JM.Subject = "HTKS AutoMail  " & Subject
    strBodyinfo = "您好！" & vbCrLf & "         " & mailContent

    JM.AppendText (strBodyinfo)
    
    MailDetailBaoFei = JM.Send("mail.qtechglobal.com")
    
ErrHandler:
    Set JM = Nothing
    Exit Function
End Function

Public Function MailDetailBaoFeiRejectMail(ByVal Subject As String, ByVal MailType As String, ByVal mailContent As String, toMailName As String) As Boolean
Dim JM              As New jmail.Message
Dim strBodyinfo     As String
Dim i               As Integer

Dim strSql      As String
Dim j As Integer
Dim rs          As New ADODB.Recordset

On Error GoTo ErrHandler
    MailDetailBaoFeiRejectMail = False
    

    
    '调用jmail
    JM.Charset = "gb2312"
    JM.Silent = False
    JM.Priority = 1
    JM.MailServerUserName = "Qtech_Report@qtechglobal.com" '帐号
    JM.MailServerPassWord = "qtech8" '密码
    JM.From = "Qtech_Report@qtechglobal.com"    '名字
    JM.FromName = "Qtech_Report"  '发件人姓名
    
'    '多人邮件区分
'    Recipients = Split(Recipient, ",")
'    For i = 0 To UBound(Recipients)
'        If InStr(1, Recipients(i), "@") <= 1 Or InStr(1, Recipients(i), "@") >= Len(Recipients(i)) Then
'            Exit Function
'        Else
'            JM.AddRecipient Recipients(i)
'        End If
'    Next i
'
'    '操送人
'    RecipientCCs = Split(RecipientCC, ",")
'    For i = 0 To UBound(RecipientCCs)
'        If InStr(1, RecipientCCs(i), "@") <= 1 Or InStr(1, RecipientCCs(i), "@") >= Len(RecipientCCs(i)) Then
'            Exit Function
'        Else
'            JM.AddRecipientCC RecipientCCs(i)
'        End If
'    Next i
    
    '2011-11-15  jiayunzhang modify 收件人，抄送人
'    '收件人
'    strSql = "select mailname from  AutoMailList where SystemName='" & MailType & "' and SendType='To' and flag='Y' order by DeptName,MailName "
'    If Rs.State = adStateOpen Then Rs.Close
'    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
'    If Not Rs.EOF Then
'        For j = 1 To Rs.RecordCount
'            JM.AddRecipient Rs.fields(0).Value
'            Rs.MoveNext
'        Next
'    Else
'        JM.AddRecipient Recipient
'    End If
'    Rs.Close
'    Set Rs = Nothing
    
    
      JM.AddRecipient toMailName
      
     strSql = "select mailname from  AutoMailList where SystemName='BaoFeiSys_Reject' and SendType='Cc' and flag='Y' order by DeptName,MailName "
            
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Function
     
    For j = 1 To rs.RecordCount
        JM.AddRecipientCC rs.Fields(0).Value
        rs.MoveNext
    Next

    rs.Close
    Set rs = Nothing
      
      
      
      
   
    
    JM.Subject = "HTKS AutoMail  " & Subject
    strBodyinfo = "您好！" & vbCrLf & "         " & mailContent

    JM.AppendText (strBodyinfo)
    
    MailDetailBaoFeiRejectMail = JM.Send("mail.qtechglobal.com")
    
ErrHandler:
    Set JM = Nothing
    Exit Function
End Function



'1
Public Function MailDetail(ByVal Subject As String, ByVal Recipient As String, ByVal Attachment As String, ByVal RecipientCC As String) As Boolean
Dim JM              As New jmail.Message
Dim Recipients()    As String
Dim RecipientCCs()    As String
Dim strBodyinfo     As String
Dim i               As Integer

Dim strSql      As String
Dim j As Integer
Dim rs          As New ADODB.Recordset
Dim RsD         As New ADODB.Recordset


On Error GoTo ErrHandler
    MailDetail = False
    

   JM.Charset = "gb2312"
    JM.Silent = False
    JM.Priority = 1
   JM.MailServerUserName = "sqladmin" '帐号
    JM.MailServerPassWord = "ksitadmin" '密码
    JM.From = "sqladmin@htkjks.com"    '名字
    JM.FromName = "sqladmin"  '发件人姓名
    
    '收件人
    strSql = "select mailname from  AutoMailList where SystemName='GC_Dev_Report' and SendType='To' and flag='Y' order by DeptName,MailName "
    If rs.State = adStateOpen Then rs.Close
    
    If Cnn.State = 0 Then
    ConOracle
    End If
    
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        For j = 1 To rs.RecordCount
            JM.AddRecipient rs.Fields(0).Value
            rs.MoveNext
        Next
    Else
        JM.AddRecipient Recipient
    End If
    rs.Close
    Set rs = Nothing
    '抄送人
    strSql = "select mailname from  AutoMailList where SystemName='GC_Dev_Report' and SendType='Cc' and flag='Y' order by DeptName,MailName "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        For j = 1 To rs.RecordCount
            JM.AddRecipientCC rs.Fields(0).Value
            rs.MoveNext
        Next
    Else
        JM.AddRecipientCC RecipientCC
    End If
    
    rs.Close
    Set rs = Nothing
    
    '附件
    If Dir(Attachment, vbNormal Or vbArchive) = "" Then
        Exit Function
    Else
        JM.AddAttachment Attachment
    End If
    
    JM.Subject = "HTKS AutoMail  " & Subject
    strBodyinfo = "尊敬的客户，您好！" & vbCrLf & "        附件为HTKS " & Format(g_Date, "YYYY-MM-DD") & "的GC 发货报表 ，详见附件 。" & vbCrLf & "请查收 "

    JM.AppendText (strBodyinfo)
    
    MailDetail = JM.Send("mail.htkjks.com")
    
ErrHandler:
    Set JM = Nothing
    Exit Function
End Function

Public Function MailDetail_TW(ByVal Subject As String, _
                              ByVal Recipient As String, _
                              ByVal Attachment As String, _
                              ByVal RecipientCC As String) As Boolean



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

    On Error GoTo ErrHandler

    MailDetail_TW = False

'    JM.Charset = "gb2312"
'    JM.Silent = False
'    JM.Priority = 1
'    JM.MailServerUserName = "ks015918" '帐号
'    JM.MailServerPassWord = "ks123456" '密码
'    JM.From = "ks015918@ht-tech.com"    '名字
'    JM.FromName = "sqladmin"  '发件人姓名
    
    JM.Charset = "gb2312"
    JM.Silent = False
    JM.Priority = 1
   JM.MailServerUserName = "sqladmin" '帐号
    JM.MailServerPassWord = "ksitadmin" '密码
    JM.From = "sqladmin@ht-tech.com"    '名字
    JM.FromName = "sqladmin"  '发件人姓名
    
    '收件人
    JM.AddRecipient Recipient
 
    '抄送人
    JM.AddRecipientCC RecipientCC
    JM.AddRecipientCC "jian.pan_ks@ht-tech.com"

    '附件
    If Attachment <> "" Then
        If Dir(Attachment, vbNormal Or vbArchive) = "" Then
            Exit Function
        Else
            JM.AddAttachment Attachment

        End If

    End If
    
    JM.Subject = "HTKS AutoMail  " & Subject
    strBodyinfo = "请查收 "

    JM.AppendText (strBodyinfo)
    
    MailDetail_TW = JM.Send("mail.ht-tech.com")
    
ErrHandler:
    Set JM = Nothing
    Exit Function

End Function



Public Function MailDetailSX(ByVal Subject As String, ByVal Recipient As String, ByVal Attachment As String, ByVal RecipientCC As String) As Boolean
Dim JM              As New jmail.Message
Dim Recipients()    As String
Dim RecipientCCs()    As String
Dim strBodyinfo     As String
Dim i               As Integer

Dim strSql      As String
Dim j As Integer
Dim rs          As New ADODB.Recordset
Dim RsD         As New ADODB.Recordset


On Error GoTo ErrHandler
    MailDetailSX = False
    

    
    '调用jmail
    JM.Charset = "gb2312"
    JM.Silent = False
    JM.Priority = 1
    JM.MailServerUserName = "Qtech_Report@qtechglobal.com" '帐号
    JM.MailServerPassWord = "qtech8" '密码
    JM.From = "Qtech_Report@qtechglobal.com"    '名字
    JM.FromName = "Qtech_Report"  '发件人姓名
    
    '收件人
    strSql = "select mailname from  AutoMailList where SystemName='SX_Dev_Report' and SendType='To' and flag='Y' order by DeptName,MailName "
    If rs.State = adStateOpen Then rs.Close
    
    If Cnn.State = 0 Then
    ConOracle
    End If
    
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        For j = 1 To rs.RecordCount
            JM.AddRecipient rs.Fields(0).Value
            rs.MoveNext
        Next
    Else
        JM.AddRecipient Recipient
    End If
    rs.Close
    Set rs = Nothing
    '抄送人
    strSql = "select mailname from  AutoMailList where SystemName='SX_Dev_Report' and SendType='Cc' and flag='Y' order by DeptName,MailName "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        For j = 1 To rs.RecordCount
            JM.AddRecipientCC rs.Fields(0).Value
            rs.MoveNext
        Next
    Else
        JM.AddRecipientCC RecipientCC
    End If
    
    rs.Close
    Set rs = Nothing
    
    
    '附件
    If Dir(Attachment, vbNormal Or vbArchive) = "" Then
        Exit Function
    Else
        JM.AddAttachment Attachment
    End If
    
    JM.Subject = "HTKS AutoMail  " & Subject
    strBodyinfo = "尊敬的客户，您好！" & vbCrLf & "        附件为HTKS " & Format(g_Date, "YYYY-MM-DD") & "的SX 发货报表 ，详见附件 。" & vbCrLf & "请查收 "

    JM.AppendText (strBodyinfo)
    
    MailDetailSX = JM.Send("mail.qtechglobal.com")
    
ErrHandler:
    Set JM = Nothing
    Exit Function
End Function

'biao
Public Function MailDetail56(ByVal Subject As String, ByVal Recipient As String, ByVal Attachment As String, ByVal RecipientCC As String) As Boolean
Dim JM              As New jmail.Message
Dim Recipients()    As String
Dim RecipientCCs()    As String
Dim strBodyinfo     As String
Dim i               As Integer

Dim strSql      As String
Dim j As Integer
Dim rs          As New ADODB.Recordset
Dim RsD         As New ADODB.Recordset


On Error GoTo ErrHandler
    MailDetail56X = False
    

    
    '调用jmail
    JM.Charset = "gb2312"
    JM.Silent = False
    JM.Priority = 1
    JM.MailServerUserName = "Qtech_Report@qtechglobal.com" '帐号
    JM.MailServerPassWord = "qtech8" '密码
    JM.From = "Qtech_Report@qtechglobal.com"    '名字
    JM.FromName = "Qtech_Report"  '发件人姓名
    
    '收件人
    strSql = "select mailname from  AutoMailList where SystemName='SX_Dev_Report' and SendType='To' and flag='Y' order by DeptName,MailName "
    If rs.State = adStateOpen Then rs.Close
    
    If Cnn.State = 0 Then
    ConOracle
    End If
    
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        For j = 1 To rs.RecordCount
            JM.AddRecipient rs.Fields(0).Value
            rs.MoveNext
        Next
    Else
        JM.AddRecipient Recipient
    End If
    rs.Close
    Set rs = Nothing
    '抄送人
    strSql = "select mailname from  AutoMailList where SystemName='SX_Dev_Report' and SendType='Cc' and flag='Y' order by DeptName,MailName "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        For j = 1 To rs.RecordCount
            JM.AddRecipientCC rs.Fields(0).Value
            rs.MoveNext
        Next
    Else
        JM.AddRecipientCC RecipientCC
    End If
    
    rs.Close
    Set rs = Nothing
    
    
    '附件
    If Dir(Attachment, vbNormal Or vbArchive) = "" Then
        Exit Function
    Else
        JM.AddAttachment Attachment
    End If
    
    JM.Subject = "HTKS AutoMail  " & Subject
    strBodyinfo = "尊敬的客户，您好！" & vbCrLf & "        附件为HTKS " & Format(g_Date, "YYYY-MM-DD") & "的SX 发货报表 ，详见附件 。" & vbCrLf & "请查收 "

    JM.AppendText (strBodyinfo)
    
    MailDetail56 = JM.Send("mail.qtechglobal.com")
    
ErrHandler:
    Set JM = Nothing
    Exit Function
End Function



Public Function MailDetailHD(ByVal Subject As String, ByVal Recipient As String, ByVal Attachment As String, ByVal RecipientCC As String) As Boolean
Dim JM              As New jmail.Message
Dim Recipients()    As String
Dim RecipientCCs()    As String
Dim strBodyinfo     As String
Dim i               As Integer

Dim strSql      As String
Dim j As Integer
Dim rs          As New ADODB.Recordset
Dim RsD         As New ADODB.Recordset


On Error GoTo ErrHandler
    MailDetailHD = False
    

    
    '调用jmail
    JM.Charset = "gb2312"
    JM.Silent = False
    JM.Priority = 1
    JM.MailServerUserName = "Qtech_Report@qtechglobal.com" '帐号
    JM.MailServerPassWord = "qtech8" '密码
    JM.From = "Qtech_Report@qtechglobal.com"    '名字
    JM.FromName = "Qtech_Report"  '发件人姓名
    
    '收件人
    strSql = "select mailname from  AutoMailList where SystemName='HD_Dev_Report' and SendType='To' and flag='Y' order by DeptName,MailName "
    If rs.State = adStateOpen Then rs.Close
    
    If Cnn.State = 0 Then
    ConOracle
    End If
    
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        For j = 1 To rs.RecordCount
            JM.AddRecipient rs.Fields(0).Value
            rs.MoveNext
        Next
    Else
        JM.AddRecipient Recipient
    End If
    rs.Close
    Set rs = Nothing
    '抄送人
    strSql = "select mailname from  AutoMailList where SystemName='HD_Dev_Report' and SendType='Cc' and flag='Y' order by DeptName,MailName "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        For j = 1 To rs.RecordCount
            JM.AddRecipientCC rs.Fields(0).Value
            rs.MoveNext
        Next
    Else
        JM.AddRecipientCC RecipientCC
    End If
    
    rs.Close
    Set rs = Nothing
    
    
    '附件
    If Dir(Attachment, vbNormal Or vbArchive) = "" Then
        Exit Function
    Else
        JM.AddAttachment Attachment
    End If
    
    JM.Subject = "HTKS AutoMail  " & Subject
    strBodyinfo = "尊敬的客户，您好！" & vbCrLf & "        附件为HTKS " & Format(g_Date, "YYYY-MM-DD") & "的HD 发货报表 ，详见附件 。" & vbCrLf & "请查收 "

    JM.AppendText (strBodyinfo)
    
    MailDetailHD = JM.Send("mail.qtechglobal.com")
    
ErrHandler:
    Set JM = Nothing
    Exit Function
End Function


'2014-04-09 add
Public Function GetWODataCode() As String
'莉ビ叫筁掸计
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select to_char(sysdate, 'yyyy') || to_char(sysdate, 'WW') AS DATA_CODE1  from dual "
        
     
RSResult = getStr2(cmdStr)
GetWODataCode = RSResult
End Function


'2014-04-09 add
Public Function GetGCOutRpt_Ver(waferIdTemp As String, ptTemp As String, verTemp As String) As String
'莉ビ叫筁掸计
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select  Get_GCOutRpt('" & waferIdTemp & "','" & ptTemp & "','" & verTemp & "')  as vertemp  from dual "

RSResult = getStr2(cmdStr)
GetGCOutRpt_Ver = RSResult
End Function


'2014-04-09 add
Public Function GetDateToCode(timeTemp As String) As String
'莉ビ叫筁掸计
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select to_char(to_date('" & timeTemp & "','YYYY-MM-DD'),'YYWW') as dcode from dual "



RSResult = getStr2(cmdStr)
GetDateToCode = RSResult
End Function


Public Function GetNPICustomerPt(productTemp As String) As String
'莉ビ叫筁掸计
Dim cmdStr As String
Dim RSResult As String
productTemp = Trim(Replace(Replace(productTemp, Chr(13), ""), Chr(10), ""))
cmdStr = "select a.customerptno1 from  TBLTsvNpiProduct a where a.qtechptno2='" & productTemp & "' "


RSResult = getStr2(cmdStr)
GetNPICustomerPt = RSResult
End Function


Public Function GetNPICustomerHTPt(productTemp As String) As String
'莉ビ叫筁掸计
Dim cmdStr As String
Dim RSResult As String
productTemp = Trim(Replace(Replace(productTemp, Chr(13), ""), Chr(10), ""))
cmdStr = "select  a.qtechptno  from  TBLTsvNpiProduct a where a.qtechptno2='" & productTemp & "' "


RSResult = getStr2(cmdStr)
GetNPICustomerHTPt = RSResult
End Function



Public Function GetOILotQty(productTemp As String) As Long


'莉ビ叫筁掸计
productTemp = Trim(Replace(Replace(productTemp, Chr(13), ""), Chr(10), ""))
Dim cmdStr As String
Dim RSResult As Long


cmdStr = "select count(a.substrateid) ID from mappingdatatest a where a.lotid='" & productTemp & "' "
     
RSResult = GetSeq(cmdStr)
GetOILotQty = RSResult

End Function



Public Function GetLotInFirstTime(lotIdTemp As String) As String
'莉ビ叫筁掸计
Dim cmdStr As String
Dim RSResult As String

cmdStr = "select to_char(min(a.qtech_created_date),'YYYY-MM-DD') from customeroitbl_test a where a.source_batch_id='" & lotIdTemp & "' "


RSResult = getStr2(cmdStr)
GetLotInFirstTime = RSResult
End Function





Public Function GetOracleCustomerName(factoryTemp As String) As ADODB.Recordset
'客户代码
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

'cmdStr = "select distinct b.productname PID,b.productname from product a ,PRODUCTBASE b" & _
'" where a.productbaseid=b.productbaseid and b.objectcategory='PN' and a.objecttype='PN' and b.productname like '18%' order by b.productname"
'

cmdStr = "select b.customerid as pid from TSV_MDRpt_Work_type a, TSV_MDRpt_Customer_type b  Where b.workid = a.id and b.flag = 'Y' and a.flag = 'Y'  and a.workname='" & factoryTemp & "' order by b.customerid"

Set RSResult = getStr(cmdStr)
Set GetOracleCustomerName = RSResult
End Function

Public Sub Modifywo(Userid As String, id As Long, wafertemp As WoWafer)
Dim sqlTemp As String

Dim sqlTemp1 As String


 sqlTemp = " insert into customeroitbl_test (id,source_batch_id,po_num,mtrl_num, mpn_desc,fab_conv_id,test_site,ship_site,flag,qtech_created_by,qtech_created_date, customershortname ) " & _
          "  (select '" & id & "', source_batch_id,'" & wafertemp.PO & "',mtrl_num, '" & wafertemp.device & "',fab_conv_id,test_site,ship_site,flag,'" & Userid & "',sysdate,customershortname " & _
         "   from customeroitbl_test c  where c.customershortname = '" & wafertemp.CustName & "' and c.source_batch_id = '" & wafertemp.lot & "' and  c.id =  (select max(id) from customeroitbl_test where customershortname = '" & wafertemp.CustName & "' and  source_batch_id = '" & wafertemp.lot & "')) "
         
sqlTemp1 = " insert into erpbase..tblCustomerOI (id,source_batch_id,po_num,mtrl_num, mpn_desc,fab_conv_id,test_site,ship_site,flag,qtech_created_by,qtech_created_date,qtech_lastupdate_date,customershortname ) " & _
          "  (select '" & id & "', source_batch_id,'" & wafertemp.PO & "',mtrl_num, '" & wafertemp.device & "',fab_conv_id,test_site,ship_site,flag,'" & Userid & "',GETDATE(),GETDATE(),customershortname " & _
         "   from erpbase .. tblCustomerOI c  where c.customershortname = '" & wafertemp.CustName & "' and c.source_batch_id = '" & wafertemp.lot & "' and  c.id =  (select max(id) from erpbase..tblCustomerOI where customershortname = '" & wafertemp.CustName & "' and  source_batch_id = '" & wafertemp.lot & "')) "
           
 AddSql (sqlTemp)
 AddSql2 (sqlTemp1)
 
 
End Sub


Public Sub modifwafer(Userid As String, id As Long, wafertemp As WoWafer)

    
 sqlTemp = " update mappingdatatest set filename = '" & id & "' , qtech_created_by = '" & Userid & "' , passbincount = '" & wafertemp.gooddie & "', failbincount = '" & wafertemp.ngdie & "',  qtech_lastupdate_date = sysdate " & _
           " where substrateid = '" & wafertemp.WAFER & "' and lotid = '" & wafertemp.lot & "' and customershortname = '" & wafertemp.CustName & "'"
           
     
 sqlTemp1 = " update erpbase..tblmappingData set filename = '" & id & "' , qtech_created_by = '" & Userid & "' ,passbincount = '" & wafertemp.gooddie & "', failbincount = '" & wafertemp.ngdie & "',qtech_lastupdate_date = GETDATE() " & _
           " where substrateid = '" & wafertemp.WAFER & "' and lotid = '" & wafertemp.lot & "' and customershortname = '" & wafertemp.CustName & "'"
           
 AddSql (sqlTemp)
 AddSql2 (sqlTemp1)

End Sub


Public Function Getwo_wafer(custTemp As String, lottemp As String) As ADODB.Recordset
'查询表中数据

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


    
cmdStr = " select  c.customershortname,c.source_batch_id, d.substrateid,d.wafer_id,c.po_num,c.mpn_desc,d.passbincount,d.failbincount,d.filename,'' " & _
        "  from customeroitbl_test c,mappingdatatest d where c.customershortname = '" & custTemp & "' and d.filename = to_char(c.id)  and c.source_batch_id = '" & lottemp & "' "
               
Set RSResult = getStr(cmdStr)
Set Getwo_wafer = RSResult

End Function

Public Sub ApprovalPO(nPoPriceTemp As POPrice)

Dim sqlTemp As String
Dim sqlTemp1 As String
Dim userName As String


Dim id As Long

id = nPoPriceTemp.id
userName = nPoPriceTemp.CreateBy

 
    sqlTemp = "insert into erptemp..tblBB_CSRPO values('" & nPoPriceTemp.CUSTOMERSHORTNAME & "','" & nPoPriceTemp.PONo & "','10','" & nPoPriceTemp.bj & "',  " & _
  " '" & nPoPriceTemp.pt & "','" & nPoPriceTemp.peaseQty & "','" & nPoPriceTemp.QTY & "','" & nPoPriceTemp.Price & "','" & nPoPriceTemp.DIE_PRICE & "' ,'" & nPoPriceTemp.unit & "','')  "
 
  
 AddSql2 (sqlTemp)
 
 
  sqlTemp1 = " Update TSV_MD_POPrice set check_flag = 'Y' where  id=" & id & " and flag='Y' "

              
 AddSql (sqlTemp1)
 

End Sub

Public Function GetMARK1(wafertemp) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


cmdStr = " select c.productid,to_char(wm_concat(c.substrateid)),to_char(count(*)),''" & _
        "  from mappingdatatest c " & _
       "   where c.customershortname in ('SX', 'HJ', 'TJ003', 'JS140', 'BJ153','GC') " & _
      "    and c.qtech_created_date > sysdate - 100 " & _
      "    and c.substrateid not like '%+' " & _
      "    group by c.productid " & _
      "    having count(*) >1 "

Set RSResult = getStr(cmdStr)

Set GetMARK1 = RSResult
End Function

Public Function GetMARK2(wafertemp As String, txtPN As String) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


cmdStr = " select to_char(length(replace(c.productid,'_',''))),b.mpn_desc,to_char(nvl(a.marking_code,0)), to_char(wm_concat(distinct c.lotid))" & _
         " from mappingdatatest c " & _
"         inner join customeroitbl_test b " & _
"         on to_char(b.id) = c.filename " & _
"     left join tbltsvnpiproduct a " & _
"     on a.customerptno1 = b.mpn_desc " & _
"    Where c.CUSTOMERSHORTNAME " & _
"    in ('SX', 'HJ', 'TJ003', 'JS140', 'BJ153','GC') " & _
"    and c.qtech_created_date > sysdate - 100 " & _
"    and c.substrateid = '" & wafertemp & "' and a.qtechptno2 = '" & txtPN & "'     " & _
 "   and  nvl(length(replace(c.productid,'_','')),0) <> nvl(a.marking_code,0) " & _
"    group by length(replace(c.productid,'_','')),b.mpn_desc,a.marking_code "

Set RSResult = getStr(cmdStr)

Set GetMARK2 = RSResult
End Function

Public Function GetMARK3(marktemp) As ADODB.Recordset

Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


'cmdStr = " select count(aa.productid ) from  mappingdatatest aa " & _
'         "  where aa.customershortname in ('SX', 'HJ', 'TJ003', 'JS140', 'BJ153') " & _
'        " and aa.productid = '" & marktemp & "' " & _
'        "  having count(aa.productid) >1 "
'
cmdStr = " select count(aa.productid ) from  mappingdatatest aa " & _
         "  where aa.customershortname in ('SX', 'HJ', 'TJ003', 'JS140', 'BJ153') " & _
        " and aa.productid = '" & marktemp & "' and aa.substrateid not like '%+%' " & _
        "  having count(aa.productid) >1 "

Set RSResult = getStr(cmdStr)

Set GetMARK3 = RSResult
End Function

Public Function Getcustpart(wafertemp As String) As String

Dim cmdStr As String
Dim RSResult  As String

cmdStr = " select b.fab_conv_id from mappingdatatest a,customeroitbl_test b where to_char(b.id) = a.filename and a.substrateid = '" & wafertemp & "'"
RSResult = getStr2(cmdStr)

Getcustpart = RSResult
End Function

Public Function Get37Bonded(wafertemp As String) As String

Dim cmdStr As String
Dim RSResult  As String


cmdStr = " select b.jobno from mappingdatatest a,customeroitbl_test b where to_char(b.id) = a.filename and a.substrateid = '" & wafertemp & "'"


RSResult = getStr2(cmdStr)



 Get37Bonded = RSResult
End Function

Public Sub AddTmpLot(strLotID As String, strWOID As String)

    Dim cmdStr As String

    On Error GoTo DealError

    Cnn.BeginTrans
                 
    cmdStr = "insert into WAFERDETAILTMP(LOTID, WOID) values('" & strLotID & "', '" & strWOID & "')"
    AddSql (cmdStr)
 
    Cnn.CommitTrans

    Exit Sub
 
DealError:

    Cnn.RollbackTrans

End Sub

Public Sub DelTmpLot()

    Dim cmdStr As String
    Dim cmdStr2 As String

    On Error GoTo DealError

    Cnn.BeginTrans
                  
    cmdStr = "delete from WAFERDETAILTMP"
  
    AddSql (cmdStr)

    Cnn.CommitTrans

    Exit Sub
 
DealError:

    Cnn.RollbackTrans

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       SentMes
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-26-10:16:33
'
' Parameters :       Subject (String)
'                    SentText (String)
'                    Recipient() (String)
'                    Attachment (String)
'                    RecipientCC() (String)
'--------------------------------------------------------------------------------
Public Function SentMes(Subject As String, SentText As String, Recipient() As String, Attachment As String, RecipientCC() As String) As Boolean

If gUserName = "07885" Then
    SentMes = True
    Exit Function
End If

    Dim JM As Object

    Set JM = CreateObject("JMAIL.Message")
    
    Dim Recipients()   As String

    Dim RecipientCCs() As String

    Dim strBodyinfo    As String

    Dim i              As Integer

    Dim strSql         As String

    Dim j              As Integer

    Dim rs             As New ADODB.Recordset

    Dim RsD            As New ADODB.Recordset

    On Error GoTo ErrHandler

    SentMes = False

    JM.Charset = "gb2312"
    JM.Silent = False
    JM.Priority = 1
    JM.MailServerUserName = "sqladmin" '帐号
    JM.MailServerPassWord = "ksitadmin" '密码
    JM.From = "sqladmin@ht-tech.com"    '名字
    JM.FromName = "sqladmin"  '发件人姓名
    
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
    
    '附件
    If Attachment <> "" Then
        If Dir(Attachment, vbNormal Or vbArchive) = "" Then
            Exit Function
        Else
            JM.AddAttachment Attachment

        End If

    End If
    
    JM.Subject = Subject
    JM.AppendText SentText
    
    SentMes = JM.Send("mail.ht-tech.com")
    
ErrHandler:
    Set JM = Nothing
    Exit Function

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckPOTerms
' Description:       检查客户代码和客户机种
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-26-10:16:46
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function CheckPOTerms(strCustCode As String, strCustPN As String) As Boolean
Dim strSql As String
CheckPOTerms = False
strSql = "select * from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and customerptno1 = '" & strCustPN & "'"
If Get_OracleCnt(strSql) = 0 Then
    MsgBox "NPI未维护该客户代码和该客户机种的对照关系,请确认是否输入有误", vbCritical, "提示"
    Exit Function
End If
CheckPOTerms = True
End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       getScanReels
' Description:       获取已扫描的卷盘数量
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-11:24:33
'
' Parameters :       strDN (String)
'--------------------------------------------------------------------------------
Public Function getScanReels(strDN As String) As Long
Dim strSql As String

strSql = "select max(seq) from PACKING_DETAILED where DN_NUM = '" & strDN & "' "
getScanReels = Get_OracleNo(strSql)
End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       getScanQty
' Description:       获取已扫描的数量
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-11:24:33
'
' Parameters :       strDN (String)
'--------------------------------------------------------------------------------
Public Function getScanQty(strDN As String) As Long
Dim strSql As String

strSql = "select sum(qty) from PACKING_DETAILED where DN_NUM = '" & strDN & "' "
getScanQty = Get_OracleNo(strSql)
End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       getlastScanReelID
' Description:       获取最后一个卷盘的REELID
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-10:58:39
'
' Parameters :       strDN (String)
'--------------------------------------------------------------------------------
Public Function getLastScanReelID(strDN As String) As String
Dim strSql As String
Dim lMaxSeq As Long

strSql = "select max(seq) from PACKING_DETAILED where DN_NUM = '" & strDN & "' "
lMaxSeq = Get_OracleNo(strSql)

If lMaxSeq <> 0 Then
    strSql = "select trayid from PACKING_DETAILED where DN_NUM = '" & strDN & "'  and seq = '" & lMaxSeq & "'"
    getLastScanReelID = Get_OracleStr(strSql)
Else
    getLastScanReelID = ""
End If

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       getlastScanReelID
' Description:       获取最后一个卷盘的JOBID
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-10:58:39
'
' Parameters :       strDN (String)
'--------------------------------------------------------------------------------
Public Function getLastScanJOBID(strDN As String) As String
Dim strSql As String
Dim lMaxSeq As Long

strSql = "select max(seq) from PACKING_DETAILED where DN_NUM = '" & strDN & "' "
lMaxSeq = Get_OracleNo(strSql)

If lMaxSeq <> 0 Then
    strSql = "select job_id from PACKING_DETAILED where DN_NUM = '" & strDN & "'  and seq = '" & lMaxSeq & "'"
    getLastScanJOBID = Get_OracleStr(strSql)
Else
    getLastScanJOBID = ""
End If

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       getlastScanReelID
' Description:       获取最后一个卷盘的MPN
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-10:58:39
'
' Parameters :       strDN (String)
'--------------------------------------------------------------------------------
Public Function getLastScanMPN(strDN As String) As String
Dim strSql As String
Dim lMaxSeq As Long

strSql = "select max(seq) from PACKING_DETAILED where DN_NUM = '" & strDN & "' "
lMaxSeq = Get_OracleNo(strSql)

If lMaxSeq <> 0 Then
    strSql = "select customer_device from PACKING_DETAILED where DN_NUM = '" & strDN & "'  and seq = '" & lMaxSeq & "'"
    getLastScanMPN = Get_OracleStr(strSql)
Else
    getLastScanMPN = ""
End If

End Function


'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       getLastInBoxNO
' Description:       获取最后一个卷盘所属内盒
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-12:30:48
'
' Parameters :       strDN (String)
'                    strReelID (String)
'--------------------------------------------------------------------------------
Public Function getLastInBoxNO(strDN As String, strReelID As String) As Integer
Dim strSql As String

strSql = "select inbox_num from PACKING_DETAILED where DN_NUM = '" & strDN & "' and trayid = '" & strReelID & "' "

getLastInBoxNO = Get_OracleNo(strSql)

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       getLastInBoxReels
' Description:       获取最后一个卷盘所属内盒的当前卷盘数量
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-12:34:59
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function getLastInBoxReels(strDN As String, strReelID) As Integer
Dim strSql As String
Dim strInBoxNO As String
Dim strOutBoxNO As String

strSql = "select inbox_num from PACKING_DETAILED where DN_NUM = '" & strDN & "' and trayid = '" & strReelID & "' "
strInBoxNO = Get_OracleStr(strSql)

strSql = "select outbox_num from PACKING_DETAILED where DN_NUM = '" & strDN & "' and trayid = '" & strReelID & "' "
strOutBoxNO = Get_OracleStr(strSql)

strSql = "select count(1) from PACKING_DETAILED where DN_NUM = '" & strDN & "' and outbox_num = '" & strOutBoxNO & "' and inbox_num = '" & strInBoxNO & "'"
getLastInBoxReels = Get_OracleNo(strSql)
End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       getReelID_ONELOT
' Description:       获取卷盘的JOBID
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-13:25:24
'
' Parameters :       strDN (String)
'                    strReelID (String)
'--------------------------------------------------------------------------------
Public Function getReel_JOBID_ONELOT(strReelID As String) As String
Dim strSql As String
Dim strRes As String

strSql = "select KEY_VALUE from erpdata..tblErpInStockDetailInfo a where CHARINDEX('" & strReelID & "',a.KEY_VALUE) > 0 and a.KEY_NAME = 'CONTAINER_NAME' AND a.KEY_TYPE = 'T'"
strRes = Get_SqlStr(strSql)

getReel_JOBID_ONELOT = Mid(strRes, InStr(strRes, "|") + 1)

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       checkJobID_ONELOT
' Description:       检查JOBID是否属于本DN
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-14:08:02
'
' Parameters :       strDN (String)
'                    strJOBID (Variant)
'--------------------------------------------------------------------------------
Public Function checkJobID_ONELOT(strDN As String, strJobID As String) As Boolean
Dim strSql As String
checkJobID_ONELOT = False

If strJobID = "" Then
    Exit Function
End If

strSql = "select * from customershippinguptbl where delivery = '" & strDN & "' and batchnumber  = '" & strJobID & "'"
If Get_OracleCnt(strSql) > 0 Then
    Exit Function
End If

checkJobID_ONELOT = True
End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       getReelQty_ONELOT
' Description:       获取卷盘的入库数量
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-14:13:57
'
' Parameters :       strReelID (String)
'--------------------------------------------------------------------------------
Public Function getReelQty_ONELOT(strReelID As String) As Long
Dim strSql As String

strSql = "select SUM(数量) from erpdata..tblStockNumSub  where 箱号 = '" & strReelID & "'"

getReelQty_ONELOT = Get_SqlserverNo(strSql)

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       getReelMPN_ONELOT
' Description:       获取卷盘的机种
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-28-14:16:40
'
' Parameters :       strJOBID (String)
'--------------------------------------------------------------------------------
Public Function getReelMPN_ONELOT(strDN As String, strJobID As String) As String
Dim strSql As String

strSql = "select distinct marketingpn as mpn from CUSTOMERSHIPPINGUPTBL where delivery = '" & strDN & "' and batchnumber = '" & strJobID & "'"

getReelMPN_ONELOT = Get_OracleStr(strSql)
End Function


Public Function Check37JOB(waferid As String) As Boolean
Dim strSql As String

Check37JOB = True

strSql = " SELECT * FROM MAPPINGDATATEST a,customeroitbl_test b, tbl37testdc c  WHERE  a.substrateid =  '" & waferid & "' and  to_char(b.id) = a.filename  and c.jobid = b.test_mtrl_desc "

If Get_OracleCnt(strSql) > 0 Then
    Check37JOB = False
End If

End Function


Public Function Checkwaferqty(waferid As String) As Boolean
Dim strSql As String

Checkwaferqty = False

strSql = " SELECT * FROM ERPBASE..tblToInRec_Wafer a WHERE a.晶圆ID = REPLACE('" & waferid & "','+','' ) "

If Get_SqlserverCnt(strSql) > 0 Then
    Checkwaferqty = True
End If

End Function


Public Function Checklotqty(lot As String, i As Integer) As Boolean
Dim strSql As String

Checklotqty = False

strSql = "  SELECT a.批号,a.当前存量 FROM ERPBASE..tblstocknum a WHERE a.批号 = '" & lot & "' AND a.当前存量 >= " & i & "  union " & _
         "  SELECT a.工单号,COUNT(DISTINCT a.流程卡编号) FROM erpdata..tblStockNumSub a  WHERE a.工单号 = '" & lot & "'   GROUP BY a.工单号 HAVING  COUNT(DISTINCT a.流程卡编号)  >= " & i & "  "

If Get_SqlserverCnt(strSql) > 0 Then
    Checklotqty = True
End If

End Function




