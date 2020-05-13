VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Frm_Normal_WIP 
   Caption         =   "通用WIP报表实时查询"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
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
   ScaleHeight     =   5535
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CommdQuery 
      Caption         =   "报表查询"
      Height          =   480
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo CmbCustomer 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户："
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   45
   End
End
Attribute VB_Name = "Frm_Normal_WIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommdQuery_Click()

Dim cusTemp As String

If CmbCustomer.Text = "" Then

    MsgBox "请先选择客户", vbInformation, "友情提示"
    Exit Sub
  Else
  
  cusTemp = Trim(UCase(CmbCustomer.Text))
End If


Label1.Caption = "系统正在查询中，请稍等…… "

Dim strSql As String
'If Cnn.State = 0 Then
'ConOracle
'End If
'
''执行Oracle中的Wip存储过程
' Cnn.Execute ("MesReport_WIPData_New")



'         strSql = "select distinct PTno as Qtech_PTNO,min(OILotID) OI_Lot ,wafernumber Actual_Lot,waferscribenumber WAFERID,designid,pspt PROBE_SHIP_PART_TYPE,pt MTRL_NUM,pt_desc MTRL_DESC,salesordernumber PO_NUM, to_char(to_date(min(createddate),'yyyy-mm-dd HH24:mi:ss'),'mm/dd/yyyy HH24:mi:ss') CREATED_DATE,to_char(to_date(max(lotpriority),'yyyy-mm-dd'),'mm/dd/yyyy') LOT_PRIORITY,to_char(min(releasedate),'mm/dd/yyyy HH24:mi:ss') RELEASEDATE,max(ageddate) AGED_DATE,max (stagetime) STAGE_TIME, ownername LOT_TYPE, wipstatus STATUS,  to_char( min(customerrequestdate),'mm/dd/yyyy') E_SOD, to_char( min(PMCdate),'mm/dd/yyyy') RE_SOD,workcentername SpecName, min(NDPW) NDPW, min(NG) NG_Number, holdreason HOLD_REASON, holdtime HOLD_DATE " & _
'         " from ( select * from (select Get_WIPReport_PTNo(containername) as PTno,OILotID,wafernumber,waferscribenumber,designid,pspt, pt, pt_desc,salesordernumber, min(created_date) as createddate, max(lot_priority) as lotpriority, min(releasedate) as releasedate, max(aged_date) as ageddate," & _
'         " max(stage_time) as stagetime,ownername,wipstatus, min(customer_request_date) as customerrequestdate,min(PMCUpdateDate) as PMCdate,workcentername," & _
'         " sum(case when substr(containername, -2, 2) = '-A' or instr(containername, '-') <= 0 then NDPW Else 0 end) NDPW ,  sum(case  when substr(containername, instr(containername, '-'), 2) = '-E' then NDPW Else 0 end) NG,  holdreason, holdtime " & _
'         " from MesWipData_AutoMail group by Get_WIPReport_PTNo(containername), OILotID, wafernumber,  waferscribenumber, designid, pspt,  pt, pt_desc, salesordernumber, ownername,  wipstatus, workcentername, holdreason,  holdtime) " & _
'         " Where NDPW <> 0 ) Group By PTno,wafernumber,waferscribenumber,designid , pspt, pt, pt_desc, salesordernumber, ownername, wipstatus, workcentername, holdreason, holdtime  " & _
'         " union all  select '' PTno,source_batch_id OILotID, source_batch_id Actual_Lot, '' WAFERID, design_id designid, probe_ship_part_type pspt, mtrl_num pt, mtrl_desc pt_desc, po_num salesordernumber," & _
'         " to_char(to_date(to_char(to_date(created_date || ' ' || created_time, 'YYYY-MM-DD hh24:mi:ss') + 15 / 24, 'YYYY-MM-DD HH24:MI:SS'), 'yyyy-mm-dd HH24:mi:ss'), 'mm/dd/yyyy HH24:mi:ss')," & _
'         " to_char(to_date((lot_priority), 'yyyy-mm-dd'), 'mm/dd/yyyy'),  '' releasedate, null ageddate, trunc(sysdate - QTECH_CREATED_DATE, 2) as Stage_Time, '' ownername, 'Pending Qtech OI' wipstatus, to_char( Get_Report_E_Sod_New(to_date( to_char( to_date(created_date||' '||created_time,'YYYY-MM-DD hh24:mi:ss') +15/24 ,'YYYY-MM-DD HH24:MI:SS') ,'YYYY-MM-DD HH24:MI:SS')),'mm/dd/yyyy') customerrequestdate," & _
'         " '' PMCdate, 'Bank' workcentername, die_qty NDPW, 0 NG, '' holdreason, null holdtime " & _
'         "  From CustomerOItbl_test " & _
'         " where customershortname = 'AA' and qtech_created_date > to_date('2013-01-02', 'yyyy-mm-dd') and downqty = 0 " & _
'         " order by SpecName, WAFERID, Actual_Lot "
         
         
 If CmbCustomer.Text = "GT" Then
 
 strSql = " select SUB_NAME,TO_SUB_NAME,EVENT_DATE,TARGET_DEVICE,FAB_DEVICE,Version,WAFER_LOT,sum(WAFER_QTY) WAFER_QTY " & _
         " ,SUB_STAGE,Replace('#'||Ltrim(MAX(SYS_CONNECT_BY_PATH(WAFER_ID, ' ')),' '),' ','') WAFER_ID,PKG_TYPE,PIN,PRI,OWNERNAME,STATUS " & _
         " ,sum(GOOD_QTY) GOOD_QTY,WORKORDER_ID,PO_NO,V_LOTID,RECEIVE_DATE,SOD,FLOW " & _
         " from (  SELECT 'HTKS' SUB_NAME, '" & cusTemp & "' TO_SUB_NAME, to_char(sysdate,'yyyy-mm-dd') EVENT_DATE,oi.mpn_desc TARGET_DEVICE, " & _
         " oi.fab_conv_id FAB_DEVICE,   oi.imager_customer_rev Version,  lw.WAFERNUMBER WAFER_LOT,  1 WAFER_QTY,   wc.description SUB_STAGE, " & _
         " substr(lw.WAFERSCRIBENUMBER,-2,2) WAFER_ID,  'TSV' as PKG_TYPE,'' PIN,1 PRI,'PROD' OWNERNAME, " & _
         " case when a.currentholdcount>0 then 'HOLD' else 'RUN' end STATUS,  sum(lw.NDPW) GOOD_QTY,  lw.WORKORDERNAME WORKORDER_ID, " & _
         " oi.po_num PO_NO,'' V_LOTID,   oi.created_date RECEIVE_DATE,'' SOD,'TSV' FLOW, " & _
         " ROW_NUMBER() OVER(PARTITION BY lw.WAFERNUMBER,wc.description,case when a.currentholdcount>0 then 'HOLD' else 'RUN' end,lw.WORKORDERNAME,oi.po_num " & _
         " ORDER BY substr(lw.WAFERSCRIBENUMBER,-2,2)) RN " & _
         " FROM CONTAINER A inner join  CURRENTSTATUS B on A.CurrentStatusId = B.CurrentStatusId inner join SPEC S on  B.SpecId = S.SpecId and S.ObjectCategory = 'WIP' " & _
         " inner join A_SCHEDULEDATA SD on A.ScheduleDataId = SD.ScheduleDataId and SD.ObjectType = 'WAFER' inner join PRODUCT P on A.ProductId = P.ProductId and P.objecttype = 'PN' " & _
         " inner join PRODUCTBASE PB on P.PRODUCTBASEID = PB.PRODUCTBASEID inner join SPECBASE SB on S.SpecBaseId = SB.SpecBaseId inner join A_LOTWAFERS lw on a.CONTAINERID = lw.CONTAINERID " & _
         " inner join OPERATION op on s.OPERATIONID = op.OPERATIONID inner join WORKCENTER wc on op.WORKCENTERID = wc.WORKCENTERID inner join A_LotAttributes al on a.CONTAINERID = al.CONTAINERID " & _
         " inner join mfgorder mfg on a.mfgorderid = mfg.mfgorderid inner join CustomerOItbl_test oi on lw.wafernumber=oi.source_batch_id and oi.customershortname='" & cusTemp & "' " & _
         " inner join MAPPINGDATATEST mapp on oi.source_batch_id = mapp.lotid and mapp.customershortname= '" & cusTemp & "' and mapp.substrateid=lw.waferscribenumber " & _
         " WHERE  A.Status = 1 and al.customername = '" & cusTemp & "'  and (A.containername not like '%-E%' or A.containername not like '%-F%') group by oi.mpn_desc,oi.fab_conv_id,oi.imager_customer_rev " & _
         " ,lw.WAFERNUMBER,wc.description,substr(lw.WAFERSCRIBENUMBER,-2,2),case when a.currentholdcount>0 then 'HOLD' else 'RUN' end ,lw.WORKORDERNAME,oi.po_num,oi.created_date order by lw.WAFERNUMBER,substr(lw.WAFERSCRIBENUMBER,-2,2)) " & _
         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN  AND WAFER_LOT = PRIOR WAFER_LOT AND SUB_STAGE = PRIOR SUB_STAGE AND STATUS=PRIOR STATUS AND WORKORDER_ID=PRIOR WORKORDER_ID and PO_NO = PRIOR PO_NO Group by SUB_NAME,TO_SUB_NAME,EVENT_DATE,TARGET_DEVICE,FAB_DEVICE,Version " & _
         " ,WAFER_LOT,SUB_STAGE,PKG_TYPE,PIN,PRI,OWNERNAME,STATUS ,WORKORDER_ID,PO_NO,V_LOTID,RECEIVE_DATE,SOD,FLOW "
         
  

Else

         
         
strSql = " select SUB_NAME,TO_SUB_NAME,EVENT_DATE,TARGET_DEVICE,FAB_DEVICE,Version,WAFER_LOT,sum(WAFER_QTY) WAFER_QTY " & _
         " ,SUB_STAGE,Replace('#'||Ltrim(MAX(SYS_CONNECT_BY_PATH(WAFER_ID, ' ')),' '),' ','') WAFER_ID,PKG_TYPE,PIN,PRI,OWNERNAME,STATUS " & _
         " ,sum(GOOD_QTY) GOOD_QTY,WORKORDER_ID,PO_NO,V_LOTID,RECEIVE_DATE,SOD,FLOW " & _
         " from (  SELECT 'HTKS' SUB_NAME, '" & cusTemp & "' TO_SUB_NAME, to_char(sysdate,'yyyy-mm-dd') EVENT_DATE,oi.mpn_desc TARGET_DEVICE, " & _
         " oi.fab_conv_id FAB_DEVICE,   oi.imager_customer_rev Version,  lw.WAFERNUMBER WAFER_LOT,  1 WAFER_QTY,   wc.description SUB_STAGE, " & _
         " substr(lw.WAFERSCRIBENUMBER,-2,2) WAFER_ID,  'TSV' as PKG_TYPE,'' PIN,1 PRI,'PROD' OWNERNAME, " & _
         " case when a.currentholdcount>0 then 'HOLD' else 'RUN' end STATUS,  sum(lw.NDPW) GOOD_QTY,  lw.WORKORDERNAME WORKORDER_ID, " & _
         " oi.po_num PO_NO,'' V_LOTID,   oi.created_date RECEIVE_DATE,'' SOD,'TSV' FLOW, " & _
         " ROW_NUMBER() OVER(PARTITION BY lw.WAFERNUMBER,wc.description,case when a.currentholdcount>0 then 'HOLD' else 'RUN' end,lw.WORKORDERNAME,oi.po_num " & _
         " ORDER BY substr(lw.WAFERSCRIBENUMBER,-2,2)) RN " & _
         " FROM CONTAINER A inner join  CURRENTSTATUS B on A.CurrentStatusId = B.CurrentStatusId inner join SPEC S on  B.SpecId = S.SpecId and S.ObjectCategory = 'WIP' " & _
         " inner join A_SCHEDULEDATA SD on A.ScheduleDataId = SD.ScheduleDataId and SD.ObjectType = 'WAFER' inner join PRODUCT P on A.ProductId = P.ProductId and P.objecttype = 'PN' " & _
         " inner join PRODUCTBASE PB on P.PRODUCTBASEID = PB.PRODUCTBASEID inner join SPECBASE SB on S.SpecBaseId = SB.SpecBaseId inner join A_LOTWAFERS lw on a.CONTAINERID = lw.CONTAINERID " & _
         " inner join OPERATION op on s.OPERATIONID = op.OPERATIONID inner join WORKCENTER wc on op.WORKCENTERID = wc.WORKCENTERID inner join A_LotAttributes al on a.CONTAINERID = al.CONTAINERID " & _
         " inner join mfgorder mfg on a.mfgorderid = mfg.mfgorderid inner join CustomerOItbl_test oi on lw.wafernumber=oi.source_batch_id and oi.customershortname='" & cusTemp & "' " & _
         " inner join MAPPINGDATATEST mapp on to_char(oi.id) = mapp.filename and mapp.customershortname= '" & cusTemp & "' and mapp.substrateid=lw.waferscribenumber " & _
         " WHERE  A.Status = 1 and al.customername = '" & cusTemp & "'  and (A.containername not like '%-E%' or A.containername not like '%-F%') group by oi.mpn_desc,oi.fab_conv_id,oi.imager_customer_rev " & _
         " ,lw.WAFERNUMBER,wc.description,substr(lw.WAFERSCRIBENUMBER,-2,2),case when a.currentholdcount>0 then 'HOLD' else 'RUN' end ,lw.WORKORDERNAME,oi.po_num,oi.created_date order by lw.WAFERNUMBER,substr(lw.WAFERSCRIBENUMBER,-2,2)) " & _
         " START WITH RN = 1 CONNECT BY RN - 1 = PRIOR RN  AND WAFER_LOT = PRIOR WAFER_LOT AND SUB_STAGE = PRIOR SUB_STAGE AND STATUS=PRIOR STATUS AND WORKORDER_ID=PRIOR WORKORDER_ID and PO_NO = PRIOR PO_NO Group by SUB_NAME,TO_SUB_NAME,EVENT_DATE,TARGET_DEVICE,FAB_DEVICE,Version " & _
         " ,WAFER_LOT,SUB_STAGE,PKG_TYPE,PIN,PRI,OWNERNAME,STATUS ,WORKORDER_ID,PO_NO,V_LOTID,RECEIVE_DATE,SOD,FLOW "
         
  End If
         
         
         
         

 
    ExporToExcel (strSql)
    
Label1.Caption = "查询成功！ "
End Sub

Private Sub Form_Load()


Set mainItemRS = GetWoCustName()
Set CmbCustomer.RowSource = mainItemRS
CmbCustomer.ListField = mainItemRS("name").Name
CmbCustomer.BoundColumn = mainItemRS("id").Name


End Sub
