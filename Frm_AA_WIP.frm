VERSION 5.00
Begin VB.Form Frm_AA_WIP 
   Caption         =   "AA客户WIP报表实时查询"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5655
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
   ScaleHeight     =   2430
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CommdQuery 
      Caption         =   "报表查询"
      Height          =   720
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   45
   End
End
Attribute VB_Name = "Frm_AA_WIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommdQuery_Click()

    Label1.Caption = "系统正在查询中，请稍等…… "

    Dim strSql As String

    If Cnn.State = 0 Then
        ConOracle

    End If

    '执行Oracle中的Wip存储过程
    Cnn.Execute ("MesReport_WIPData_New")

    strSql = "select distinct PTno as Qtech_PTNO,min(OILotID) OI_Lot ,wafernumber Actual_Lot,waferscribenumber WAFERID,designid,pspt PROBE_SHIP_PART_TYPE,pt MTRL_NUM,pt_desc MTRL_DESC,salesordernumber PO_NUM, to_char(to_date(min(createddate),'yyyy-mm-dd HH24:mi:ss'),'mm/dd/yyyy HH24:mi:ss') CREATED_DATE,to_char(to_date(max(lotpriority),'yyyy-mm-dd'),'mm/dd/yyyy') LOT_PRIORITY,to_char(min(releasedate),'mm/dd/yyyy HH24:mi:ss') RELEASEDATE,max(ageddate) AGED_DATE,max (stagetime) STAGE_TIME, ownername LOT_TYPE, wipstatus STATUS,  to_char( min(customerrequestdate),'mm/dd/yyyy') E_SOD, to_char( min(PMCdate),'mm/dd/yyyy') RE_SOD,workcentername SpecName, min(NDPW) NDPW, min(NG) NG_Number, holdreason HOLD_REASON, holdtime HOLD_DATE " & _
       " from ( select * from (select Get_WIPReport_PTNo(containername) as PTno,OILotID,wafernumber,waferscribenumber,designid,pspt, pt, pt_desc,salesordernumber, min(created_date) as createddate, max(lot_priority) as lotpriority, min(releasedate) as releasedate, max(aged_date) as ageddate," & _
       " max(stage_time) as stagetime,ownername,wipstatus, min(customer_request_date) as customerrequestdate,min(PMCUpdateDate) as PMCdate,workcentername," & _
       " sum(case when substr(containername, -2, 2) = '-A' or instr(containername, '-') <= 0 then NDPW Else 0 end) NDPW ,  sum(case  when substr(containername, instr(containername, '-'), 2) = '-E' then NDPW Else 0 end) NG,  holdreason, holdtime " & _
       " from MesWipData_AutoMail group by Get_WIPReport_PTNo(containername), OILotID, wafernumber,  waferscribenumber, designid, pspt,  pt, pt_desc, salesordernumber, ownername,  wipstatus, workcentername, holdreason,  holdtime) " & _
       " Where NDPW <> 0 ) Group By PTno,wafernumber,waferscribenumber,designid , pspt, pt, pt_desc, salesordernumber, ownername, wipstatus, workcentername, holdreason, holdtime  " & _
       " union all  select '' PTno,source_batch_id OILotID, source_batch_id Actual_Lot, '' WAFERID, design_id designid, probe_ship_part_type pspt, mtrl_num pt, mtrl_desc pt_desc, po_num salesordernumber," & _
       " to_char(to_date(to_char(to_date(created_date || ' ' || created_time, 'YYYY-MM-DD hh24:mi:ss') + 15 / 24, 'YYYY-MM-DD HH24:MI:SS'), 'yyyy-mm-dd HH24:mi:ss'), 'mm/dd/yyyy HH24:mi:ss')," & _
       " to_char(to_date((lot_priority), 'yyyy-mm-dd'), 'mm/dd/yyyy'),  '' releasedate, null ageddate, trunc(sysdate - QTECH_CREATED_DATE, 2) as Stage_Time, '' ownername, 'Pending Qtech OI' wipstatus, to_char( Get_Report_E_Sod_New(to_date( to_char( to_date(created_date||' '||created_time,'YYYY-MM-DD hh24:mi:ss') +15/24 ,'YYYY-MM-DD HH24:MI:SS') ,'YYYY-MM-DD HH24:MI:SS')),'mm/dd/yyyy') customerrequestdate," & _
       " '' PMCdate, 'Bank' workcentername, die_qty NDPW, 0 NG, '' holdreason, null holdtime " & _
       "  From CustomerOItbl_test " & _
       " where customershortname = 'AA' and qtech_created_date > to_date('2013-01-02', 'yyyy-mm-dd') and downqty = 0 " & _
       " order by SpecName, WAFERID, Actual_Lot "
 
    ExporToExcel (strSql)
    
    Label1.Caption = "查询成功！ "

End Sub

