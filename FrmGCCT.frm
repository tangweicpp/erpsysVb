VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmGCCT 
   Caption         =   "�ͻ�ÿ��CT��ѯ"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab CTTAB 
      Height          =   7815
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   13785
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "ͨ��CT"
      TabPicture(0)   =   "FrmGCCT.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblInfor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmbCustomer"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DTP2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DTP1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "MPS CT"
      TabPicture(1)   =   "FrmGCCT.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fps(1)"
      Tab(1).Control(1)=   "CmdOut"
      Tab(1).Control(2)=   "CmdMPSOut"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Semtech CT��Daily RPT"
      TabPicture(2)   =   "FrmGCCT.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(2)=   "fps(0)"
      Tab(2).Control(3)=   "DTPicker2"
      Tab(2).Control(4)=   "DTPicker1"
      Tab(2).Control(5)=   "Cmd37Out"
      Tab(2).Control(6)=   "Cmd37Query"
      Tab(2).Control(7)=   "CmdInput"
      Tab(2).ControlCount=   8
      Begin VB.CommandButton CmdInput 
         Caption         =   "Daily I/O RPT"
         Height          =   600
         Left            =   -62520
         TabIndex        =   19
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton Cmd37Query 
         Caption         =   "��ѯ"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -67440
         TabIndex        =   17
         Top             =   720
         Width           =   990
      End
      Begin VB.CommandButton Cmd37Out 
         Caption         =   "����"
         Height          =   360
         Left            =   -65640
         TabIndex        =   16
         Top             =   720
         Width           =   990
      End
      Begin VB.CommandButton CmdMPSOut 
         Caption         =   "����"
         Height          =   360
         Left            =   -70800
         TabIndex        =   11
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton CmdOut 
         Caption         =   "��ѯ"
         Height          =   360
         Left            =   -73320
         TabIndex        =   9
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����Excel"
         Height          =   480
         Left            =   2160
         TabIndex        =   1
         Top             =   3360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   264568833
         CurrentDate     =   41424
      End
      Begin MSComCtl2.DTPicker DTP2 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   264568833
         CurrentDate     =   41424
      End
      Begin MSDataListLib.DataCombo CmbCustomer 
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5895
         Index           =   1
         Left            =   -74760
         TabIndex        =   10
         Top             =   1200
         Width           =   14895
         _Version        =   524288
         _ExtentX        =   26273
         _ExtentY        =   10398
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
         SpreadDesigner  =   "FrmGCCT.frx":0054
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -73200
         TabIndex        =   12
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   264568833
         CurrentDate     =   41424
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   -69840
         TabIndex        =   14
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   264568833
         CurrentDate     =   41424
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5895
         Index           =   0
         Left            =   -73560
         TabIndex        =   18
         Top             =   1440
         Width           =   14895
         _Version        =   524288
         _ExtentX        =   26273
         _ExtentY        =   10398
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
         SpreadDesigner  =   "FrmGCCT.frx":04C4
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ڣ�"
         Height          =   195
         Left            =   -71040
         TabIndex        =   15
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ���ڣ� "
         Height          =   195
         Left            =   -74400
         TabIndex        =   13
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ���ڣ� "
         Height          =   195
         Left            =   840
         TabIndex        =   8
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ڣ�"
         Height          =   195
         Left            =   840
         TabIndex        =   7
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ���"
         Height          =   195
         Left            =   1200
         TabIndex        =   6
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label LblInfor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ڴ����У����Եȡ���"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   2250
      End
   End
End
Attribute VB_Name = "FrmGCCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail�֭�
    E_Customer = 1           '�ͻ�
    E_PT                     '�Ϻ�
    E_HTPT                     '�Ϻ�
    E_LotID                  'LotID
    E_RecDate               '��������
    E_LotStart              '��ʼ
    E_OutDate               '���
    E_ShipDate              '��������
    E_ProductCT             ' Product CT
    E_CT                     'CT
    E_Hold                   'Hold
    E_End
    
End Enum


Private Enum E_FPS1          'Detail�֭�
   ' E_ID = 1                 'id��
    E_Customer = 1           '�ͻ�
    E_PT                     '�Ϻ�
    E_HTPT                     '�Ϻ�
    E_LotID                  'LotID
    E_RecDate               '��������
    E_ShipDate              '��������
    E_CT                     'CT
    E_Hold                   'Hold
    E_End
    
End Enum




Dim reportRS As New ADODB.Recordset

Dim listRS As New ADODB.Recordset

Public g_Date           As String



Private Sub CmdAdd_Click()
'����
Dim tempKey As String
Dim tempValue As String
Dim getValue As String
Dim otherValue As String

Dim sqlTemp As String

tempKey = UCase(Trim(txtdelNote.Text))
tempValue = Trim(txtawb.Text)
getValue = CombMo.Text
otherValue = Trim(TxtPackage.Text)

'�ж��Ƿ�������
 If tempKey = "" Or getValue = "" Then
    MsgBox "�������������ύ��", vbInformation, "������ʾ"
    Exit Sub
 
 End If


 
sqlTemp = " insert into  tblsetpf(fieldName,fieldValue,resultValue,other,flag,createby,createdate) values ('" & tempKey & "','" & tempValue & "','" & getValue & "','" & otherValue & "','Y','Auto',sysdate)"
AddSql (sqlTemp)

 MsgBox "��ӳɹ�!", vbInformation, "������ʾ"
 
ShowData_Where

End Sub



Private Sub Cmd37Out_Click()
 If Not ExportFpspreadToExcel(fps(0), "SemtechCT", "SemtechCT") Then Exit Sub
End Sub

Private Sub Cmd37Query_Click()
Dim beginTime As Date
Dim endTime As Date

beginTime = DTPicker1.Value
endTime = DTPicker2.Value

Dim sqlTemp As String





Set reportRS = Get37CTData(Format(beginTime, "YYYY-MM-DD"), Format(endTime, "YYYY-MM-DD"))

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With



End Sub

Private Sub CmdInput_Click()
Dim beginTime37 As String
Dim endTime37 As String
Dim beginTime As String
Dim endTime As String
Dim userid As String

userid = UCase(gUserName)
'37 input

 'ExporToExcel ("  select customerpn as CustomerDeviceName ,product as HTKSDeviceName,waferlot as LOTID,erpcreatedate as InputDate,sum(dieqty) as Quantity from (" & _
   '            " select a.customerpn,a.product,b.waferlot, b.waferid,a.erpcreatedate as erpcreatedate ,b.dieqty from ib_wohistory a ,ib_waferlist b  where customer='37' and substr(a.ordername,3,1)<>'D'" & _
'"and b.ordername=a.ordername and b.waferlot not like '%TEST%' ) X group by customerpn,product,waferlot,erpcreatedate  order by erpcreatedate desc,customerpn,product,waferlot")
               
               
                'ExporToExcel ("SELECT   customerpn AS customerdevicename, product AS htksdevicename," & _
       ' " waferlot AS joblot, x.htlotid, x.movein," & _
        '" SUM(DieQty) As Quantity" & _
  '"  FROM (SELECT a.customerpn, a.product, b.waferlot, b.waferid," & _
    '          "   c.firstname AS htlotid," & _
      ''         "  CCS_37B4765IN(c.containername) AS movein," & _
       '        "   b.dieqty,c.CONTAINERNAME con" & _
       '    " FROM ib_wohistory a, ib_waferlist b, container c, a_lotwafers d" & _
         ' " WHERE customer = '37'" & _
         '   " AND SUBSTR (a.ordername, 3, 1) <> 'D'" & _
          '' "  AND b.ordername = a.ordername" & _
          ' "  AND c.containerid = d.containerid" & _
           ' " AND d.waferscribenumber = b.waferid" & _
          ' "  AND b.waferlot NOT LIKE '%TEST%') x" & _
' " GROUP BY customerpn, product, waferlot, htlotid, movein" & _
'" ORDER BY movein DESC," & _
       ' " customerpn," & _
       ' " product," & _
        '" waferlot," & _
       ' " x.htlotid," & _
       ' " x.movein")
    
beginTime = Format(DTPicker1.Value, "YYYY/MM/DD")
endTime = Format(DTPicker2.Value + 1, "YYYY/MM/DD")
beginTime37 = Format(beginTime, "YYYY-MM-DD")
endTime37 = Format(endTime, "YYYY-MM-DD")

Call del37CTData(userid)

 Call Insert37CTData(beginTime37, userid, endTime37)
       
        
               ExporToExcel ("select distinct * from (select y.�ͻ�����, y.�ͻ�����, y.���ڻ���,y.��������, y.����, x.Ƭ��,x.DIE��, y.���Ƭ��, y.���DIE�� from (select '37' as �ͻ�����, " & _
 " d.mpn_desc as �ͻ�����,substr(a.product, 3, 7) as ���ڻ���,pj_combin_qty.wip_flow(a.product) as ��������,to_char(a.erpcreatedate, 'YYYY-MM-DD') as ����, " & _
" count(b.waferid) as Ƭ��, sum(b.dieqty) as DIE�� from ib_wohistory a,ib_waferlist b, mappingdatatest c, customeroitbl_test d where a.customer = '37' " & _
"  and b.ordername = a.ordername and c.substrateid = b.waferid and to_char(d.id) = c.filename group by '37', d.mpn_desc, substr(a.product, 3, 7), pj_combin_qty.wip_flow(a.product), " & _
"  a.erpcreatedate) x right join (select '37' as �ͻ�����, substr(e.product, 3, 7) as ���ڻ���,to_char(e.packdate, 'YYYY-MM-DD') as ����, " & _
"  pj_combin_qty.wip_flow(e.product) as ��������, count(distinct e.waferid) as ���Ƭ��, sum(e.qty) as ���DIE��, h.customerptno1 as �ͻ�����  from pj_37_ct e , TBLTsvNpiProduct h " & _
" where  h.qtechptno2 =  e.product  group by substr(e.product, 3, 7), to_char(e.packdate, 'YYYY-MM-DD'),h.customerptno1, " & _
"  pj_combin_qty.wip_flow(e.product)) y  on x.���ڻ��� = y.���ڻ��� and x.���� = y.���� and Ƭ�� is null union select x.�ͻ�����,  x.�ͻ�����, x.���ڻ���,x.��������, " & _
"  x.����, x.Ƭ��, x.DIE��, y.���Ƭ��, y.���DIE�� from (select '37' as �ͻ�����, d.mpn_desc as �ͻ�����, substr(a.product, 3, 7) as ���ڻ���, " & _
"  pj_combin_qty.wip_flow(a.product) as ��������, to_char(a.erpcreatedate, 'YYYY-MM-DD') as ����,count(b.waferid) as Ƭ��, sum(b.dieqty) as DIE�� " & _
"  from ib_wohistory a, ib_waferlist b, mappingdatatest c,customeroitbl_test d where a.customer = '37' and b.ordername = a.ordername and c.substrateid = b.waferid " & _
"  and to_char(d.id) = c.filename group by '37', d.mpn_desc, substr(a.product, 3, 7), pj_combin_qty.wip_flow(a.product), a.erpcreatedate) x " & _
"  left join (select '37' as �ͻ�����,substr(e.product, 3, 7) as ���ڻ���, to_char(e.packdate, 'YYYY-MM-DD') as ����, count(distinct e.waferid) as ���Ƭ��, " & _
"   sum(e.qty) as ���DIE��  from pj_37_ct e  group by substr(e.product, 3, 7), to_char(e.packdate, 'YYYY-MM-DD')) y  on x.���ڻ��� = y.���ڻ��� " & _
"  and x.���� = y.���� where x.���� >= '" & beginTime37 & "'  and x.���� <= '" & endTime37 & "') order by  �ͻ�����,���� ")


        
End Sub

Private Sub CmdMPSOut_Click()

  If Not ExportFpspreadToExcel(fps(1), "MPSCT", "MPSCT") Then Exit Sub
  
End Sub

Private Sub CmdOTD_Click()
Dim beginTime As String
Dim endTime As String

beginTime = Format(DTPicker1.Value, "YYYY/MM/DD")
endTime = Format(DTPicker2.Value, "YYYY/MM/DD")


 SqlServerExporToExcel (" SELECT distinct  'Semtech' as Customer , d.CUSTOMERPN as CustomerDeviceName," & _
" b.�Ϻ� as HTKSDeviceName,b.������ as LOTID,CONVERT(varchar(100), d.ERPCREATEDATE, 23) as LotStartDate," & _
" CONVERT(varchar(50), e.CREATE_DATE, 120) as OutputDate,'' as TargetCSD, CONVERT(varchar(50), a.��������, 120) as ActualShippedDate,'' as Meet FROM   [erpdata].[dbo].[tblStockMove]  a  , [erpdata].[dbo].[tblStockMovesub] b, [erpdata].[dbo].[tblTSVwaferlist] c, " & _
" [erpdata].[dbo].[tblTSVworkorder] d, [erpdata].[dbo].[TblQBOXNUMBER_TSV] e " & _
" where a.�ͻ�����='37' and a.ʵ����Ʒ��>0 and a.���߱��=1 and a.��������=1 and  CONVERT(varchar(100), a.��������, 111) >= '" & beginTime & "'  and CONVERT(varchar(100), a.��������, 111) <= '" & endTime & "'" & _
"and b.���ݱ��=a.���ݱ�� and b.������=a.������ and c.WAFERID=b.���̿���� and d.ORDERNAME=c.ORDERNAME and e.WAFERSCRIBENUMBER=b.���̿����  ")



End Sub

Private Sub CmdOut_Click()
'Dim tempBillNo As String
'tempBillNo = UCase(Trim(TxtBillNo.Text))
'
'If tempBillNo = "" Then
'    MsgBox "�����뵥�ݱ��!", vbInformation, "������ʾ"
'    Exit Sub
'End If
'
'
'  Dim judgeEmp As Boolean
'
'judgeEmp = JudgeGRBillNo2(tempBillNo)
' If judgeEmp = False Then
' MsgBox "��ѯ�����˵��ݱ��ά�����������Ϣ����ȷ��!", vbInformation, "������ʾ"
' Exit Sub
'
'End If
'
'
'
' Dim sqlTemp As String
'
' sqlTemp = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & _
'           " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & _
'           " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & _
'           " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & _
'           " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & _
'           " Where a.���ݱ�� = b.���ݱ�� and a.���ݱ��='" + tempBillNo + "' "
'
'  SqlServerExporToExcel (sqlTemp)




'��ϸ����
Dim i As Integer
Dim j As Integer
Dim waferIdTemp As String
Dim lotIdTemp As String

Dim woType As String
Dim baofeiQty As Long
Dim outQty As Long
Dim sendTimeTemp As Long
Dim maxBillNoTemp As String
Dim sendTimes As Long
Dim inTime As String
Dim sendTime2 As String
Dim allCtTemp As Long


Dim allQty As Long


    Set listRS = GetFpsMPSCt()


If listRS.RecordCount <= 0 Then
    MsgBox "��ϸ����û��������ݣ���ȷ��"
    Exit Sub
    
Else


    fps(1).MaxRows = listRS.RecordCount
    For i = 0 To listRS.RecordCount - 1
         baofeiQty = 0
         outQty = 0
         sendTimeTemp = 0
         maxBillNoTemp = ""
         allCtTemp = 0
         
          lotIdTemp = Trim(CStr(listRS.Fields(4).Value))
          outQty = CLng(listRS.Fields(5).Value)
          
          baofeiQty = GetQty37Baofei(lotIdTemp)
          
          allQty = GetOILotQty(lotIdTemp)
          
          
          If (allQty + baofeiQty = allQty) Then
          
          '�ж�lotID��û�з��������������ˣ�����Ct
          
           
         With fps(1)
                 .Row = i + 1
                 .Col = E_FPS1.E_Customer
                 
                  If Trim(listRS.Fields(0).Value) = "68" Or Trim(listRS.Fields(0).Value) = "70" Then
                  .Text = "MPS"
                  Else
                  
                  .Text = CStr(listRS.Fields(0).Value)
                  End If
                 
              
              '��Ʒ�Ϻ�
                  
                 
                .Row = i + 1
                .Col = E_FPS1.E_PT
                .Text = GetNPICustomerPt(Trim(CStr(listRS.Fields(3).Value)))

                 .Row = i + 1
                .Col = E_FPS1.E_HTPT
                .Text = GetNPICustomerHTPt(Trim(CStr(listRS.Fields(3).Value)))

                  .Row = i + 1
                 .Col = E_FPS1.E_LotID
                 .Text = lotIdTemp
                 
                  .Row = i + 1
                 .Col = E_FPS1.E_RecDate
                 
                 inTime = GetLotInFirstTime(lotIdTemp)
                 
                 .Text = inTime
                 
                 .Row = i + 1
                 .Col = E_FPS1.E_ShipDate
                 .Text = Trim(CStr(listRS.Fields(2).Value))
                 
                 
                 '�ж� ���lotid�Ǽ��η���
                 
                 
                 sendTimes = GetLotOutTimes(lotIdTemp)
                 
                 If sendTimes = 1 Then
                 
                   .Row = i + 1
                 .Col = E_FPS1.E_CT
                 .Text = CDate(Trim(CStr(listRS.Fields(2).Value))) - CDate(inTime)
                 
                 Else
                 
                 '�ֱ�������ÿ�γ�����ʱ���
                 
                 
                 
                  Set reportRS = GetLotSendTimes(lotIdTemp)
                  
                   For j = 0 To reportRS.RecordCount - 1
                   
                   
                        sendTime2 = CDate(Trim(CStr(reportRS.Fields(1).Value))) - CDate(inTime)
                        allCtTemp = allCtTemp + sendTime2
                   
                   
                  
                   reportRS.MoveNext

                   Next
                   
                  .Row = i + 1
                 .Col = E_FPS1.E_CT
                 .Text = CStr(allCtTemp / sendTimes)
                 
                 
                 End If
                 
                 
                  .Row = i + 1
                 .Col = E_FPS1.E_Hold
                 .Text = ""
                 
                 
                 
'                   .Row = i + 1
'                 .Col = E_FPS1.E_GDQty
'                 .Text = CStr(listRS.fields(6).Value)
'
'                .Row = i + 1
'                 .Col = E_FPS3.E_NGQty
'                 '��ѯ������û�б���
'                 baofeiQty = GetQty37Baofei(CStr(listRS.fields(8).Value))
'                 If baofeiQty = 0 Then
'                 .Text = ""
'                 Else
'                   .Text = CStr(baofeiQty)
'                 End If
'
'
'                  .Row = i + 1
'                 .Col = E_FPS3.E_Status
'                 '�ж���û�г���
'
'                 outQty = GetQty37OutQty(CStr(listRS.fields(8).Value))
'
'                 If outQty + baofeiQty = CLng(listRS.fields(5).Value) Then
'
'                 '�жϽ����м��η���������ж�Σ������һ�β���ʾX
'                 sendTimeTemp = GetQty37OutTimes(CStr(listRS.fields(8).Value))
'
'                        If sendTimeTemp > 1 Then
'
'                        '�ж�������������Ƿ�Ϊ���ķ������ţ�����ΪX ,����Ϊ0
'                        maxBillNoTemp = GetQty37OutMaxBill(CStr(listRS.fields(8).Value))
'
'                                If maxBillNoTemp = UCase(Trim(Txt37BillNo.Text)) Then
'                                 .Text = "X"
'                                Else
'                                .Text = ""
'                                End If
'
'                        End If
'
'
'
'                 Else
'                  .Text = ""
'
'                 End If
'
'
'                  .Row = i + 1
'                 .Col = E_FPS1.E_LotNumber
'                 .Text = CStr(listRS.fields(8).Value)
'
                 
                
        
        End With
        
        End If
    
NextRecord:
       
        listRS.MoveNext

    Next


End If







End Sub

Private Sub CmdSaver_Click()
'���浽SqlServer��

Dim tempBillNo As String
Dim tempdelNote As String
Dim tempAwb As String

Dim tempWeight As Single
Dim tempPackage As Integer

Dim cmdStrSql As String

tempBillNo = ""
tempdelNote = ""
tempAwb = ""

tempBillNo = UCase(Trim(TxtBillNo.Text))
tempBillNo = Replace(tempBillNo, vbCrLf, "")
tempBillNo = Replace(tempBillNo, vbCr, "")
tempBillNo = Replace(tempBillNo, vbLf, "")


tempdelNote = UCase(Trim(txtdelNote.Text))
tempdelNote = Replace(tempdelNote, vbCrLf, "")
tempdelNote = Replace(tempdelNote, vbCr, "")
tempdelNote = Replace(tempdelNote, vbLf, "")


tempAwb = UCase(Trim(txtawb.Text))
tempAwb = Replace(tempAwb, vbCrLf, "")
tempAwb = Replace(tempAwb, vbCr, "")
tempAwb = Replace(tempAwb, vbLf, "")


If tempBillNo = "" Or tempdelNote = "" Or tempAwb = "" Or Trim(TxtWeight.Text) = "" Or Trim(TxtPackage.Text) = "" Then
    MsgBox "��������������!", vbInformation, "������ʾ"
    Exit Sub
End If



tempWeight = CSng(Trim(TxtWeight.Text))
tempWeight = Replace(tempWeight, vbCrLf, "")
tempWeight = Replace(tempWeight, vbCr, "")
tempWeight = Replace(tempWeight, vbLf, "")


tempPackage = CInt(UCase(Trim(TxtPackage.Text)))
tempPackage = Replace(tempPackage, vbCrLf, "")
tempPackage = Replace(tempPackage, vbCr, "")
tempPackage = Replace(tempPackage, vbLf, "")


'2013-11-21 �жϵ��ݱ�� �Ƿ����

  Dim judgeEmp As Boolean
  judgeEmp = JudgeGRBillNo(tempBillNo)

    If judgeEmp = False Then
    
     MsgBox "�ⵥ�ݱ�Ż�û����GR����ʱ������ά�������Ϣ!", vbInformation, "������ʾ"
     Exit Sub
     
    End If
    
   '�Ƿ���ά����
    judgeEmp = JudgeGRBillNo2(tempBillNo)
     If judgeEmp = True Then
    
     MsgBox "�ⵥ�ݱ����ά�����������ٴ�ά������ȷ��!", vbInformation, "������ʾ"
     Exit Sub
     
    End If
    

    

cmdStrSql = " insert into [erpdata].[dbo].[GRdetailSetUp](���ݱ��,Del_Note,AWB,[Weight],Package) values('" & tempBillNo & "'," & _
             " '" & tempdelNote & "','" & tempAwb & "'," & tempWeight & "," & tempPackage & " )"



AddSql2 (cmdStrSql)

MsgBox "������Ϣ�ɹ� !", vbInformation, "��ʾ"


End Sub

Private Sub CmdSend_Click()
'����

Dim tempBillNo As String
tempBillNo = UCase(Trim(TxtBillNo.Text))

If tempBillNo = "" Then
    MsgBox "�����뵥�ݱ��!", vbInformation, "������ʾ"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNo2(tempBillNo)
 If judgeEmp = False Then
 MsgBox "��ѯ�����˵��ݱ��ά�����������Ϣ����ȷ��!", vbInformation, "������ʾ"
 Exit Sub
 
End If


'    SaveFileSend
    SaveFileSendTest

End Sub

Private Sub Combo2_Change()
TxtBillNoGC.SetFocus

End Sub

Private Sub Combo2_Click()
TxtBillNoGC.SetFocus
End Sub


Private Sub cmdOutput_Click()

 'ExporToExcel ("   select b.mpn_desc as CUSTOMERDEVICENAME,a.productname as HTKSDEVICENAME,b.source_batch_id as LOTID, a.mfgdate as OutputDate,a.qty as Quantity " & _
  '             "  from historymainline a ,customeroitbl_test b  Where a.mfgdate > sysdate - 1 and a.productname like '18X37%' and a.specname = '5272'  " & _
'"and a.callbycdoname = 'WaferWIPMain' and a.cdoname = 'MoveInLot' and a.containername like '%-A%' and b.mtrl_num=substr(a.containername,1,instr(a.containername,'-')-1) ")
           
            ' ExporToExcel ("  SELECT b.mpn_desc AS customerdevicename, a.productname AS htksdevicename," & _
    '"  b.source_batch_id AS lotid, c.firstname AS htlotid, " & _
     ' " ccs_37starge (a.containername) AS stage, " & _
     ' " CCS_37B4765IN (a.containername) AS finishtestout, " & _
     ' " CCS_37_B4770OUT (a.containername) AS inspection, " & _
     ' " CCS_37_5272OUT (a.containername) AS packingoutput, a.qty AS quantity " & _
 '" FROM historymainline a, customeroitbl_test b, container c " & _
'" WHERE   a.productname LIKE '18X37%' " & _
 '"  AND a.specname in ('B4670','B4675','B4770') " & _
 ' " AND a.cdoname = 'MoveInLot' " & _
  '" AND a.containername LIKE '%-A%' " & _
  '" AND a.containerid = c.containerid " & _
  '" and c.STATUS=1 " & _
 '"  AND b.mtrl_num =SUBSTR (a.containername, 1, INSTR (a.containername, '-') - 1) ")
 
          ExporToExcel (" SELECT * FROM (SELECT b.mpn_desc AS customerdevicename," & _
             "  a.productname AS htksdevicename, b.source_batch_id AS lotid," & _
             "  c.firstname AS htlotid, ccs_37starge (a.containername) AS stage," & _
             "  ccs_37b4765in (a.containername) AS finishtestout," & _
              " ccs_37_b4770out (a.containername) AS inspection," & _
             "  ccs_37_5272out (a.containername) AS packingoutput," & _
             "  sum(a.qty) As Quantity" & _
         " FROM historymainline a, customeroitbl_test b, container c,a_lotwafers d" & _
        " WHERE a.productname LIKE '18X37%'" & _
          " AND a.specname IN ('B4735', 'B4765', 'B4770')" & _
          " AND a.cdoname = 'MoveInLot'" & _
          " AND a.containerid = c.containerid" & _
          " AND c.status = 1" & _
          " and d.CONTAINERID=c.CONTAINERID" & _
          " and d.WAFERSCRIBENUMBER=b.mtrl_num" & _
         "  group by b.mpn_desc," & _
            "   a.productname, b.source_batch_id," & _
             "  c.firstname , ccs_37starge (a.containername)," & _
             "  ccs_37b4765in (a.containername) ," & _
             "  ccs_37_b4770out (a.containername) ," & _
             "  ccs_37_5272out (a.containername)" & _
          " ) x WHERE x.stage <> '�ֿ� ����' AND x.stage <> '��װ����ǩ'" & _
 "  AND x.stage <> '�ֿ� TRAY'" & _
 "  AND x.stage <> 'TSV ���'" & _
 "  AND x.stage <> 'OQC'")

End Sub

Private Sub Command1_Click()
LblInfor.Visible = True


Dim i As Integer
Dim sqlTemp As String

''wla wip
'Dim beginTime As String
'Dim endTime As String
'Dim woTemp As String
'Dim productTemp As String
'Dim sqlTemp As String
'Dim cusPTTemp As String
'
'
'
'
'beginTime = Format(DTP1.Value, "YYYY/MM/DD")
'endTime = Format(DTP2.Value, "YYYY/MM/DD")
'cusPTTemp = CusPT.Text
'
'
'  sqlTemp = " select  row_number() over(order by 1) as ""No."" , X.SubName as ""Sub Name"",X.ShipTo as ""Ship To"",X.CustomerDevice as ""Customer Device"",X.GCVersion as ""GC Version"",X.CSTID as ""CST ID"",X.CSTQTY as ""CST QTY"",X.BondPro as ""Bond Pro."",X.FABLotID as ""FAB Lot ID"",X.WaferID as ""Wafer ID"",X.GrossDies as ""Gross Dies"",X.PONO as ""PO NO"",X.WO as ""WO"",X.InvoiceNO as ""Invoice NO"",X.FABDevice as ""FAB Device"",X.PacklotID as ""Pack lot ID"",X.FABOutDate as ""FAB-Out Date"", " & _
' " X.SamplingQty as ""Sampling Qty"",X.PassDies as ""Pass Dies"",X.Yield as ""Yield"",X.Remark as ""Remark""  from ( " & _
' " select distinct 'HTKS' as SubName, 'GC_LG' as ShipTo, replace(a.mpn_desc,'-3','-2.5') as CustomerDevice, a.imager_customer_rev as GCVersion, " & _
'        "   Get_GCWLA_LotID(b.lotid,b.substrateid,to_date('" + beginTime + "','YYYY/MM/DD'),to_date('" + endTime + "','YYYY/MM/DD'),'" + cusPTTemp + "') as CSTID,   Get_GCWLA_Qty(b.lotid,b.substrateid,to_date('" + beginTime + "','YYYY/MM/DD'),to_date('" + endTime + "','YYYY/MM/DD'),'" + cusPTTemp + "') as CSTQTY, 'SH' as BondPro, b.lotid as FABLotID,  substr(b.substrateid,length(b.substrateid)-1,2) as WaferID, b.passbincount as GrossDies, " & _
'        " a.po_num as PONO,a.mtrl_num as WO,  '' InvoiceNO, a.fab_conv_id as FABDevice, c.firstname as PacklotID,to_char(sysdate, 'YYYY-MM-DD') as FABOutDate, " & _
'        " b.passbincount as SamplingQty,  '' as PassDies, '' as Yield, 'A' as Remark " & _
'        " from  tsv_qboxnumber_GC d, mappingdatatest b, customeroitbl_test a, container c " & _
'        " Where d.create_date >=to_date('" + beginTime + "','YYYY/MM/DD') and  d.create_date <to_date('" + endTime + "','YYYY/MM/DD') and b.customershortname = 'GC' and b.substrateid =d.waferscribenumber and b.filename = a.id " & _
'        " and a.customershortname = 'GC' and c.containername = b.substrateid and a.mpn_desc='" + cusPTTemp + "'  " & _
'        " order by   b.lotid,  9 ) X"
'
'
'     ExporToExcel (sqlTemp)

'�Ȳ���������ݣ����浽Oracle��ٵ�����

Dim beginTime As String
Dim endTime As String
Dim beginTime37 As String
Dim endTime37 As String
Dim containerID As String
Dim WaferID As String
Dim dateTemp As String
Dim dateTemp2 As String
Dim userid As String
Dim deptId As String
userid = UCase(gUserName)
deptId = UCase(Trim(CmbCustomer.Text))



If CmbCustomer.Text = "" Then
    MsgBox "����ѡ��ͻ�!", vbInformation, "������ʾ"
    Exit Sub
End If



beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value + 1, "YYYY/MM/DD")

If CmbCustomer.Text = "37" Then

Call del37CTData(userid)

Else

Call delGCCTData(userid)    'ɾ����ǰ����

End If

If CmbCustomer.Text <> "37" Then

Set reportRS = GetGCCTSqlData(beginTime, endTime, deptId)



If (reportRS.RecordCount < 0) Then
    MsgBox "��ѯ�������ݣ� ��ȷ�ϲ�ѯʱ�䣡", vbInformation, "������ʾ"
    Exit Sub

End If



  For i = 1 To reportRS.RecordCount
      
       containerID = CStr(reportRS.Fields(0).Value)
       WaferID = CStr(reportRS.Fields(1).Value)
       dateTemp = Format(CStr(reportRS.Fields(2).Value), "YYYY/MM/DD")
       dateTemp2 = Format(CStr(reportRS.Fields(3).Value), "YYYY/MM/DD")
       
      Call InsertGCCTData(containerID, WaferID, dateTemp, userid, dateTemp2)
       
  
     reportRS.MoveNext
     
  Next i
  
End If
   
If CmbCustomer.Text = "37" Then

beginTime37 = Format(beginTime, "YYYY-MM-DD")
endTime37 = Format(endTime, "YYYY-MM-DD")

 Call Insert37CTData(beginTime37, userid, endTime37)
 
 sqlTemp = " select x.customershortname as �ͻ�����, x.mpn_desc as �ͻ�����,x.product as ���ϱ���,x.source_batch_id as LOT,'#' || to_char(wmsys.wm_concat(x.waferid)) as Wafer_id, count(x.waferid) as wafer_qty,x.ordtemp as ��������," & _
"  x.test_mtrl_desc as JOB, x.po_num as PO,x.ordername as ������,x.poup as PO�ϴ�ʱ��,x.woup as WO�ϴ�ʱ��,x.comdate as Ԥ�����ʱ��,x.packdate as ���ʱ��,pj_37ct.ct_mfg(x.ordername) as ����ʱ��, " & _
"  round((x.packdate - to_date(pj_37ct.ct_mfg(x.ordername), 'YYYY-MM-DD HH24:mi:ss')),1) as MFGCT,round(x.d2dct / 24, 1),round(pj_37ct.ct_holdsum(x.ordername, x.source_batch_id) / 24, 1) as HOLDʱ��, " & _
"  round((x.packdate - to_date(pj_37ct.ct_mfg(x.ordername), 'YYYY-MM-DD HH24:mi:ss')),1) - round(pj_37ct.ct_holdsum(x.ordername, x.source_batch_id) / 24, 1) as CT2, decode(instr((to_char(x.comdate, 'YYYYMMDD') - " & _
"  to_char(x.packdate, 'YYYYMMDD')),'-'),1, 0, 1) as OTD from (select distinct c.customershortname,c.mpn_desc,e.product,c.source_batch_id, decode(instr(b.substrateid, '+'),'0', substr(b.substrateid, -2), " & _
"  substr(b.substrateid, -3)) as waferid,  pj_combin_qty.wip_flow(a.product) as ordtemp,c.test_mtrl_desc,c.po_num,d.ordername, substr(to_char(c.qtech_created_date,'YYYY-MM-DD HH24:mi:ss'), 1, 16) as poup, " & _
"  to_date(e.para7, 'YYYY-MM-DD HH24:mi:ss') as woup, e.planenddate as comdate, a.packdate, round((a.packdate - c.qtech_created_date) * 24, 0) as d2dct from pj_37_ct a, mappingdatatest b, customeroitbl_test c, " & _
"  ib_waferlist d,ib_wohistory  e where b.substrateid = a.waferid and to_char(c.id) = b.filename and d.waferid = a.waferid  and e.ordername = d.ordername) x group by x.customershortname, x.mpn_desc, x.product, " & _
"  x.source_batch_id , x.ordtemp, x.test_mtrl_desc, x.po_num, x.OrderName, x.poup, x.woup, x.comdate, x.packdate, pj_37ct.ct_mfg(x.OrderName), x.d2dct, pj_37ct.ct_holdsum(x.OrderName, x.source_batch_id) "

   
   
 Else
' sqlTemp = "   select  distinct  replace(a.wafernumber,'-A','') wafer ,  a.wafernumber,to_char(Get_GCCT_3010(a.wafernumber), 'YYYY-MM-DD') as date3010,to_char(Get_GCCT_3180(a.wafernumber),'YYYY-MM-DD') as date3180 ,to_char(Get_GCCT_5270(a.wafernumber),'YYYY-MM-DD') as date5270,b.lotid,c.mpn_desc,d.firstname,substr(a.customername,1,10) as outdate,f.product ,b.qtech_created_date ,to_char( to_date(a.fabdate,'YYYY-MM-DD HH24:MI:SS'),'YYYY-MM-DD') as �ֿ��վ�Բ����,f.ordername " & _
'           " from  TSV_GC_CT  a ,mappingdatatest b ,customeroitbl_test c ,container d ,ib_waferlist e ,ib_wohistory f " & _
'           " Where b.SubstrateId = a.productName and c.id=b.filename" & _
'           " and b.customershortname='" & deptId & "' and c.customershortname='" & deptId & "'" & _
'           " and c.source_batch_id=b.lotid and d.containername=a.productname" & _
'           " and a.productname not like 'GXS%' and a.createby='" & userid & "'  and  e.waferid=a.productname  and f.ordername=e.ordername   "
'
sqlTemp = "   select  distinct  replace(a.wafernumber,'-A','') wafer ,  a.wafernumber,to_char(Get_GCCT_3010(a.wafernumber), 'YYYY-MM-DD') as date3010,to_char(Get_GCCT_3180(a.wafernumber),'YYYY-MM-DD') as date3180 ," & _
           " to_char(Get_GCCT_5230(replace(a.wafernumber, '-A', '')),'YYYY-MM-DD') as date5230,to_char(Get_GCCT_5270(a.wafernumber),'YYYY-MM-DD') as date5270,b.lotid,c.mpn_desc,d.firstname,substr(a.customername,1,10) as outdate,f.product ,b.qtech_created_date ," & _
           " to_char( to_date(a.fabdate,'YYYY-MM-DD HH24:MI:SS'),'YYYY-MM-DD') as �ֿ��վ�Բ����,f.ordername " & _
           " from  TSV_GC_CT  a ,mappingdatatest b ,customeroitbl_test c ,container d ,ib_waferlist e ,ib_wohistory f " & _
           " Where b.SubstrateId = a.productName and c.id=b.filename" & _
           " and b.customershortname='" & deptId & "' and c.customershortname='" & deptId & "'" & _
           " and c.source_batch_id=b.lotid and d.containername=a.productname" & _
           " and a.productname not like 'GXS%' and a.createby='" & userid & "'  and  e.waferid=a.productname  and f.ordername=e.ordername   "
  End If

 ExporToExcel (sqlTemp)

LblInfor.Visible = False


End Sub

Private Sub Command2_Click()
'WLT ����  2015-11-11



'����
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtBillNoGCWlt.Text))



If tempBillNo = "" Then
    MsgBox "�����뵥�ݱ��!", vbInformation, "������ʾ"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(tempBillNo)
 If judgeEmp = False Then
 MsgBox "��ѯ�����˵��ݱ�ŵ������Ϣ����ȷ��!", vbInformation, "������ʾ"
 Exit Sub
 
End If



SaveFileSendGCNewWLT




End Sub

Private Sub Command3_Click()
'WLA ERP

'ERP�ĵ���


Dim billNoTemp As String

 billNoTemp = Trim(TxtBillNoGCWLAErp.Text)
 
 
 
  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(billNoTemp)
 If judgeEmp = False Then
 MsgBox "��ѯ�����˵��ݱ�ŵ������Ϣ����ȷ��!", vbInformation, "������ʾ"
 Exit Sub
 
End If

 
 
  
      sqlTemp = "  SELECT row_number() OVER(ORDER BY a.������,a.���̿����) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
" replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.���ݱ��,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST ID]," & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.���ݱ��,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST QTY]," & _
" 'SH' as [Bond Pro.],a.������ as [FAB Lot ID],right(rtrim(ltrim(a.���̿����)),2) as [Wafer ID]," & _
" a.���� as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.���� as [Sampling Qty]," & _
" ''as [Pass Dies],''as [Yield],'A'as [Remark] " & _
" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c WHERE a.���ݱ��='" + billNoTemp + "'" & _
" and b.SOURCE_BATCH_ID=a.������ and c.WAFERSCRIBENUMBER=a.���̿���� and c.WAFERNUMBER=a.������ "
        
        
        
     SqlServerExporToExcel (sqlTemp)


End Sub

Private Sub Command4_Click()

'WLD ERP

'ERP�ĵ���


'Dim billnoTemp As String
'
' billnoTemp = Trim(TxtBillNoGCWLDErp.Text)
'
'
'
'  Dim judgeEmp As Boolean
'
'judgeEmp = JudgeGRBillNoGCWlt(billnoTemp)
' If judgeEmp = False Then
' MsgBox "��ѯ�����˵��ݱ�ŵ������Ϣ����ȷ��!", vbInformation, "������ʾ"
' Exit Sub
'
'End If
'
'
'
'      sqlTemp = "  SELECT row_number() OVER(ORDER BY a.������,a.���̿����) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
'" replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotID](a.���ݱ��,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST ID]," & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty](a.���ݱ��,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST QTY]," & _
'" 'SH' as [Bond Pro.],a.������ as [FAB Lot ID],right(rtrim(ltrim(a.���̿����)),2) as [Wafer ID], d.PRODUCTID as [Wafer Mark]," & _
'" a.���� as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
'" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.���� as [Sampling Qty]," & _
'" ''as [Pass Dies],''as [Yield],''as [Remark] " & _
'" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c , [ERPBASE].[dbo].[tblmappingData] d WHERE a.���ݱ��='" + billnoTemp + "'" & _
'" and b.SOURCE_BATCH_ID=a.������ and c.WAFERSCRIBENUMBER=a.���̿���� and c.WAFERNUMBER=a.������ and d.CUSTOMERSHORTNAME='GC' and d.FILENAME=b.ID and d.SUBSTRATEID=a.���̿���� and d.LOTID=a.������ "
'
'
'
'     SqlServerExporToExcel (sqlTemp)
     
     '--------------------------
     
    
'����
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtBillNoGCWLDErp.Text))



If tempBillNo = "" Then
    MsgBox "�����뵥�ݱ��!", vbInformation, "������ʾ"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(tempBillNo)
 If judgeEmp = False Then
 MsgBox "��ѯ�����˵��ݱ�ŵ������Ϣ����ȷ��!", vbInformation, "������ʾ"
 Exit Sub
 
End If



'Call SaveFileSendGCWLDNew(tempBillNo)

Call SaveFileSendGCWLD(tempBillNo)
 

End Sub

Private Sub Form_Activate()
DTP1.Value = Now - 7

DTP2.Value = Now - 1


DTPicker1.Value = Now - 7

DTPicker2.Value = Now - 1



 g_Date = Format(Now, "YYYY-MM-DD hh:mm:ss")
End Sub




Private Sub IniFpsCT()
    With fps(1)
        .ReDraw = False
        .MaxCols = E_FPS1.E_End - 1
        .MaxRows = 0
        
        '�]�m�榡
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
        .SetText E_FPS1.E_Customer, 0, "Customer"
        .SetText E_FPS1.E_PT, 0, "ShippingDeviceName"
        .SetText E_FPS1.E_HTPT, 0, "HTDeviceName"
        .SetText E_FPS1.E_LotID, 0, "Lot#"
        .SetText E_FPS1.E_RecDate, 0, "ReceiveDate"
        .SetText E_FPS1.E_ShipDate, 0, "ShippingDate"
        .SetText E_FPS1.E_CT, 0, "D2D"
         .SetText E_FPS1.E_Hold, 0, "HoldTime"
    
    
    

        .ColWidth(E_FPS1.E_Customer) = 15
        .ColWidth(E_FPS1.E_PT) = 15
        .ColWidth(E_FPS1.E_HTPT) = 15
        .ColWidth(E_FPS1.E_LotID) = 15
        .ColWidth(E_FPS1.E_RecDate) = 15
        .ColWidth(E_FPS1.E_ShipDate) = 15
         .ColWidth(E_FPS1.E_CT) = 15
        .ColWidth(E_FPS1.E_Hold) = 15
        
        


        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
    
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub IniFpsCT37()
    With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        '�]�m�榡
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
        .SetText E_FPS0.E_Customer, 0, "Customer"
        .SetText E_FPS0.E_PT, 0, "CustomerDeviceName"
        .SetText E_FPS0.E_HTPT, 0, "HTKSDeviceName"
        .SetText E_FPS0.E_LotID, 0, "LotID"
        .SetText E_FPS0.E_RecDate, 0, "WaferReceiveDate"
        .SetText E_FPS0.E_LotStart, 0, "LotStartDate"
        .SetText E_FPS0.E_OutDate, 0, "OutputDate"
        .SetText E_FPS0.E_ShipDate, 0, "ShippingDate"
        .SetText E_FPS0.E_ProductCT, 0, "ProductionCT"
        .SetText E_FPS0.E_CT, 0, "D2D"
        .SetText E_FPS0.E_Hold, 0, "HoldTime"
    
    

        .ColWidth(E_FPS0.E_Customer) = 8
        .ColWidth(E_FPS0.E_PT) = 15
        .ColWidth(E_FPS0.E_HTPT) = 8
        .ColWidth(E_FPS0.E_LotID) = 8
        .ColWidth(E_FPS0.E_RecDate) = 15
        .ColWidth(E_FPS0.E_LotStart) = 8
        .ColWidth(E_FPS0.E_OutDate) = 15
        
        .ColWidth(E_FPS0.E_ShipDate) = 12
        .ColWidth(E_FPS0.E_ProductCT) = 8
        .ColWidth(E_FPS0.E_CT) = 8
        .ColWidth(E_FPS0.E_Hold) = 8
        
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
    
        .ReDraw = True
    End With
    
    
    

End Sub





Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CmbCustomer.RowSource = mainItemRS
CmbCustomer.ListField = mainItemRS("productname").Name
CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub



Private Sub SaveFileSendTest()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim rs          As New ADODB.Recordset

On Error GoTo ErrHandler
    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path_GR & "\" & "QT_FG_CSP_" & Format(g_Date, "YYYYMMDD") & "_" & Format(g_Date, "hhmmss") & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "PO_num,PO_Item,Previous_Batch_ID,Previous_Mtrl_Num,Batch_ID,mtrl_num,mtrl_desc,Mtrl_num_Mtrlgrp,Output_Qty,Consumed_Qty,Reject_Qty,Current_Wafer_Qty" & _
                    ",Film_Frame_Qty,Optical_Quality,Country_of_Assembly,Offshore_ASM_Company,Asm_Containment_type,Date_code,asm_conv_id,asm_excr_id,assembly_facility,Country_of_Test,Offshore_TEST_Company" & _
                    ",Tst_Containment_type,Tst_Program_rev,Created_date,Created_time,Del_Note,AWB,weight(kgs),package" & vbCrLf
    '��ϸ����
    strSql = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & _
           " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & _
           " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & _
           " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & _
           " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & _
           " Where a.���ݱ�� = b.���ݱ�� and a.���ݱ��='" + UCase(Trim(TxtBillNo.Text)) + "' "

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    For i = 1 To rs.RecordCount
        strColData = ""
        For j = 0 To rs.Fields.Count - 1
            If j = 26 Then
             strColData = strColData + Format(g_Date, "hh:mm:ss") + ","
            Else
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
            
            End If
        
           
        Next
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '��������
    'д���ļ�
    LogFile.WriteLine (strDatas)
    
    MsgBox ("���ͳɹ���")
    
    LogFile.Close
    Set LogFile = Nothing
    
ErrHandler:
    Set FSO = Nothing
End Sub


Private Sub SaveFileSendSX()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

On Error GoTo ErrHandler
'��ѯ�����������

fileNo = GetGC_FileNo("SX")

Dim kk As String

    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "SX_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "NO,������,�ͻ�,��Ʒ����,�ͻ�������,�ͻ�Lot,WaferNo,GoodDieQty,BadDieQty,Yield,��������,LaserMark,���,��ע" & vbCrLf
    '��ϸ����
  
  strSql = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & _
          " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & _
          " [Wafer_Mark]as [Laser Mark],CartonNO as [���], [Remark] as [��ע] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
          
          
           
    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        For j = 0 To rs.Fields.Count - 1

             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '��������
    'д���ļ�
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '���ʼ�
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetailSX("SX ��������", strRecipient, g_Path & "\" & "SX_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
    Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'SX') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub

Private Sub SaveFileSendHD()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

On Error GoTo ErrHandler
'��ѯ�����������

fileNo = GetGC_FileNo("HD")

Dim kk As String

    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "NO,������,�ͻ�,�汾,��Ʒ����,�ͻ�������,�ͻ�Lot,WaferNo,GoodDieQty,NGDieQty,ShipmentGoodDie,Yield,��������,��ע" & vbCrLf
    '��ϸ����
  
  strSql = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Fab_Device] as [�汾],[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & _
          " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[Gross_Dies]as [NGDieQty],[NG_Die]as [ShipmentGoodDie] ,[Yield] ,[PACK_Out_Date]as[��������], " & _
          " [Remark] as [��ע] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
          
          
           
    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        For j = 0 To rs.Fields.Count - 1

             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '��������
    'д���ļ�
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '���ʼ�
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetailHD("HD ��������", strRecipient, g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
    Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub



Private Sub SaveFileSendGC()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim waferidMain As String
Dim waferPT As String
Dim waferVer As String
Dim waferVerResult As String

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

On Error GoTo ErrHandler
'��ѯ�����������

fileNo = GetGC_FileNo("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim kk As String

    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PP_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "NO,Sub Name,Ship To,Fab Device,Customer Device,PO NO,WO,GC Version,Invoice NO,PACK-Out Date,PACK Lot ID,FAB Lot ID" & _
               ",Wafer ID,Wafer Mark,Gross Dies,Pass Dies,NG Die,Yield,Remark,System CartonNO,PACK Device,CartonNO,MaskType" & vbCrLf
    '��ϸ����
    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
           " Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        waferVerResult = ""
        
            waferidMain = Trim("" & rs.Fields(0).Value) & "-A"
            
            waferPT = Trim("" & rs.Fields(1).Value)
            
            waferVer = Trim("" & rs.Fields(2).Value)
            
            waferVerResult = GetGCOutRpt_Ver(waferidMain, waferPT, waferVer)
            
            If Len(waferVerResult) <> 3 Then
                MsgBox waferidMain & " ��Ƭ�������볤�Ȳ�����3����ȷ�Ϻú���ܵ�����", vbInformation, "������ʾ"
                'Exit Sub
            
            End If
            
        
        For j = 3 To rs.Fields.Count - 1
             
             If j = 10 Then
             
             strColData = strColData + waferVerResult + ","
             
             Else
             
             
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
             
             End If
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '��������
    'д���ļ�
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '���ʼ�
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetail("GC ��������", strRecipient, g_Path & "\" & "PP_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub




Private Sub SaveFileSendGCNormaNew()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim waferidMain As String
Dim waferPT As String
Dim waferVer As String
Dim waferVerResult As String

Dim dateTemp As String

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

On Error GoTo ErrHandler
'��ѯ�����������

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim kk As String

    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����

    strDatas = "NO,Sub Name,Ship To,Fab Device,Customer Device,GC Version,PO NO,Invoice NO,Ship Out Date,FAB Lot ID," & _
               "Wafer ID,Gross Dies,Sampling Qty,Pass Dies,NG Die,Yield,Pack Lot ID,Wafer Mark,Grade,Carton NO,WO,Remark" & vbCrLf
               
    
    '��ϸ����
'    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
'           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
'           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
'           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
'           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
'           " Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "
           
           
    strSql = " select  rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion,  cast([NO] as int), " & _
             " [Sub_Name],'GCSH' as [Ship_To],[Fab_Device],[Customer_Device],[GC_Version], " & _
             " [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [PACK_Out_Date],[FAB_Lot_ID],[Wafer_ID],[Gross_Dies], " & _
             " '' as [Sampling Qty] ,[Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as  [Pass_Dies],[NG_Die],[Yield] ,[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark] " & _
             " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "'   order by 4 "
 
           
           

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        waferVerResult = ""
        
            waferidMain = Trim("" & rs.Fields(0).Value) & "-A"
            
            waferPT = Trim("" & rs.Fields(1).Value)
            
            waferVer = Trim("" & rs.Fields(2).Value)
            
            waferVerResult = GetGCOutRpt_Ver(waferidMain, waferPT, waferVer)
            
            If Len(waferVerResult) <> 3 Then
                MsgBox waferidMain & " ��Ƭ�������볤�Ȳ�����3����ȷ�Ϻú���ܵ�����", vbInformation, "������ʾ"
                'Exit Sub
            
            End If
            
        
        For j = 3 To rs.Fields.Count - 1
             
'             If j = 10 Then
'
'             strColData = strColData + waferVerResult + ","
'
'             Else
             
             
               If j = 8 Then
             
             strColData = strColData + waferVerResult + ","
             
             ElseIf j = 11 Then
             
             dateTemp = Trim("" & rs.Fields(j).Value)
             
                strColData = strColData + Format(dateTemp, "YYYY-MM-DD") + ","
             
             Else
             
             
             
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
             
             End If
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '��������
    'д���ļ�
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '���ʼ�
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@htkjks.com"
    strRecipientCC = "wanli.ma@htkjks.com"
        
    Call MailDetail("GC ��������", strRecipient, g_Path & "\" & "PL_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub



Private Sub SaveFileSendGCNewWLT()
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim waferidMain As String
Dim waferPT As String
Dim waferVer As String
Dim waferVerResult As String
Dim dateTemp As String

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

On Error GoTo ErrHandler
'��ѯ�����������

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim kk As String

'PL_HTKS_WLT_20151111001.csv
    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_WLT_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "NO,Sub Name,Ship To,Fab Device,Customer Device,GC Version,PO NO,Invoice NO,Ship Out Date,FAB Lot ID," & _
               "Wafer ID,Gross Dies,Sampling Qty,Pass Dies,NG Die,Yield,Pack Lot ID,Wafer Mark,Grade,Carton NO,WO,Remark" & vbCrLf
    '��ϸ����
'    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
'           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
'           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
'           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
'           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
'           " Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "
           
           
           
     strSql = "  select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int), " & _
    " [Sub_Name],'GCSH' as [Ship_To],[Fab_Device],[Customer_Device],[GC_Version], " & _
    " [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [PACK_Out_Date],[FAB_Lot_ID],[Wafer_ID],[Gross_Dies]," & _
    " erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as [Sampling Qty] ,[Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as  [Pass_Dies],[NG_Die],[Yield] ,[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark]" & _
    " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGCWlt.Text)) + "'  order by 4 "
               
               

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        waferVerResult = ""
        
            waferidMain = Trim("" & rs.Fields(0).Value) & "-A"
            
            waferPT = Trim("" & rs.Fields(1).Value)
            
            waferVer = Trim("" & rs.Fields(2).Value)
            
            waferVerResult = GetGCOutRpt_Ver(waferidMain, waferPT, waferVer)
            
            If Len(waferVerResult) <> 3 Then
                MsgBox waferidMain & " ��Ƭ�������볤�Ȳ�����3����ȷ�Ϻú���ܵ�����", vbInformation, "������ʾ"
                'Exit Sub
            
            End If
            
        
        For j = 3 To rs.Fields.Count - 1
             
             If j = 8 Then
             
             strColData = strColData + waferVerResult + ","
             
             ElseIf j = 11 Then
             
             dateTemp = Trim("" & rs.Fields(j).Value)
             
                strColData = strColData + Format(dateTemp, "YYYY-MM-DD") + ","
             
             Else
             
             
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
             
             End If
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '��������
    'д���ļ�
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '���ʼ�
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@htkjks.com"
    strRecipientCC = "wanli.ma@htkjks.com"
        
    Call MailDetail("GC ��������", strRecipient, g_Path & "\" & "PL_HTKS_WLT_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub




Private Sub SaveFileSendGCWLD(billNoTemp As String)
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim waferidMain As String
Dim waferPT As String
Dim waferVer As String
Dim waferVerResult As String

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

On Error GoTo ErrHandler
'��ѯ�����������

fileNo = GetGC_FileNo("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim kk As String

    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PP_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "No.,Sub Name,Ship To,Customer Device,GC Version,CST ID,CST QTY,Bond Pro.,FAB Lot ID,Wafer ID,Wafer Mark,Gross Dies" & _
               ",PO NO,WO,Invoice NO,FAB Device,Pack lot ID,FAB-Out Date,Sampling Qty,Pass Dies,Yield" & vbCrLf
    '��ϸ����
'    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
'           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
'           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
'           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
'           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
'           " Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "
           
           
           
           
     
      strSql = "  SELECT rtrim(ltrim(a.���̿����)) as waferidMain,b.MPN_DESC as device,b.IMAGER_CUSTOMER_REV as gcversion,   row_number() OVER(ORDER BY a.������,a.���̿����) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
" b.MPN_DESC as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.���ݱ��,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST ID]," & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.���ݱ��,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST QTY]," & _
" 'SH' as [Bond Pro.],a.������ as [FAB Lot ID],right(rtrim(ltrim(a.���̿����)),2) as [Wafer ID], d.PRODUCTID as [Wafer Mark]," & _
" a.���� as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.���� as [Sampling Qty]," & _
" ''as [Pass Dies],''as [Yield],'A'as [Remark] " & _
" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c , [ERPBASE].[dbo].[tblmappingData] d WHERE a.���ݱ��='" + billNoTemp + "'" & _
" and b.SOURCE_BATCH_ID=a.������ and c.WAFERSCRIBENUMBER=a.���̿���� and c.WAFERNUMBER=a.������ and d.CUSTOMERSHORTNAME='GC' and d.FILENAME=b.ID and d.SUBSTRATEID=a.���̿���� and d.LOTID=a.������ and a.���=c.QBOXNUMBER "
        
              
           
           
           

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        waferVerResult = ""
        
            waferidMain = Trim("" & rs.Fields(0).Value) & "-A"
            
            waferPT = Trim("" & rs.Fields(1).Value)
            
            waferVer = Trim("" & rs.Fields(2).Value)
            
            waferVerResult = GetGCOutRpt_Ver(waferidMain, waferPT, waferVer)
            
            
            If Len(waferVerResult) <> 3 Then
                MsgBox waferidMain & " ��Ƭ�������볤�Ȳ�����3����ȷ�Ϻú���ܵ�����", vbInformation, "������ʾ"
                Exit Sub
            
            End If
            
            
        
        For j = 3 To rs.Fields.Count - 1
             
             If j = 7 Then
             
             strColData = strColData + waferVerResult + ","
             
             Else
             
             
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
             
             End If
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '��������
    'д���ļ�
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '���ʼ�
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetail("GC ��������", strRecipient, g_Path & "\" & "PP_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub



Private Sub SaveFileSendGCWLDNew(billNoTemp As String)
Dim FSO         As New FileSystemObject
Dim LogFile     As TextStream
Dim strDatas    As String
Dim strRowData  As String
Dim strColData  As String
Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim waferidMain As String
Dim waferPT As String
Dim waferVer As String
Dim waferVerResult As String

Dim maxRow As Integer

Dim rs          As New ADODB.Recordset

Dim fileNo As String

On Error GoTo ErrHandler
'��ѯ�����������

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim kk As String

    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_WLD_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
'    strDatas = "No.,Sub Name,Ship To,Customer Device,GC Version,CST ID,CST QTY,Bond Pro.,FAB Lot ID,Wafer ID,Wafer Mark,Gross Dies" & _
'               ",PO NO,WO,Invoice NO,FAB Device,Pack lot ID,FAB-Out Date,Sampling Qty,Pass Dies,Yield" & vbCrLf
    
    
   strDatas = "NO,Sub Name,Ship To,Fab Device,Customer Device,GC Version,PO NO,Invoice NO,Ship Out Date,FAB Lot ID," & _
               "Wafer ID,Gross Dies,Sampling Qty,Pass Dies,NG Die,Yield,Pack Lot ID,Wafer Mark,Grade,Carton NO,WO,Remark" & vbCrLf
    
    
    '��ϸ����
'    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
'           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
'           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
'           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
'           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
'           " Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "
           
           
           
           
     
      strSql = "  SELECT rtrim(ltrim(a.���̿����)) as waferidMain,b.MPN_DESC as device,b.IMAGER_CUSTOMER_REV as gcversion,   row_number() OVER(ORDER BY a.������,a.���̿����) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
" b.MPN_DESC as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.���ݱ��,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST ID]," & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.���ݱ��,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST QTY]," & _
" 'SH' as [Bond Pro.],a.������ as [FAB Lot ID],right(rtrim(ltrim(a.���̿����)),2) as [Wafer ID], d.PRODUCTID as [Wafer Mark]," & _
" a.���� as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.���� as [Sampling Qty]," & _
" ''as [Pass Dies],''as [Yield],'A'as [Remark] " & _
" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c , [ERPBASE].[dbo].[tblmappingData] d WHERE a.���ݱ��='" + billNoTemp + "'" & _
" and b.SOURCE_BATCH_ID=a.������ and c.WAFERSCRIBENUMBER=a.���̿���� and c.WAFERNUMBER=a.������ and d.CUSTOMERSHORTNAME='GC' and d.FILENAME=b.ID and d.SUBSTRATEID=a.���̿���� and d.LOTID=a.������ and a.���=c.QBOXNUMBER "
        
              
           
           
           

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        waferVerResult = ""
        
            waferidMain = Trim("" & rs.Fields(0).Value) & "-A"
            
            waferPT = Trim("" & rs.Fields(1).Value)
            
            waferVer = Trim("" & rs.Fields(2).Value)
            
            waferVerResult = GetGCOutRpt_Ver(waferidMain, waferPT, waferVer)
            
            
            If Len(waferVerResult) <> 3 Then
                MsgBox waferidMain & " ��Ƭ�������볤�Ȳ�����3����ȷ�Ϻú���ܵ�����", vbInformation, "������ʾ"
                Exit Sub
            
            End If
            
            
        
        For j = 3 To rs.Fields.Count - 1
             
             If j = 7 Then
             
             strColData = strColData + waferVerResult + ","
             
             Else
             
             
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","
             
             End If
           
        Next
        
        If i = maxRow Then
          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
        
        Else
        
        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
        
        End If
        
        rs.MoveNext
    Next
    strDatas = strDatas + strRowData '��������
    'д���ļ�
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '���ʼ�
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@htkjks.com"
    strRecipientCC = "wanli.ma@htkjks.com"
        
    Call MailDetail("GC ��������", strRecipient, g_Path & "\" & "PL_HTKS_WLD_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub





Private Sub SaveFileSend()
'Excel����

Dim strSql      As String
Dim i           As Integer, j           As Integer
Dim rs          As New ADODB.Recordset
Dim RsD         As New ADODB.Recordset
Dim xlApp       As New Excel.Application
Dim xlBook      As Excel.Workbook
Dim xlSheet     As Excel.Worksheet
Dim currentSheetRow As Long

Dim txtHeaderTemp As String



On Error GoTo ErrHandle
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Name = "GrData"
    xlSheet.Activate
    xlApp.Visible = False
'
'
'    '��һ�б���
'    xlSheet.Cells(1, 1) = "PO_num"
'    xlSheet.Cells(1, 2) = "PO_Item"
'    xlSheet.Cells(1, 3) = "Previous_Batch_ID"
'    xlSheet.Cells(1, 4) = "Previous_Mtrl_Num"
'    xlSheet.Cells(1, 5) = "Batch_ID"
'    xlSheet.Cells(1, 6) = "mtrl_num"
'    xlSheet.Cells(1, 7) = "mtrl_desc"
'    xlSheet.Cells(1, 8) = "Mtrl_num_Mtrlgrp"
'    xlSheet.Cells(1, 9) = "Output_Qty"
'    xlSheet.Cells(1, 10) = "Consumed_Qty"
'
'    xlSheet.Cells(1, 11) = "Reject_Qty"
'    xlSheet.Cells(1, 12) = "Current_Wafer_Qty"
'
'    xlSheet.Cells(1, 13) = "Film_Frame_Qty"
'    xlSheet.Cells(1, 14) = "Optical_Quality"
'    xlSheet.Cells(1, 15) = "Country_of_Assembly"
'    xlSheet.Cells(1, 16) = "Offshore_ASM_Company"
'
'    xlSheet.Cells(1, 17) = "Asm_Containment_type"
'
'    xlSheet.Cells(1, 18) = "Date_code"
'    xlSheet.Cells(1, 19) = "asm_conv_id"
'
'    xlSheet.Cells(1, 20) = "asm_excr_id"
'    xlSheet.Cells(1, 21) = "assembly_facility"
'    xlSheet.Cells(1, 22) = "Country_of_Test"
'    xlSheet.Cells(1, 23) = "Offshore_TEST_Company"
'
'    xlSheet.Cells(1, 24) = "Tst_Containment_type"
'    xlSheet.Cells(1, 25) = "Tst_Program_rev"
'    xlSheet.Cells(1, 26) = "Created_date"
'    xlSheet.Cells(1, 27) = "Created_time"
'
'    xlSheet.Cells(1, 28) = "Del_Note"
'    xlSheet.Cells(1, 29) = "AWB"
'    xlSheet.Cells(1, 30) = "weight(kgs)"
'    xlSheet.Cells(1, 31) = "package"
    
    txtHeaderTemp = "PO_num,PO_Item,Previous_Batch_ID,Previous_Mtrl_Num,Batch_ID,mtrl_num,mtrl_desc,Mtrl_num_Mtrlgrp,Output_Qty,Consumed_Qty,Reject_Qty,Current_Wafer_Qty" & _
                    " Film_Frame_Qty,Optical_Quality,Country_of_Assembly,Asm_Containment_type,Date_code,asm_conv_id,asm_excr_id,assembly_facility,Country_of_Test,Offshore_TEST_Company" & _
                    " Tst_Containment_type,Tst_Program_rev,Created_date,Created_time,Del_Note,AWB,weight(kgs),package" & vbCrLf
       xlSheet.Cells(1, 1) = txtHeaderTemp
    
Dim tempBillNo As String
tempBillNo = UCase(Trim(TxtBillNo.Text))

 Dim sqlTemp As String

 strSql = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & _
           " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & _
           " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & _
           " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & _
           " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & _
           " Where a.���ݱ�� = b.���ݱ�� and a.���ݱ��='" + tempBillNo + "' "


    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
    INIConnectSTART
    End If

    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
'     xlSheet.Range("a2:K" & Rs.RecordCount + 1).NumberFormatLocal = "@"
     currentSheetRow = rs.RecordCount + 1
    For i = 2 To rs.RecordCount + 1
        For j = 0 To rs.Fields.Count - 1
            xlSheet.Cells(i, j + 1) = Trim("" & rs.Fields(j).Value)
        Next
        rs.MoveNext
    Next

'
 

  
'    xlSheet.SaveAs g_Path_GR & "\" & Format(g_Date, "YYYY-MM-DD hhmmss") & "WipReport.xls"
    
    xlSheet.SaveAs g_Path_GR & "\" & "QT_FG_CSP_" & Format(g_Date, "YYYYMMDD") & "_" & Format(g_Date, "hhmmss") & ".csv"
    
    
    xlBook.Close
    
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    rs.Close
    Set rs = Nothing
    
    g_IsShouldSend = True
    
    Exit Sub
ErrHandle:
    Set xlApp = Nothing  '"���ٱ��Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Sub



Private Sub Form_Load()

IniCustomerName

IniFpsCT

IniFpsCT37

'txtKey.Text = "PROTECTIVE_FILM_APLD"
'TxtAttri.Text = "BB��"
'
' With fps(0)
'        .ReDraw = False
'        .MaxCols = E_FPS0.E_End - 1
'        .MaxRows = 0
'
'        '�]�m�榡
'        .DAutoHeadings = False
'        .DAutoCellTypes = False
'        .DAutoSizeCols = DAutoSizeColsNone
'
'        .Col = -1
'        .Row = -1
'        .Lock = True
'
'
'        .OperationMode = OperationModeNormal
'        .TypeVAlign = TypeVAlignCenter
'        .SelForeColor = &HFF8080
'
'
'
'        .SetText E_FPS0.E_Key, 0, "�ֶ���"
'        .SetText E_FPS0.E_Value, 0, "�ֶ�ֵ"
'        .SetText E_FPS0.E_getValue, 0, "�Ƿ���Ĥ"
'        .SetText E_FPS0.E_otherValue, 0, "��ע"
'
'
'        .ColWidth(E_FPS0.E_Key) = 20
'        .ColWidth(E_FPS0.E_Value) = 15
'        .ColWidth(E_FPS0.E_getValue) = 15
'        .ColWidth(E_FPS0.E_otherValue) = 25
'
'
'
'        .RowHeight(0) = 20
'        .RowHeight(-1) = 15
'
'
'
'
'        .ReDraw = True
'    End With
'
'
'ShowData_Where





End Sub


Private Sub ShowData_Where()
'Set reportRS = GetpfData()
'
'With fps(0)
'        .MaxRows = 0
'        If reportRS.RecordCount > 0 Then
'            Set .DataSource = reportRS
'
'        End If
'End With

End Sub


Private Sub GCCmdOut_Click()


Dim tempBillNo As String
Dim custNameTemp As String


tempBillNo = UCase(Trim(TxtBillNoGC.Text))
custNameTemp = UCase(Trim(Combo2.Text))

If tempBillNo = "" Or custNameTemp = "" Then
    MsgBox "��ѡ��ͻ����룬���뵥�ݱ��!", vbInformation, "������ʾ"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGC(tempBillNo, custNameTemp)
 If judgeEmp = False Then
 MsgBox "��ѯ�����˵��ݱ�ŵ������Ϣ����ȷ��!", vbInformation, "������ʾ"
 Exit Sub
 
End If
    


 Dim sqlTemp As String
      
 If custNameTemp = "GC" Then
           
'sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [Sub Name],[Ship_To]as [Ship To] ,[Fab_Device]as [Fab Device] ,[Customer_Device] as [Customer Device],[PO_NO] as [PO NO]," & _
'          " [WO],[GC_Version]as [GC Version],[Invoice_NO]as [Invoice NO] ,[PACK_Out_Date]as[PACK-Out Date],[PACK_Lot_ID]as[PACK Lot ID],[FAB_Lot_ID]as[FAB Lot ID] ," & _
'          " [Wafer_ID]as [Wafer ID],[Wafer_Mark]as [Wafer Mark],[Gross_Dies]as [Gross Dies],[Pass_Dies]as [Pass Dies],[NG_Die]as [NG Die] ,[Yield] ," & _
'          " [Remark] , [System_CartonNO]as [System CartonNO], [PACK_Device]as [PACK Device], [CartonNO]as [CartonNO], [MaskType] " & _
'          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
'
'
          
sqlTemp = "  select cast([NO] as int) as NO, " & _
" [Sub_Name],'GCSH' as [Ship To],[Fab_Device],[Customer_Device],[GC_Version], " & _
" [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [Ship Out Date],[FAB_Lot_ID],[Wafer_ID],[Gross_Dies], " & _
" '' as [Sampling Qty] ,[Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as  [Pass_Dies],[NG_Die],[Yield] ,[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark] " & _
" FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "'  order by 1 "
           
          
          
          
          
          
          
    Dim judgeEmp2 As Boolean
    judgeEmp2 = JudgeGRBillNoGCCodeLen(tempBillNo)
     If judgeEmp2 = True Then
     MsgBox "�˱ʷ����� " & tempBillNo & " �к��ж������볤�Ȳ���3����ȷ�ϣ�", vbInformation, "������ʾ"
     Exit Sub
     
    End If
        
                  
ElseIf custNameTemp = "SX" Then

sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & _
          " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & _
          " [Wafer_Mark]as [Laser Mark],CartonNO as [���], [Remark] as [��ע] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
          
          
ElseIf custNameTemp = "HD" Then

sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Fab_Device] as [�汾],[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & _
          " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[Gross_Dies]as [NGDieQty],[NG_Die]as [ShipmentGoodDie] ,[Yield] ,[PACK_Out_Date]as[��������], " & _
          "  [Remark] as [��ע] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
End If

  SqlServerExporToExcel (sqlTemp)

End Sub

Private Sub GCCmdSend_Click()



'��������
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtBillNoGC.Text))
custNameTemp = UCase(Trim(Combo2.Text))


If tempBillNo = "" Or custNameTemp = "" Then
    MsgBox "��ѡ��ͻ����룬���뵥�ݱ��!", vbInformation, "������ʾ"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGC(tempBillNo, custNameTemp)
 If judgeEmp = False Then
 MsgBox "��ѯ�����˵��ݱ�ŵ������Ϣ����ȷ��!", vbInformation, "������ʾ"
 Exit Sub
 
End If

If custNameTemp = "GC" Then

SaveFileSendGCNormaNew

ElseIf custNameTemp = "SX" Then
SaveFileSendSX

ElseIf custNameTemp = "HD" Then
SaveFileSendHD


End If


    
End Sub

Private Sub TxtPackage_KeyPress(KeyAscii As Integer)
Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46) + Chr(13)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
CmdSaver.SetFocus
End If
End Sub

Private Sub TxtWeight_KeyPress(KeyAscii As Integer)

Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46) + Chr(13)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
TxtPackage.SetFocus
End If

End Sub

