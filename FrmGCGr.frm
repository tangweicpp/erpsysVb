VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmGCGr 
   Caption         =   "GC客户发货信息"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
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
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "正常产品出货"
      TabPicture(0)   =   "FrmGCGr.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Combo2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "GCCmdSend"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "GCCmdOut"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtBillNoGC"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "WLT出货"
      TabPicture(1)   =   "FrmGCGr.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "TxtBillNoGCWlt"
      Tab(1).Control(2)=   "Label1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "WLA MES在线 出货"
      TabPicture(2)   =   "FrmGCGr.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(3)=   "DTP2"
      Tab(2).Control(4)=   "DTP1"
      Tab(2).Control(5)=   "Command1"
      Tab(2).Control(6)=   "CusPT"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "WLA 进ERP系统 出货"
      TabPicture(3)   =   "FrmGCGr.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label7"
      Tab(3).Control(1)=   "Command3"
      Tab(3).Control(2)=   "TxtBillNoGCWLAErp"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "WLD发货"
      TabPicture(4)   =   "FrmGCGr.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command4"
      Tab(4).Control(1)=   "TxtBillNoGCWLDErp"
      Tab(4).Control(2)=   "Label8"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "COG发货"
      TabPicture(5)   =   "FrmGCGr.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin VB.CommandButton Command4 
         Caption         =   "发送报表"
         Height          =   480
         Left            =   -67560
         TabIndex        =   21
         Top             =   1920
         Width           =   990
      End
      Begin VB.TextBox TxtBillNoGCWLDErp 
         Height          =   375
         Left            =   -73560
         TabIndex        =   20
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox TxtBillNoGCWLAErp 
         Height          =   375
         Left            =   -73560
         TabIndex        =   18
         Top             =   1920
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "导出Excel"
         Height          =   480
         Left            =   -67560
         TabIndex        =   17
         Top             =   1920
         Width           =   990
      End
      Begin VB.ComboBox CusPT 
         Height          =   315
         Left            =   -72600
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "导出Excel"
         Height          =   480
         Left            =   -72600
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "发送报表"
         Height          =   480
         Left            =   -67560
         TabIndex        =   8
         Top             =   1920
         Width           =   990
      End
      Begin VB.TextBox TxtBillNoGCWlt 
         Height          =   375
         Left            =   -73560
         TabIndex        =   7
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox TxtBillNoGC 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   1920
         Width           =   3495
      End
      Begin VB.CommandButton GCCmdOut 
         Caption         =   "导出报表"
         Height          =   480
         Left            =   5400
         TabIndex        =   3
         Top             =   1920
         Width           =   990
      End
      Begin VB.CommandButton GCCmdSend 
         Caption         =   "发送报表"
         Height          =   480
         Left            =   7440
         TabIndex        =   2
         Top             =   1920
         Width           =   990
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmGCGr.frx":00A8
         Left            =   1440
         List            =   "FrmGCGr.frx":00AA
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1320
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   375
         Left            =   -72600
         TabIndex        =   12
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   527237121
         CurrentDate     =   41424
      End
      Begin MSComCtl2.DTPicker DTP2 
         Height          =   375
         Left            =   -72600
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   527237121
         CurrentDate     =   41424
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据编号："
         Height          =   195
         Left            =   -74640
         TabIndex        =   22
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据编号："
         Height          =   195
         Left            =   -74640
         TabIndex        =   19
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户机种： "
         Height          =   195
         Left            =   -73800
         TabIndex        =   16
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期： "
         Height          =   195
         Left            =   -73800
         TabIndex        =   15
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期："
         Height          =   195
         Left            =   -73800
         TabIndex        =   14
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据编号："
         Height          =   195
         Left            =   -74640
         TabIndex        =   9
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据编号："
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户："
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   1440
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmGCGr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
'    E_ID = 1                 'id
    E_Key = 1                'Key
    E_Value                  'Value
    E_getValue               'getValue
    E_otherValue             '备注
    E_End
    
End Enum
Dim reportRS As New ADODB.Recordset

Public g_Date           As String



Private Sub CmdAdd_Click()
'新增
Dim tempKey As String
Dim tempValue As String
Dim getValue As String
Dim otherValue As String

Dim sqlTemp As String

tempKey = UCase(Trim(txtdelNote.Text))
tempValue = Trim(txtawb.Text)
getValue = CombMo.Text
otherValue = Trim(TxtPackage.Text)

'判断是否已输入
 If tempKey = "" Or getValue = "" Then
    MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
    Exit Sub
 
 End If


 
sqlTemp = " insert into  tblsetpf(fieldName,fieldValue,resultValue,other,flag,createby,createdate) values ('" & tempKey & "','" & tempValue & "','" & getValue & "','" & otherValue & "','Y','Auto',sysdate)"
AddSql (sqlTemp)

 MsgBox "添加成功!", vbInformation, "友情提示"
 
ShowData_Where

End Sub

Private Sub CmdOut_Click()
Dim tempBillNo As String
tempBillNo = UCase(Trim(TxtBillNo.Text))

If tempBillNo = "" Then
    MsgBox "请输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNo2(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号维护过的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If
    


 Dim sqlTemp As String

 sqlTemp = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & _
           " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & _
           " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & _
           " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & _
           " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & _
           " Where a.单据编号 = b.单据编号 and a.单据编号='" + tempBillNo + "' "

  SqlServerExporToExcel (sqlTemp)


End Sub

Private Sub CmdSaver_Click()
'保存到SqlServer中

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
    MsgBox "请输入完整资料!", vbInformation, "友情提示"
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


'2013-11-21 判断单据编号 是否存在

  Dim judgeEmp As Boolean
  judgeEmp = JudgeGRBillNo(tempBillNo)

    If judgeEmp = False Then
    
     MsgBox "这单据编号还没生成GR，暂时不可以维护相关信息!", vbInformation, "友情提示"
     Exit Sub
     
    End If
    
   '是否已维护过
    judgeEmp = JudgeGRBillNo2(tempBillNo)
     If judgeEmp = True Then
    
     MsgBox "这单据编号已维护过，不可再次维护，请确认!", vbInformation, "友情提示"
     Exit Sub
     
    End If
    

    

cmdStrSql = " insert into [erpdata].[dbo].[GRdetailSetUp](单据编号,Del_Note,AWB,[Weight],Package) values('" & tempBillNo & "'," & _
             " '" & tempdelNote & "','" & tempAwb & "'," & tempWeight & "," & tempPackage & " )"



AddSql2 (cmdStrSql)

MsgBox "保存信息成功 !", vbInformation, "提示"


End Sub

Private Sub CmdSend_Click()
'发送

Dim tempBillNo As String
tempBillNo = UCase(Trim(TxtBillNo.Text))

If tempBillNo = "" Then
    MsgBox "请输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNo2(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号维护过的相关信息，请确认!", vbInformation, "友情提示"
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

Private Sub Command1_Click()
'wla wip
Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim cusPTTemp As String




beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")
cusPTTemp = CusPT.Text

 
  sqlTemp = " select  row_number() over(order by 1) as ""No."" , X.SubName as ""Sub Name"",X.ShipTo as ""Ship To"",X.CustomerDevice as ""Customer Device"",X.GCVersion as ""GC Version"",X.CSTID as ""CST ID"",X.CSTQTY as ""CST QTY"",X.BondPro as ""Bond Pro."",X.FABLotID as ""FAB Lot ID"",X.WaferID as ""Wafer ID"",X.GrossDies as ""Gross Dies"",X.PONO as ""PO NO"",X.WO as ""WO"",X.InvoiceNO as ""Invoice NO"",X.FABDevice as ""FAB Device"",X.PacklotID as ""Pack lot ID"",X.FABOutDate as ""FAB-Out Date"", " & _
 " X.SamplingQty as ""Sampling Qty"",X.PassDies as ""Pass Dies"",X.Yield as ""Yield"",X.Remark as ""Remark""  from ( " & _
 " select distinct 'HTKS' as SubName, 'GC_LG' as ShipTo, replace(a.mpn_desc,'-3','-2.5') as CustomerDevice, a.imager_customer_rev as GCVersion, " & _
        "   Get_GCWLA_LotID(b.lotid,b.substrateid,to_date('" + beginTime + "','YYYY/MM/DD'),to_date('" + endTime + "','YYYY/MM/DD'),'" + cusPTTemp + "') as CSTID,   Get_GCWLA_Qty(b.lotid,b.substrateid,to_date('" + beginTime + "','YYYY/MM/DD'),to_date('" + endTime + "','YYYY/MM/DD'),'" + cusPTTemp + "') as CSTQTY, 'SH' as BondPro, b.lotid as FABLotID,  substr(b.substrateid,length(b.substrateid)-1,2) as WaferID, b.passbincount as GrossDies, " & _
        " a.po_num as PONO,a.mtrl_num as WO,  '' InvoiceNO, a.fab_conv_id as FABDevice, c.firstname as PacklotID,to_char(sysdate, 'YYYY-MM-DD') as FABOutDate, " & _
        " b.passbincount as SamplingQty,  '' as PassDies, '' as Yield, 'A' as Remark " & _
        " from  tsv_qboxnumber_GC d, mappingdatatest b, customeroitbl_test a, container c " & _
        " Where d.create_date >=to_date('" + beginTime + "','YYYY/MM/DD') and  d.create_date <to_date('" + endTime + "','YYYY/MM/DD') and b.customershortname = 'GC' and b.substrateid =d.waferscribenumber and b.filename = a.id " & _
        " and a.customershortname = 'GC' and c.containername = b.substrateid and a.mpn_desc='" + cusPTTemp + "'  " & _
        " order by   b.lotid,  9 ) X"

 
     ExporToExcel (sqlTemp)






End Sub

Private Sub Command2_Click()
'WLT



'发送
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtBillNoGCWlt.Text))



If tempBillNo = "" Then
    MsgBox "请输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If



SaveFileSendGC






End Sub

Private Sub Command3_Click()
'WLA ERP

'ERP的导出


Dim billNoTemp As String

 billNoTemp = Trim(TxtBillNoGCWLAErp.Text)
 
 
 
  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(billNoTemp)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If

 
 
  
      sqlTemp = "  SELECT row_number() OVER(ORDER BY a.工单号,a.流程卡编号) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
" replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID]," & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST QTY]," & _
" 'SH' as [Bond Pro.],a.工单号 as [FAB Lot ID],right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID]," & _
" a.数量 as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.数量 as [Sampling Qty]," & _
" ''as [Pass Dies],''as [Yield],'A'as [Remark] " & _
" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c WHERE a.单据编号='" + billNoTemp + "'" & _
" and b.SOURCE_BATCH_ID=a.工单号 and c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 "
        
        
        
     SqlServerExporToExcel (sqlTemp)


End Sub

Private Sub Command4_Click()

'WLD ERP

'ERP的导出


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
' MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
' Exit Sub
'
'End If
'
'
'
'      sqlTemp = "  SELECT row_number() OVER(ORDER BY a.工单号,a.流程卡编号) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
'" replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotID](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID]," & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST QTY]," & _
'" 'SH' as [Bond Pro.],a.工单号 as [FAB Lot ID],right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID], d.PRODUCTID as [Wafer Mark]," & _
'" a.数量 as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
'" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.数量 as [Sampling Qty]," & _
'" ''as [Pass Dies],''as [Yield],''as [Remark] " & _
'" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c , [ERPBASE].[dbo].[tblmappingData] d WHERE a.单据编号='" + billnoTemp + "'" & _
'" and b.SOURCE_BATCH_ID=a.工单号 and c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 and d.CUSTOMERSHORTNAME='GC' and d.FILENAME=b.ID and d.SUBSTRATEID=a.流程卡编号 and d.LOTID=a.工单号 "
'
'
'
'     SqlServerExporToExcel (sqlTemp)
     
     '--------------------------
     
    
'发送
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtBillNoGCWLDErp.Text))



If tempBillNo = "" Then
    MsgBox "请输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If



Call SaveFileSendGCWLD(tempBillNo)
 
     



End Sub

Private Sub Form_Activate()
DTP1.Value = Now - 1

DTP2.Value = Now

CusPT.AddItem ("GC0310-3")
CusPT.AddItem ("GC0312-3")
CusPT.AddItem ("GC6123-3")

 g_Date = Format(Now, "YYYY-MM-DD hh:mm:ss")
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
    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path_GR & "\" & "QT_FG_CSP_" & Format(g_Date, "YYYYMMDD") & "_" & Format(g_Date, "hhmmss") & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "PO_num,PO_Item,Previous_Batch_ID,Previous_Mtrl_Num,Batch_ID,mtrl_num,mtrl_desc,Mtrl_num_Mtrlgrp,Output_Qty,Consumed_Qty,Reject_Qty,Current_Wafer_Qty" & _
                    ",Film_Frame_Qty,Optical_Quality,Country_of_Assembly,Offshore_ASM_Company,Asm_Containment_type,Date_code,asm_conv_id,asm_excr_id,assembly_facility,Country_of_Test,Offshore_TEST_Company" & _
                    ",Tst_Containment_type,Tst_Program_rev,Created_date,Created_time,Del_Note,AWB,weight(kgs),package" & vbCrLf
    '明细数据
    strSql = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & _
           " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & _
           " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & _
           " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & _
           " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & _
           " Where a.单据编号 = b.单据编号 and a.单据编号='" + UCase(Trim(TxtBillNo.Text)) + "' "

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
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    MsgBox ("发送成功！")
    
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
'查询报表名的序号

fileNo = GetGC_FileNo("SX")

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "SX_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "NO,供货方,客户,产品名称,客户订单号,客户Lot,WaferNo,GoodDieQty,BadDieQty,Yield,出货日期,LaserMark,箱号,备注" & vbCrLf
    '明细数据
  
  strSql = " select cast([NO] as int) as NO,[Sub_Name] as [供货方],[Ship_To]as [客户] ,[Customer_Device] as [产品名称],[PO_NO] as [客户订单号]," & _
          " [FAB_Lot_ID]as[客户Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[出货日期], " & _
          " [Wafer_Mark]as [Laser Mark],CartonNO as [箱号], [Remark] as [备注] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
          
          
           
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
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetailSX("SX 发货报表", strRecipient, g_Path & "\" & "SX_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
    Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'SX') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
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
'查询报表名的序号

fileNo = GetGC_FileNo("HD")

Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "NO,供货方,客户,版本,产品名称,客户订单号,客户Lot,WaferNo,GoodDieQty,NGDieQty,ShipmentGoodDie,Yield,出货日期,备注" & vbCrLf
    '明细数据
  
  strSql = " select cast([NO] as int) as NO,[Sub_Name] as [供货方],[Ship_To]as [客户] ,[Fab_Device] as [版本],[Customer_Device] as [产品名称],[PO_NO] as [客户订单号]," & _
          " [FAB_Lot_ID]as[客户Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[Gross_Dies]as [NGDieQty],[NG_Die]as [ShipmentGoodDie] ,[Yield] ,[PACK_Out_Date]as[出货日期], " & _
          " [Remark] as [备注] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
          
          
           
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
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetailHD("HD 发货报表", strRecipient, g_Path & "\" & "HD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
    Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
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
'查询报表名的序号

fileNo = GetGC_FileNo("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PP_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "NO,Sub Name,Ship To,Fab Device,Customer Device,PO NO,WO,GC Version,Invoice NO,PACK-Out Date,PACK Lot ID,FAB Lot ID" & _
               ",Wafer ID,Wafer Mark,Gross Dies,Pass Dies,NG Die,Yield,Remark,System CartonNO,PACK Device,CartonNO,MaskType" & vbCrLf
    '明细数据
    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
           " Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "

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
                MsgBox waferidMain & " 此片二级代码长度不等于3，请确认好后才能导报表！", vbInformation, "友情提示"
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
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetail("GC 发货报表", strRecipient, g_Path & "\" & "PP_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
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
'查询报表名的序号

fileNo = GetGC_FileNo("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim kk As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PP_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "No.,Sub Name,Ship To,Customer Device,GC Version,CST ID,CST QTY,Bond Pro.,FAB Lot ID,Wafer ID,Wafer Mark,Gross Dies" & _
               ",PO NO,WO,Invoice NO,FAB Device,Pack lot ID,FAB-Out Date,Sampling Qty,Pass Dies,Yield" & vbCrLf
    '明细数据
'    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
'           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
'           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
'           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
'           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
'           " Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "
           
           
           
           
     
      strSql = "  SELECT rtrim(ltrim(a.流程卡编号)) as waferidMain,b.MPN_DESC as device,b.IMAGER_CUSTOMER_REV as gcversion,   row_number() OVER(ORDER BY a.工单号,a.流程卡编号) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
" b.MPN_DESC as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID]," & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST QTY]," & _
" 'SH' as [Bond Pro.],a.工单号 as [FAB Lot ID],right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID], d.PRODUCTID as [Wafer Mark]," & _
" a.数量 as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.数量 as [Sampling Qty]," & _
" ''as [Pass Dies],''as [Yield],'A'as [Remark] " & _
" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c , [ERPBASE].[dbo].[tblmappingData] d WHERE a.单据编号='" + billNoTemp + "'" & _
" and b.SOURCE_BATCH_ID=a.工单号 and c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 and d.CUSTOMERSHORTNAME='GC' and d.FILENAME=b.ID and d.SUBSTRATEID=a.流程卡编号 and d.LOTID=a.工单号 and a.箱号=c.QBOXNUMBER "
        
              
           
           
           

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
                MsgBox waferidMain & " 此片二级代码长度不等于3，请确认好后才能导报表！", vbInformation, "友情提示"
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
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetail("GC 发货报表", strRecipient, g_Path & "\" & "PP_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub



Private Sub SaveFileSend()
'Excel附件

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
'    '第一行标题
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
           " Where a.单据编号 = b.单据编号 and a.单据编号='" + tempBillNo + "' "


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
    Set xlApp = Nothing  '"ユ临北倒Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Sub



Private Sub Form_Load()

'txtKey.Text = "PROTECTIVE_FILM_APLD"
'TxtAttri.Text = "BB栏"
'
' With fps(0)
'        .ReDraw = False
'        .MaxCols = E_FPS0.E_End - 1
'        .MaxRows = 0
'
'        '砞竚Α
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
'        .SetText E_FPS0.E_Key, 0, "字段名"
'        .SetText E_FPS0.E_Value, 0, "字段值"
'        .SetText E_FPS0.E_getValue, 0, "是否贴膜"
'        .SetText E_FPS0.E_otherValue, 0, "备注"
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


Combo2.AddItem ("GC")
Combo2.AddItem ("SX")
Combo2.AddItem ("HD")


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
    MsgBox "请选择客户代码，输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGC(tempBillNo, custNameTemp)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If
    


 Dim sqlTemp As String
      
 If custNameTemp = "GC" Then
           
sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [Sub Name],[Ship_To]as [Ship To] ,[Fab_Device]as [Fab Device] ,[Customer_Device] as [Customer Device],[PO_NO] as [PO NO]," & _
          " [WO],[GC_Version]as [GC Version],[Invoice_NO]as [Invoice NO] ,[PACK_Out_Date]as[PACK-Out Date],[PACK_Lot_ID]as[PACK Lot ID],[FAB_Lot_ID]as[FAB Lot ID] ," & _
          " [Wafer_ID]as [Wafer ID],[Wafer_Mark]as [Wafer Mark],[Gross_Dies]as [Gross Dies],[Pass_Dies]as [Pass Dies],[NG_Die]as [NG Die] ,[Yield] ," & _
          " [Remark] , [System_CartonNO]as [System CartonNO], [PACK_Device]as [PACK Device], [CartonNO]as [CartonNO], [MaskType] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + tempBillNo + "' order by a.CartonNO  "
          
          
    Dim judgeEmp2 As Boolean
    judgeEmp2 = JudgeGRBillNoGCCodeLen(tempBillNo)
     If judgeEmp2 = True Then
     MsgBox "此笔发货单 " & tempBillNo & " 中含有二级代码长度不是3，请确认！", vbInformation, "友情提示"
     Exit Sub
     
    End If
        
                  
ElseIf custNameTemp = "SX" Then

sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [供货方],[Ship_To]as [客户] ,[Customer_Device] as [产品名称],[PO_NO] as [客户订单号]," & _
          " [FAB_Lot_ID]as[客户Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[出货日期], " & _
          " [Wafer_Mark]as [Laser Mark],CartonNO as [箱号], [Remark] as [备注] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + tempBillNo + "' order by a.CartonNO  "
          
          
ElseIf custNameTemp = "HD" Then

sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [供货方],[Ship_To]as [客户] ,[Fab_Device] as [版本],[Customer_Device] as [产品名称],[PO_NO] as [客户订单号]," & _
          " [FAB_Lot_ID]as[客户Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[Gross_Dies]as [NGDieQty],[NG_Die]as [ShipmentGoodDie] ,[Yield] ,[PACK_Out_Date]as[出货日期], " & _
          "  [Remark] as [备注] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + tempBillNo + "' order by a.CartonNO  "
End If

  SqlServerExporToExcel (sqlTemp)

End Sub

Private Sub GCCmdSend_Click()



'发送
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtBillNoGC.Text))
custNameTemp = UCase(Trim(Combo2.Text))


If tempBillNo = "" Or custNameTemp = "" Then
    MsgBox "请选择客户代码，输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGC(tempBillNo, custNameTemp)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If

If custNameTemp = "GC" Then

SaveFileSendGC

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
