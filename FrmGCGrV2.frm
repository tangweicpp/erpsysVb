VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmGCGrV2 
   Caption         =   "GC客户发货信息 新版格式"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13380
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
   ScaleWidth      =   13380
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
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "正常产品出货"
      TabPicture(0)   =   "FrmGCGrV2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TxtBillNoGC"
      Tab(0).Control(1)=   "GCCmdOut"
      Tab(0).Control(2)=   "GCCmdSend"
      Tab(0).Control(3)=   "Combo2"
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(5)=   "Label5"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "WLT出货"
      TabPicture(1)   =   "FrmGCGrV2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "TxtBillNoGCWlt"
      Tab(1).Control(2)=   "Label1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "WLA MES在线 出货"
      TabPicture(2)   =   "FrmGCGrV2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CusPT"
      Tab(2).Control(1)=   "Command1"
      Tab(2).Control(2)=   "DTP1"
      Tab(2).Control(3)=   "DTP2"
      Tab(2).Control(4)=   "Label6"
      Tab(2).Control(5)=   "Label3"
      Tab(2).Control(6)=   "Label2"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "WLA 进ERP系统 出货"
      TabPicture(3)   =   "FrmGCGrV2.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "TxtBillNoGCWLAErp"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "WLD发货"
      TabPicture(4)   =   "FrmGCGrV2.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command4"
      Tab(4).Control(1)=   "TxtBillNoGCWLDErp"
      Tab(4).Control(2)=   "Label8"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "COG发货"
      TabPicture(5)   =   "FrmGCGrV2.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "CmdCOGOutRpt"
      Tab(5).Control(1)=   "CmbReportType"
      Tab(5).Control(2)=   "Command6"
      Tab(5).Control(3)=   "TxtCogNo"
      Tab(5).Control(4)=   "Label11"
      Tab(5).Control(5)=   "Label10"
      Tab(5).Control(6)=   "Label9"
      Tab(5).ControlCount=   7
      Begin VB.CommandButton CmdCOGOutRpt 
         Caption         =   "发送报表"
         Height          =   360
         Left            =   -69360
         TabIndex        =   28
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox CmbReportType 
         Height          =   315
         ItemData        =   "FrmGCGrV2.frx":00A8
         Left            =   -73440
         List            =   "FrmGCGrV2.frx":00B5
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "跑基础数据"
         Height          =   360
         Left            =   -69360
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TxtCogNo 
         Height          =   375
         Left            =   -73440
         TabIndex        =   23
         Top             =   840
         Width           =   3495
      End
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
         Left            =   1440
         TabIndex        =   18
         Top             =   1920
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "导出Excel"
         Height          =   480
         Left            =   7440
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
         Left            =   -69480
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
         Left            =   -73560
         TabIndex        =   4
         Top             =   1920
         Width           =   3495
      End
      Begin VB.CommandButton GCCmdOut 
         Caption         =   "导出报表"
         Height          =   480
         Left            =   -69600
         TabIndex        =   3
         Top             =   1920
         Width           =   990
      End
      Begin VB.CommandButton GCCmdSend 
         Caption         =   "发送报表"
         Height          =   480
         Left            =   -67560
         TabIndex        =   2
         Top             =   1920
         Width           =   990
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmGCGrV2.frx":00E9
         Left            =   -73560
         List            =   "FrmGCGrV2.frx":00EB
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
         Format          =   109510657
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
         Format          =   109510657
         CurrentDate     =   41424
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "注：先跑基础数据，等系统提示成功后，再选择报表类型，再发送报表。"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -74400
         TabIndex        =   29
         Top             =   2520
         Width           =   5760
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报表类型："
         Height          =   195
         Left            =   -74520
         TabIndex        =   26
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据编号："
         Height          =   195
         Left            =   -74520
         TabIndex        =   25
         Top             =   960
         Width           =   900
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
         Left            =   360
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
         Left            =   -74640
         TabIndex        =   6
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户："
         Height          =   195
         Left            =   -74280
         TabIndex        =   5
         Top             =   1440
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmGCGrV2"
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
    E_END
    
End Enum
Dim reportRS As New ADODB.Recordset

Public g_Date           As String



Private Sub cmdADD_Click()
'新增
Dim tempKey As String
Dim tempValue As String
Dim getValue As String
Dim otherValue As String

Dim sqlTemp As String

tempKey = UCase(Trim(txtdelNote.text))
tempValue = Trim(txtawb.text)
getValue = CombMo.text
otherValue = Trim(txtPackage.text)

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
tempBillNo = UCase(Trim(TxtBillNo.text))

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

tempBillNo = UCase(Trim(TxtBillNo.text))
tempBillNo = Replace(tempBillNo, vbCrLf, "")
tempBillNo = Replace(tempBillNo, vbCr, "")
tempBillNo = Replace(tempBillNo, vbLf, "")


tempdelNote = UCase(Trim(txtdelNote.text))
tempdelNote = Replace(tempdelNote, vbCrLf, "")
tempdelNote = Replace(tempdelNote, vbCr, "")
tempdelNote = Replace(tempdelNote, vbLf, "")


tempAwb = UCase(Trim(txtawb.text))
tempAwb = Replace(tempAwb, vbCrLf, "")
tempAwb = Replace(tempAwb, vbCr, "")
tempAwb = Replace(tempAwb, vbLf, "")


If tempBillNo = "" Or tempdelNote = "" Or tempAwb = "" Or Trim(txtWeight.text) = "" Or Trim(txtPackage.text) = "" Then
    MsgBox "请输入完整资料!", vbInformation, "友情提示"
    Exit Sub
End If



tempWeight = CSng(Trim(txtWeight.text))
tempWeight = Replace(tempWeight, vbCrLf, "")
tempWeight = Replace(tempWeight, vbCr, "")
tempWeight = Replace(tempWeight, vbLf, "")


tempPackage = CInt(UCase(Trim(txtPackage.text)))
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
tempBillNo = UCase(Trim(TxtBillNo.text))

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

Private Sub CmdCOGOutRpt_Click()

Call Command6_Click

'发送
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtCogNo.text))



If tempBillNo = "" Then
    MsgBox "请输入单据编号，再点跑报表数据，最后再发送报表！", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If


If CmbReportType.text = "PL_HTKS_COG" Then

SaveFileSendCOG_01

ElseIf CmbReportType.text = "PL_HTKS_COG_TRC" Then

SaveFileSendCOG_02

ElseIf CmbReportType.text = "PLP_ERP_COG_HTKJ" Then
 
SaveFileSendCOG_03


End If

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
cusPTTemp = CusPT.text

 
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





  sqlTemp = "  select row_number() over(order by 1) as ""NO"", X.SubName as ""Sub Name"", X.ShipTo as ""Ship To"", X.FABDevice as ""Fab Device"",  X.CustomerDevice as ""Customer Device""," & _
      "  X.GCVersion as ""GC Version"", X.CSTID as ""CST ID"", X.CSTQTY as ""Wafer Qty"", X.BondPro as ""Bond Pro"", X.PONO as ""PO NO"", X.InvoiceNO as ""Invoice NO""," & _
       "    X.FABOutDate as ""Ship Out Date"", X.FABLotID as ""FAB Lot ID"", X.WaferID as ""Wafer ID"", X.GrossDies as ""Gross Dies"", 0 as ""Sampling Qty""," & _
      "  X.PassDies as ""Pass Dies"", 0 as ""NG Die"", X.Yield as ""Yield"", X.PacklotID as ""Pack Lot ID"",  X.mark as ""Wafer Mark"", Grade as ""Grade""," & _
        "  X.CartonBox  as ""Carton NO"",  workordername as ""WO"", '' as ""Remark"" " & _
 "  from (select distinct 'HTKS' as SubName, 'GC_LG' as ShipTo, replace(a.mpn_desc, '-3', '-2.5') as CustomerDevice, a.imager_customer_rev as GCVersion," & _
   "  Get_GCWLA_LotID(b.lotid, b.substrateid, to_date('" + beginTime + "','YYYY/MM/DD'), to_date('" + endTime + "','YYYY/MM/DD'), '" + cusPTTemp + "') as CSTID," & _
  "  Get_GCWLA_Qty(b.lotid, b.substrateid, to_date('" + beginTime + "','YYYY/MM/DD'), to_date('" + endTime + "','YYYY/MM/DD'),  '" + cusPTTemp + "') as CSTQTY," & _
  "   'SH' as BondPro,   b.lotid as FABLotID, substr(b.substrateid, length(b.substrateid) - 1, 2) as WaferID,  b.passbincount as GrossDies,  a.po_num as PONO," & _
  "     a.mtrl_num as WO,  '' InvoiceNO,  a.fab_conv_id as FABDevice, c.firstname as PacklotID, to_char(sysdate, 'YYYY-MM-DD') as FABOutDate, b.passbincount as SamplingQty," & _
    "   '' as PassDies, '' as Yield, 'A' as Remark, b.productid as Mark, lot.workordername, substr(a.imager_customer_rev, 3, 1) As Grade,d.qboxnumbernew as CartonBox " & _
      "     from tsv_qboxnumber_GC  d,  mappingdatatest  b, customeroitbl_test a, container   c,  a_lotwafers lot" & _
       "   Where d.create_date >= to_date('" + beginTime + "','YYYY/MM/DD')  and d.create_date < to_date('" + endTime + "','YYYY/MM/DD')" & _
         "   and b.customershortname = 'GC' and b.substrateid = d.waferscribenumber and b.filename = a.id and a.customershortname = 'GC'" & _
         "   and c.containername = b.substrateid  and a.mpn_desc = '" + cusPTTemp + "'  and lot.containerid=c.containerid order by b.lotid, 9) X "



 
     ExporToExcel (sqlTemp)






End Sub

Private Sub Command2_Click()
'WLT 发货  2015-11-11

Dim strSql As String
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtBillNoGCWlt.text))

If tempBillNo = "" Then
    MsgBox "请输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If

Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(tempBillNo)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
' Exit Sub
 
End If

' strSql = "SELECT row_number() over(order by t.lot_id,t.waferid) AS NO ,t.sub_name,t.SHIP_SITE,t.FAB_CONV_ID,t.cust_device,t.gcversion,t.PO_NUM,t.invoice,t.create_date,t.lot_id,t.waferid,t.gross_die,ISNULL(t.BIN3,L.NDPW) as Sampling_Qty " & _
'",ISNULL(ISNULL(t.BIN1,t.A),K.NDPW) as Pass_Dies1,ISNULL(T.BIN2,m.NDPW) as Pass_Dies2,'' AS Pass_Dies3,ISNULL(ISNULL(T.E,n.NDPW),0) as NG_Die,CONVERT(VARCHAR(10),CONVERT(decimal(18,2), (t.gross_die - ISNULL(ISNULL(T.E,n.NDPW),0))*1.0/(t.gross_die )*100)) + '%' AS Yield   " & _
'",ISNULL(t.sfc,k.FIRSTNAME) as Pack_Lot_ID,t.PRODUCTID,'A' as Grade,rtrim(t.箱号) as 箱号,t.大工单 FROM ( SELECT  'HTKS' AS sub_name,d.SHIP_SITE,RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID " & _
'",a.cust_device,a.gcversion,d.PO_NUM,'' AS invoice, a.create_date,rtrim(a.lot_id) as lot_id,SUBSTRING(REPLACE(b.流程卡编号,'+',''),LEN(a.lot_id)+1,2) as waferid,c.FAILBINCOUNT+c.PASSBINCOUNT AS gross_die " & _
'",e.GRADES,e.QTY,c.PRODUCTID,'A' as Grade,rtrim(ay.箱号) as 箱号,b.大工单,a.qbox,b.流程卡编号,SUBSTRING( e.SFC,12,8) AS SFC " & _
'"FROM erptemp..tblshipreport_new a INNER JOIN erpdata..tblStockNumTree ax  ON ax.箱号 =a.qbox  INNER JOIN erpdata..tblStockNumTree ay  ON ay.序号 = ax.上级序号 " & _
'"INNER JOIN  erpdata..tblStocksqfhsub b  ON b.单据编号 = a.ship_order AND b.箱号 = a.qbox AND b.工单号 = a.lot_id " & _
'"INNER JOIN  ERPBASE..tblmappingData c ON  c.SUBSTRATEID = b.流程卡编号 AND c.LOTID = b.工单号 INNER JOIN  erpbase..tblCustomerOI d ON  CONVERT(VARCHAR(20), d.ID) = c.FILENAME AND d.SOURCE_BATCH_ID = c.LOTID  " & _
'"left JOIN  erptemp..WAFER_BIN_LIST e ON e.WAFER_ID = b.流程卡编号 inner JOIN erpdata..tblErpInStockRelation ee ON ee.SFC_ID = e.SFC  AND CHARINDEX(e.WAFER_ID,ee.WAFER_ID) <> 0    " & _
'"WHERE a.ship_order = '" & UCase(Trim(TxtBillNoGCWlt.Text)) & "' ) AS p PIVOT(sum(qty) FOR grades IN(A,BIN1,BIN2,BIN3, E)) AS T " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV k ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.流程卡编号 AND k.CONTAINERNAME LIKE '%-A' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox AND L.WAFERSCRIBENUMBER = t.流程卡编号 AND L.CONTAINERNAME LIKE '%-A-01' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV m ON m.QBOXNUMBER = t.qbox AND m.WAFERSCRIBENUMBER = t.流程卡编号 AND m.CONTAINERNAME LIKE '%-A-02' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV n ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.流程卡编号 AND n.CONTAINERNAME LIKE '%-E' "
'
'
'SqlServerExporToExcel (strSql)
'
'
''
'
SaveFileSendGCNewWLT

End Sub

Private Sub Command3_Click()
'WLA ERP

'ERP的导出
Dim billNoTemp As String
Dim sqlTemp As String

billNoTemp = Trim(TxtBillNoGCWLAErp.text)
  
If Left$(billNoTemp, 3) = "FDP" Then
'    sqlTemp = "SELECT  row_number() OVER(ORDER BY X.[CST ID],X.[Wafer ID]) AS [NO],X.* FROM ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'"[erpdata].[dbo].[Get_TSV_GCWLA_LotID_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID],[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)), " & _
'" rtrim(ltrim(a.流程卡编号))) as [Wafer Qty], 'SH' as [Bond Pro],B.PO_NUM AS [PO NO],'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.工单号 as [FAB Lot ID], " & _
'" right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID], a.合格数 as [Gross Dies], '0' as [Sampling Qty],a.合格数 as [Pass Dies],0 as [NG Die],''as [Yield],c.FIRSTNAME as [Pack Lot ID], d.productid as [Wafer Mark], " & _
'" 'A' AS Grade,c.QBOXNUMBER as [Carton NO],f.ORDERNAME as WO,'' as [Remark] FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV c  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.调拨编号 ='" & billNoTemp & "' " & _
'" and b.SOURCE_BATCH_ID=a.工单号 and d.filename = cast(b.ID as nvarchar) and c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 and d.SUBSTRATEID=a.流程卡编号 and f.WAFERID=a.流程卡编号 )X"
    
'    sqlTemp = "SELECT  row_number() OVER(ORDER BY X.[CST ID],X.[Wafer ID]) AS [NO],X.* FROM ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID],[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)), " & _
'" rtrim(ltrim(a.流程卡编号))) as [Wafer Qty], 'SH' as [Bond Pro],B.PO_NUM AS [PO NO],'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.工单号 as [FAB Lot ID], " & _
'" right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID], a.合格数 as [Gross Dies], '0' as [Sampling Qty],a.合格数 as [Pass Dies],0 as [NG Die],''as [Yield],c.FIRSTNAME as [Pack Lot ID], d.productid as [Wafer Mark], " & _
'" 'A' AS Grade,c.QBOXNUMBER as [Carton NO],f.ORDERNAME as WO,'' as [Remark] FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV c  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.调拨编号 ='" & billNoTemp & "' " & _
'" and b.SOURCE_BATCH_ID=a.工单号 and d.filename = cast(b.ID as nvarchar) and c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 and d.SUBSTRATEID=a.流程卡编号 and f.WAFERID=a.流程卡编号 )X" & _
'" union SELECT  row_number() OVER(ORDER BY Y.[CST ID],Y.[Wafer ID]) AS [NO],Y.* FROM ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID],[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)), " & _
'" rtrim(ltrim(a.流程卡编号))) as [Wafer Qty], 'SH' as [Bond Pro],B.PO_NUM AS [PO NO],'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.工单号 as [FAB Lot ID], " & _
'" right(rtrim(ltrim(replace(a.流程卡编号,'+',''))),2) as [Wafer ID], a.合格数 as [Gross Dies], '0' as [Sampling Qty],a.合格数 as [Pass Dies],0 as [NG Die],''as [Yield],REPLACE(BB.SFC_ID,'SFCBO:1020,','') as [Pack Lot ID], d.productid as [Wafer Mark], " & _
'" 'A' AS Grade,A.箱号 as [Carton NO],f.ORDERNAME as WO,'' as [Remark] FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata..tblErpInStockDetailInfo aa,erpdata..tblErpInStockRelation bb  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.调拨编号 ='" & billNoTemp & "' " & _
'" and b.SOURCE_BATCH_ID=a.工单号 and d.filename = cast(b.ID as nvarchar)  and d.SUBSTRATEID=a.流程卡编号 and f.WAFERID=a.流程卡编号  and a.箱号 = aa.KEY_VALUE and bb.BOX_ID = aa.BOX_ID and  SUBSTRING(replace(bb.WAFER_ID,'SFCBO:1020,','') " & _
'", CHARINDEX(',', replace(bb.WAFER_ID,'SFCBO:1020,',''))+1,CHARINDEX('::', replace(bb.WAFER_ID,'SFCBO:1020,',''))-CHARINDEX(',', replace(bb.WAFER_ID,'SFCBO:1020,',''))-1) = a.流程卡编号 )Y"
'


        sqlTemp = "SELECT  row_number() OVER(ORDER BY X.[CST ID],X.[Wafer ID]) AS [NO],X.* FROM ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID],[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)), " & _
" rtrim(ltrim(a.流程卡编号))) as [Wafer Qty], 'SH' as [Bond Pro],B.PO_NUM AS [PO NO],'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.工单号 as [FAB Lot ID], " & _
" right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID], a.合格数 as [Gross Dies], '0' as [Sampling Qty],a.合格数 as [Pass Dies],0 as [NG Die],''as [Yield],c.FIRSTNAME as [Pack Lot ID], d.productid as [Wafer Mark], " & _
" 'A' AS Grade,c.QBOXNUMBER as [Carton NO],f.ORDERNAME as WO,'' as [Remark] , a.流程卡编号  FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV c  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.调拨编号 ='" & billNoTemp & "' " & _
" and b.SOURCE_BATCH_ID=a.工单号 and d.filename = cast(b.ID as nvarchar) and c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 and d.SUBSTRATEID=a.流程卡编号 and f.WAFERID=a.流程卡编号 )X  union " & _
"   SELECT row_number() OVER(ORDER BY Y.[FAB Lot ID], Y.[Wafer ID]) AS NO,Y.* FROM (SELECT DISTINCT 'HTKS' as  'Sub Name' , 'GC_LG' as 'Ship To' , " & _
"    B.FAB_CONV_ID AS  'Fab Device' , replace(b.MPN_DESC, '-3', '-2.5') as 'Customer Device' ,  b.IMAGER_CUSTOMER_REV as 'GC Version','' AS [CST ID], " & _
"   '' as [Wafer Qty] , 'SH' as 'Bond Pro', B.PO_NUM AS 'PO NO',''  AS 'Invoice NO', convert(varchar(100), getdate(), 111) AS 'Ship Out Date', " & _
"   A.工单号 as 'FAB Lot ID', right(rtrim(ltrim(replace(a.流程卡编号, '+', ''))), 2) as [Wafer ID], a.合格数 as 'Gross Dies','0' as 'Sampling Qty', " & _
"  a.合格数 AS 'Pass Dies',0 as 'NG Die','' as 'Yield',SUBSTRING( REPLACE(ab.KEYID, 'SFCBO:1020,', ''),1 " & _
"  ,CHARINDEX(rtrim(a.流程卡编号),REPLACE(ab.KEYID, 'SFCBO:1020,', ''))-2) as 'Pack Lot ID',  d.productid as 'Wafer Mark', 'A' AS Grade, A.箱号 as 'Carton NO', " & _
"   f.ORDERNAME as WO, '' as 'Remark' , a.流程卡编号 FROM erpdata..tblStockdbsub a, ERPBASE..tblCustomerOI  b, erpdata..tblErpInStockDetailInfo aa,erpdata..tblErpInStockDetailInfo ab, " & _
"   ERPBASE..tblmappingData  d, erpdata..tblTSVwaferlist  f  WHERE a.调拨编号 = '" & billNoTemp & "' AND b.SOURCE_BATCH_ID = a.工单号 and d.filename = cast(b.ID as nvarchar) " & _
"  and d.SUBSTRATEID = a.流程卡编号 AND f.WAFERID = a.流程卡编号 and SUBSTRING(a.箱号,1,CASE WHEN CHARINDEX('_VT',A.箱号 )>0 THEN  CHARINDEX('_VT',A.箱号 )-1 ELSE len(A.箱号) END )= aa.KEY_VALUE  AND ab.BOX_ID = aa.BOX_ID AND ab.KEY_TYPE = 'WAFER' AND ab.KEY_VALUE = a.流程卡编号 ) Y "
    
    
    
   Call SqlServerExporToExcel_WLA_new(sqlTemp, billNoTemp)

    Exit Sub
End If
  
  
Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(billNoTemp)
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If

      'sqlTemp = "  SELECT row_number() OVER(ORDER BY a.工单号,a.流程卡编号) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
'" replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID]," & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST QTY]," & _
'" 'SH' as [Bond Pro.],a.工单号 as [FAB Lot ID],right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID]," & _
'" a.数量 as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
'" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.数量 as [Sampling Qty]," & _
'" ''as [Pass Dies],''as [Yield],'A'as [Remark] " & _
'" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c WHERE a.单据编号='" + billNoTemp + "'" & _
'" and b.SOURCE_BATCH_ID=a.工单号 and c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 "


  sqlTemp = " SELECT  row_number() OVER(ORDER BY X.[CST ID],X.[Wafer ID]) AS [NO],X.* FROM " & _
 " ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
" B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
"[erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID], " & _
"[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.单据编号,rtrim(ltrim(a.工单号)), " & _
"   rtrim(ltrim(a.流程卡编号))) as [Wafer Qty], 'SH' as [Bond Pro],B.PO_NUM AS [PO NO], " & _
"   '' [Invoice NO],convert(varchar, getdate(), 111) AS [Ship Out Date], A.工单号 as [FAB Lot ID], " & _
 "  right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID], a.数量 as [Gross Dies], '0' as [Sampling Qty], " & _
 "  a.数量 as [Pass Dies],0 as [NG Die],''as [Yield],c.FIRSTNAME as [Pack Lot ID], d.productid as [Wafer Mark], " & _
 "  'A' AS Grade,c.QBOXNUMBER as [Carton NO],f.ORDERNAME as WO,'' as [Remark] " & _
  "  FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b , " & _
  "   erpdata.dbo.TblQBOXNUMBER_TSV c  ,[ERPBASE].[dbo].[tblmappingData] d, " & _
   "  [erpdata].[dbo].[tblTSVwaferlist] f WHERE a.单据编号='" + billNoTemp + "' " & _
   "  and b.SOURCE_BATCH_ID=a.工单号 and d.filename = cast(b.ID as nvarchar) " & _
   "  and c.WAFERSCRIBENUMBER=a.流程卡编号 " & _
   "  and c.WAFERNUMBER=a.工单号 " & _
   "  and d.SUBSTRATEID=a.流程卡编号 " & _
   "  and f.WAFERID=a.流程卡编号 )X "
        
        
        
     SqlServerExporToExcel_WLA (sqlTemp)


End Sub


Public Sub SqlServerExporToExcel_WLA_new(strOpen As String, order As String)

Dim Rs_Data As New ADODB.Recordset
Dim Irowcount As Long
Dim Icolcount As Integer

Dim i As Integer
Dim lot As String
Dim WAFER As String
Dim sqllot As String
Dim Rs_lot As New ADODB.Recordset


    
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
            Exit Sub
        End If
        Irowcount = .RecordCount
        Icolcount = .Fields.count
    End With
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    
    
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
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.name = "宋体"
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Size = 11
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = False
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
        
        For i = 2 To Irowcount + 1
        
        lot = xlSheet.Cells(i, 13)
        WAFER = xlSheet.Cells(i, 26)
        
        sqllot = "SELECT  erpdata.dbo.Get_TSV_GCWLA_LotID_NewDB('" & order & "',rtrim(ltrim('" & lot & "')) ,rtrim(ltrim('" & WAFER & "'))) as [CST ID] " & _
                  " , erpdata.dbo.Get_TSV_GCWLA_LotIDQty_NewDB('" & order & "',rtrim(ltrim('" & lot & "'))  ,  rtrim(ltrim('" & WAFER & "'))) as [Wafer Qty]"
        
        If Rs_lot.State = adStateOpen Then Rs_lot.Close
        Rs_lot.Open sqllot, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        
        xlSheet.Cells(i, 7) = Rs_lot.Fields(0).Value
        xlSheet.Cells(i, 8) = Rs_lot.Fields(1).Value
        xlSheet.Cells(i, 26) = ""

        Next i
        
       xlSheet.Cells(1, 26) = ""

    End With
    xlApp.Visible = True
    
    xlApp.Application.Visible = True
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing

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

tempBillNo = UCase(Trim(TxtBillNoGCWLDErp.text))



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



'Call SaveFileSendGCWLDNew(tempBillNo)

Call SaveFileSendGCWLD(tempBillNo)
 

End Sub

Private Sub Command5_Click()
'CoG导出

Dim tempBillNo As String
Dim custNameTemp As String


tempBillNo = UCase(Trim(TxtCogNo.text))


If tempBillNo = "" Then
    MsgBox "请输入单据编号!", vbInformation, "友情提示"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGC(tempBillNo, "GC")
 If judgeEmp = False Then
 MsgBox "查询不到此单据编号的相关信息，请确认!", vbInformation, "友情提示"
 Exit Sub
 
End If
    


 Dim sqlTemp As String
      
 If custNameTemp = "GC" Then
           


          
sqlTemp = "  select cast([NO] as int) as NO, " & _
" [Sub_Name],'GCSH' as [Ship To],[Fab_Device],[Customer_Device],[GC_Version], " & _
" [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [Ship Out Date],[FAB_Lot_ID],[Wafer_ID],[Gross_Dies], " & _
" '' as [Sampling Qty] ,[Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as  [Pass_Dies],[NG_Die],[Yield] ,[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark] " & _
" FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + tempBillNo + "'  order by 1 "
           
          
          
          
          
   
End If

  SqlServerExporToExcel (sqlTemp)


End Sub

Private Sub Command6_Click()
'COG 基础数据

'正常发送
Dim tempBillNo As String
Dim custNameTemp As String
Dim i As Integer
Dim qboxNoTemp As String
Dim containerTemp As String
Dim lvQboxTemp As String

Dim cmdStr As String

tempBillNo = UCase(Trim(TxtCogNo.text))
custNameTemp = "GC"


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


'根据ERP发货单号，中的小箱号，跑出数据

Call delCOGIniData      '删除以前数据 oracle cog 小箱号

Call delCOGRptInt01      '删除以前数据 sqlserver  跑报表1中的铝箔袋信息

Call delCOGRptInt02      '删除以前数据 oracle 跑报表2中的Tray信息

Set reportRS = GetCOGBaseData2(tempBillNo)

If (reportRS.RecordCount < 0) Then
    MsgBox "查询不到数据， 请确认单据编号！", vbInformation, "友情提示"
    Exit Sub

End If


  For i = 1 To reportRS.RecordCount
      
 
       qboxNoTemp = CStr(reportRS.Fields(0).Value)
       cmdStr = "  insert into   GR_COG_IniData (qboxnumber,containername) " & _
                   "  select distinct  a.qboxnumber,a.containername from  tsv_qboxnumber_details a " & _
                   "  where a.qboxnumber='" & qboxNoTemp & "' and a.customername='GC' and a.specname='5275' "

         AddSql (cmdStr)
          
     reportRS.MoveNext
     
  Next i
  
   reportRS.Close
Set reportRS = Nothing


  
  '把铝箔袋信息跑出，放到ERP中
  
Set reportRS = GetCOGLVDataList()

If (reportRS.RecordCount < 0) Then
    MsgBox "查询不到数据， 联系IT确认铝箔袋标签信息有无异常！", vbInformation, "友情提示"
    Exit Sub

End If

  For i = 1 To reportRS.RecordCount
      
 
       qboxNoTemp = CStr(reportRS.Fields(0).Value)
       containerTemp = CStr(reportRS.Fields(1).Value)
       lvQboxTemp = CStr(reportRS.Fields(2).Value)
       
       
       cmdStr = "  insert into   [erpdata].[dbo].[GR_COG_LV_Data]  (qboxnumber,containername,lvbarcodeqbox) values " & _
                   "  ( '" & qboxNoTemp & "','" & containerTemp & "','" & lvQboxTemp & "')"
            
         AddSql2 (cmdStr)
          
     reportRS.MoveNext
     
  Next i
  
    'end 把铝箔袋信息跑出，放到ERP中


Call AddCOGRptInt02      '插入 oracle 跑报表2中的Tray信息
Call AddCOGRptInt02_2      '插入 oracle 跑报表2中的Tray信息 02部分
Call AddCOGRptInt02_3      '插入 oracle 跑报表2中的Tray信息 03部分

'PL_HTKS_COG cog报表基础资料执行SQL SEVER存储过程
  Set adoCmd = New ADODB.Command
         Set adoCmd.ActiveConnection = INIadoCon2
             adoCmd.CommandText = "PL_HTKS_COG"
             adoCmd.Parameters.Refresh
             adoCmd.CommandType = adCmdStoredProc
        
          Set adoprm1 = New ADODB.Parameter   '参数为发货单
          adoprm1.type = adChar
          adoprm1.Size = 20
          adoprm1.Direction = adParamInput
          adoprm1.Value = tempBillNo
          adoCmd.Parameters.Append adoprm1
          adoCmd.Execute
          
'PLP_ERP_COG_HTKJ cog报表基础资料执行SQL SEVER存储过程
  Set adoCmd1 = New ADODB.Command
         Set adoCmd1.ActiveConnection = INIadoCon2
             adoCmd1.CommandText = "PLP_ERP_COG_HTKJ"
             adoCmd1.Parameters.Refresh
             adoCmd1.CommandType = adCmdStoredProc
        
          Set adoPrm11 = New ADODB.Parameter   '参数为发货单
          adoPrm11.type = adChar
          adoPrm11.Size = 20
          adoPrm11.Direction = adParamInput
          adoPrm11.Value = tempBillNo
          adoCmd1.Parameters.Append adoprm1
          adoCmd1.Execute



Cnn.Execute ("CCS_COG_SENDREPORT") '初始化TRC报表资料
MsgBox "跑报表数据已完成!", vbInformation, "友情提示"
End Sub

Private Sub Form_Activate()
DTP1.Value = Now - 1

DTP2.Value = Now

CusPT.AddItem ("GC0310-3")
CusPT.AddItem ("GC0312-3")
CusPT.AddItem ("GC6123-3")
CusPT.AddItem ("GC6133-3")
CusPT.AddItem ("GC030A-3")
CusPT.AddItem ("GC032A-3")
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
           " Where a.单据编号 = b.单据编号 and a.单据编号='" + UCase(Trim(TxtBillNo.text)) + "' "

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    For i = 1 To rs.RecordCount
        strColData = ""
        For j = 0 To rs.Fields.count - 1
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

Dim KK As String

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
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGC.text)) + "' order by 1  "
          
          
           
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
        For j = 0 To rs.Fields.count - 1

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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'SX') "
    
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

Dim KK As String

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
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGC.text)) + "' order by 1  "
          
          
           
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
        For j = 0 To rs.Fields.count - 1

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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
    
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


Dim KK As String

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
           " Where a.单据编号='" + UCase(Trim(TxtBillNoGC.text)) + "'  order by 4 "

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
            
        
        For j = 3 To rs.Fields.count - 1
             
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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
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
'查询报表名的序号

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim KK As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据

    strDatas = "NO,Sub Name,Ship To,Fab Device,Customer Device,GC Version,PO NO,Invoice NO,Ship Out Date,FAB Lot ID," & _
               "Wafer ID,Gross Dies,Sampling Qty,Pass Dies,NG Die,Yield,Pack Lot ID,Wafer Mark,Grade,Carton NO,WO,Remark" & vbCrLf
               
    
    '明细数据
'    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
'           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
'           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
'           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
'           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
'           " Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "
           
           
    strSql = " select  rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion,  cast([NO] as int), " & _
             " [Sub_Name],'GCSH' as [Ship_To],[Fab_Device],[Customer_Device],[GC_Version], " & _
             " [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [PACK_Out_Date],[FAB_Lot_ID],[Wafer_ID],[Gross_Dies], " & _
             " '' as [Sampling Qty] ,[Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as  [Pass_Dies],[NG_Die],[Yield] ,[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark] " & _
             " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGC.text)) + "'   order by 4 "
 
           
           

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
            
        
        For j = 3 To rs.Fields.count - 1
             
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
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@htkjks.com"
    strRecipientCC = "wanli.ma@htkjks.com"
        
    Call MailDetail("GC 发货报表", strRecipient, g_Path & "\" & "PL_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub



Private Sub SaveFileSendGCCOGR1()
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
'查询报表名的序号

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim KK As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_COG_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据

    strDatas = "NO,Sub Name,Ship To,Vacuum Bag ID,Vacuum Bag Qty,Box ID,Customer Device,GC Version," & _
                "Grade,Bond Pro,Ship Out Date,PO NO,Carton NO,WO,Invoice NO,Remark" & vbCrLf
       
               
    
    '明细数据
'    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
'           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
'           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
'           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
'           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
'           " Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "
           
           
    strSql = " select  rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion,  cast([NO] as int), " & _
             " [Sub_Name],'GCSH' as [Ship_To],[Fab_Device],[Customer_Device],[GC_Version], " & _
             " [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [PACK_Out_Date],[FAB_Lot_ID],[Wafer_ID],[Gross_Dies], " & _
             " '' as [Sampling Qty] ,[Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as  [Pass_Dies],[NG_Die],[Yield] ,[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark] " & _
             " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGC.text)) + "'   order by 4 "
 
           
           

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
            
        
        For j = 3 To rs.Fields.count - 1
             
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
    strDatas = strDatas + strRowData '数据连接
    '写入文件
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '发邮件
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "jiayun.zhang@htkjks.com"
    strRecipientCC = "wanli.ma@htkjks.com"
        
    Call MailDetail("GC 发货报表", strRecipient, g_Path & "\" & "PL_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
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
'查询报表名的序号

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim KK As String

'PL_HTKS_WLT_20151111001.csv
    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_WLT_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "NO,Sub Name,Ship To,Fab Device,Customer Device,GC Version,PO NO,Invoice NO,Ship Out Date,FAB Lot ID," & _
               "Wafer ID,Gross Dies,Sampling Qty,Pass Dies1,Pass Dies2, Pass Dies3,NG Die,Yield,Pack Lot ID,Wafer Mark,Grade,Carton NO,WO,Remark" & vbCrLf
    '明细数据
'    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
'           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
'           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
'           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
'           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
'           " Where a.单据编号='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "
           
           
           
'     strSql = "  select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,left(rtrim(ltrim(GC_Version)),2)+erpdata.dbo.GET_TSV_DOUBLECODE(rtrim(FAB_Lot_ID)+rtrim(Wafer_ID)) as gcversion, cast([NO] as int), " & _
'    " [Sub_Name],'GCSH' as [Ship_To],[Fab_Device],[Customer_Device],[GC_Version], " & _
'    " [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [PACK_Out_Date],[FAB_Lot_ID],CASE WHEN CHARINDEX('-',Wafer_ID)>0 THEN RIGHT(Wafer_ID,2) WHEN CHARINDEX('+',Wafer_ID)>0 THEN RIGHT(Wafer_ID,2) ELSE Wafer_ID END AS Wafer_ID,[Gross_Dies]," & _
'    " erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as [Sampling Qty] ,[Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as  [Pass_Dies],[NG_Die],[Yield] ,[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark]" & _
'    " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGCWlt.Text)) + "'  order by 4 "
               
               
               
    
    'strsql = " select rtrim(ltrim(FAB_Lot_ID)) + rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,a.GC_Version,left(rtrim(ltrim(GC_Version)), 2) +erpdata.dbo.GET_TSV_DOUBLECODE(rtrim(FAB_Lot_ID) + rtrim(Wafer_ID)) as gcversion,cast( NO  as int)," & _
    '         " Sub_Name ,'GCSH' as  Ship_To ,Fab_Device ,Customer_Device ,GC_Version ," & _
    '         " PO_NO ,Invoice_NO ,replace(PACK_Out_Date , '/', '-') as  PACK_Out_Date ,FAB_Lot_ID ,CASE WHEN CHARINDEX('-', Wafer_ID) > 0 THEN RIGHT(Wafer_ID, 2) WHEN CHARINDEX('+', Wafer_ID) > 0 THEN RIGHT(Wafer_ID, 2) Else Wafer_id END AS Wafer_ID,Gross_Dies ," & _
    '         " Pass_Dies,NG_Die ,Yield ,PACK_Lot_ID ,Wafer_Mark ,'A' as Grade,CartonNO ,WO ,Remark FROM  erpdata..GR_GC_DetailHistory  a Where a.单据编号 = 'F1708030012' order by 4"
               

'       strSql = "  select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,left(rtrim(ltrim(GC_Version)),2)+erpdata.dbo.GET_TSV_DOUBLECODE(rtrim(FAB_Lot_ID)+rtrim(Wafer_ID)) as gcversion, cast([NO] as int), " & _
'    " [Sub_Name],'GCSH' as [Ship_To],[Fab_Device],[Customer_Device],[GC_Version], " & _
'    " [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [PACK_Out_Date],[FAB_Lot_ID],CASE WHEN CHARINDEX('-',Wafer_ID)>0 THEN RIGHT(Wafer_ID,2) WHEN CHARINDEX('+',Wafer_ID)>0 THEN RIGHT(Wafer_ID,2) ELSE Wafer_ID END AS Wafer_ID,[Gross_Dies], erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as [Sampling Qty], " & _
'    " [Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) - erpdata.dbo.Get_TSV_GC_WaferGDieBin2(rtrim(ltrim(FAB_Lot_ID)) + rtrim(ltrim(Wafer_ID))) as  [Pass_Dies1],erpdata.dbo.Get_TSV_GC_WaferGDieBin2(rtrim(ltrim(FAB_Lot_ID)) + rtrim(ltrim(Wafer_ID))) as [Pass_Dies2],'' as [Pass_Dies3], [NG_Die],[Yield],[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark]" & _
'    " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + UCase(Trim(TxtBillNoGCWlt.Text)) + "'  order by 4 "
'
'

'
    strSql = "SELECT row_number() over(order by t.lot_id,t.waferid) AS rounum ,t.sub_name,t.SHIP_SITE,t.FAB_CONV_ID,t.cust_device,t.gcversion,t.PO_NUM,t.invoice,t.create_date,t.lot_id,t.waferid,t.gross_die,ISNULL(t.BIN3,L.NDPW) as Sampling_Qty " & _
",ISNULL(ISNULL(t.BIN1,t.A),K.NDPW) as Pass_Dies1,ISNULL(T.BIN2,m.NDPW) as Pass_Dies2,'' AS Pass_Dies3,ISNULL(ISNULL(T.E,n.NDPW),0) as NG_Die,CONVERT(VARCHAR(10),CONVERT(decimal(18,2), (t.gross_die - ISNULL(ISNULL(T.E,n.NDPW),0))*1.0/(t.gross_die )*100)) + '%' AS Yield   " & _
",ISNULL(t.sfc,k.FIRSTNAME) as Pack_Lot_ID,t.PRODUCTID,'A' as Grade,rtrim(t.箱号) as 箱号,t.大工单 FROM ( SELECT  'HTKS' AS sub_name,d.SHIP_SITE,RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID " & _
",a.cust_device,a.gcversion,d.PO_NUM,'' AS invoice, a.create_date,rtrim(a.lot_id) as lot_id,SUBSTRING(REPLACE(b.流程卡编号,'+',''),LEN(a.lot_id)+1,2) as waferid,c.FAILBINCOUNT+c.PASSBINCOUNT AS gross_die " & _
",e.GRADES,e.QTY,c.PRODUCTID,'A' as Grade,rtrim(ay.箱号) as 箱号,b.大工单,a.qbox,b.流程卡编号,SUBSTRING( e.SFC,12,8) AS SFC " & _
"FROM erptemp..tblshipreport_new a INNER JOIN erpdata..tblStockNumTree ax  ON ax.箱号 =a.qbox  INNER JOIN erpdata..tblStockNumTree ay  ON ay.序号 = ax.上级序号 " & _
"INNER JOIN  erpdata..tblStocksqfhsub b  ON b.单据编号 = a.ship_order AND b.箱号 = a.qbox AND b.工单号 = a.lot_id " & _
"INNER JOIN  ERPBASE..tblmappingData c ON  c.SUBSTRATEID = b.流程卡编号 AND c.LOTID = b.工单号 INNER JOIN  erpbase..tblCustomerOI d ON  CONVERT(VARCHAR(20), d.ID) = c.FILENAME AND d.SOURCE_BATCH_ID = c.LOTID  " & _
"left JOIN  erptemp..WAFER_BIN_LIST e ON e.WAFER_ID = b.流程卡编号 inner JOIN erpdata..tblErpInStockRelation ee ON ee.SFC_ID = e.SFC  AND CHARINDEX(e.WAFER_ID,ee.WAFER_ID) <> 0    " & _
"WHERE a.ship_order = '" & UCase(Trim(TxtBillNoGCWlt.text)) & "' ) AS p PIVOT(sum(qty) FOR grades IN(A,BIN1,BIN2,BIN3, E)) AS T " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV k ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.流程卡编号 AND k.CONTAINERNAME LIKE '%-A' " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox AND L.WAFERSCRIBENUMBER = t.流程卡编号 AND L.CONTAINERNAME LIKE '%-A-01' " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV m ON m.QBOXNUMBER = t.qbox AND m.WAFERSCRIBENUMBER = t.流程卡编号 AND m.CONTAINERNAME LIKE '%-A-02' " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV n ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.流程卡编号 AND n.CONTAINERNAME LIKE '%-E' "



'    strSql = "SELECT t.FAB_CONV_ID,t.cust_device,t.gcversion,t.PO_NUM,t.invoice,t.create_date,t.lot_id,t.waferid,t.gross_die,ISNULL(t.BIN3,L.NDPW) as Sampling_Qty " & _
'",ISNULL(ISNULL(t.BIN1,t.A),K.NDPW) as Pass_Dies1,ISNULL(T.BIN2,m.NDPW) as Pass_Dies2,'' AS Pass_Dies3,ISNULL(ISNULL(T.E,n.NDPW),0) as NG_Die,CONVERT(VARCHAR(10),CONVERT(decimal(18,2), (t.gross_die - ISNULL(ISNULL(T.E,n.NDPW),0))*1.0/(t.gross_die )*100)) + '%' AS Yield   " & _
'",ISNULL(t.sfc,k.FIRSTNAME) as Pack_Lot_ID,t.PRODUCTID,'A' as Grade,rtrim(t.箱号) as 箱号,t.大工单 FROM ( SELECT  'HTKS' AS sub_name,d.SHIP_SITE,RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID " & _
'",a.cust_device,a.gcversion,d.PO_NUM,'' AS invoice, a.create_date,rtrim(a.lot_id) as lot_id,SUBSTRING(REPLACE(b.流程卡编号,'+',''),LEN(a.lot_id)+1,2) as waferid,c.FAILBINCOUNT+c.PASSBINCOUNT AS gross_die " & _
'",e.GRADES,e.QTY,c.PRODUCTID,'A' as Grade,rtrim(ay.箱号) as 箱号,b.大工单,a.qbox,b.流程卡编号,SUBSTRING( e.SFC,12,8) AS SFC " & _
'"FROM erptemp..tblshipreport_new a INNER JOIN erpdata..tblStockNumTree ax  ON ax.箱号 =a.qbox  INNER JOIN erpdata..tblStockNumTree ay  ON ay.序号 = ax.上级序号 " & _
'"INNER JOIN  erpdata..tblStocksqfhsub b  ON b.单据编号 = a.ship_order AND b.箱号 = a.qbox AND b.工单号 = a.lot_id " & _
'"INNER JOIN  ERPBASE..tblmappingData c ON  c.SUBSTRATEID = b.流程卡编号 AND c.LOTID = b.工单号 INNER JOIN  erpbase..tblCustomerOI d ON  CONVERT(VARCHAR(20), d.ID) = c.FILENAME AND d.SOURCE_BATCH_ID = c.LOTID  " & _
'"left JOIN  erptemp..WAFER_BIN_LIST e ON e.WAFER_ID = b.流程卡编号 inner JOIN erpdata..tblErpInStockRelation ee ON ee.SFC_ID = e.SFC  AND CHARINDEX(e.WAFER_ID,ee.WAFER_ID) <> 0    " & _
'"WHERE a.ship_order = '" & UCase(Trim(TxtBillNoGCWlt.Text)) & "' ) AS p PIVOT(sum(qty) FOR grades IN(A,BIN1,BIN2,BIN3, E)) AS T " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV k ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.流程卡编号 AND k.CONTAINERNAME LIKE '%-A' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox AND L.WAFERSCRIBENUMBER = t.流程卡编号 AND L.CONTAINERNAME LIKE '%-A-01' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV m ON m.QBOXNUMBER = t.qbox AND m.WAFERSCRIBENUMBER = t.流程卡编号 AND m.CONTAINERNAME LIKE '%-A-02' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV n ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.流程卡编号 AND n.CONTAINERNAME LIKE '%-E' "
'





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
            
'            waferVerResult = GetGCOutRpt_Ver(waferidMain, waferPT, waferVer)
'
'            If Len(waferVerResult) <> 3 Then
'                MsgBox waferidMain & " 此片二级代码长度不等于3，请确认好后才能导报表！", vbInformation, "友情提示"
'                'Exit Sub
'
'            End If
            
        
        For j = 0 To rs.Fields.count - 1
             
'             If j = 8 Then
'
'             strColData = strColData + waferVerResult + ","
''
'             If j = 11 Then
'
'             dateTemp = Trim("" & rs.Fields(j).Value)
'
'                strColData = strColData + Format(dateTemp, "YYYY-MM-DD") + ","
'
'             Else
             
             
             strColData = strColData + Trim("" & rs.Fields(j).Value) + ","

'             End If
           
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
    
    strRecipient = "jiayun.zhang@htkjks.com"
    strRecipientCC = "wanli.ma@htkjks.com"
        
    Call MailDetail("GC 发货报表", strRecipient, g_Path & "\" & "PL_HTKS_WLT_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub


Private Sub SaveFileSendCOG_01()
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
'查询报表名的序号

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim KK As String

'PL_HTKS_WLT_20151111001.csv
    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_COG_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "NO,Sub Name,Ship To,Vacuum Bag ID,Vacuum Bag Qty,Box ID,Customer Device," & _
               "GC Version,Grade,Bond Pro,Ship Out Date,PO NO,Carton NO,WO,Invoice NO,Remark" & vbCrLf
               
'strsql = " select ROW_NUMBER() OVER(ORDER BY X.Qboxnumber ,X.lvNo) as No,Sub_Name,ShipTo,lvNo,VacuumBagQty,Qboxnumber,Customer_Device,GCVersion, " & _
'" Grade,BondPro,ShipOutDate,PO_NO,CartonNO,WOnumber,InvoiceNO,Remark from ( " & _
'" select  distinct  Sub_Name,'GC_SH' as ShipTo,e.[LVBARCODEQBOX] as lvNo,CASE WHEN dbo.GET_TSV_COG_LOTQTY(A.FAB_Lot_ID,E.LVBARCODEQBOX)=0 THEN 1500 ELSE dbo.GET_TSV_COG_LOTQTY(A.FAB_Lot_ID,E.LVBARCODEQBOX)END as VacuumBagQty,c.箱号 as Qboxnumber, " & _
'" Customer_Device,GC_Version+'D' as GCVersion,'A' as Grade,'SH' as BondPro,REPLACE(PACK_Out_Date,'/','-') as ShipOutDate,PO_NO," & _
'" '1' as CartonNO,d.MTRL_NUM as WOnumber,'' as InvoiceNO,''as Remark " & _
'" from  [erpdata].[dbo].[GR_GC_DetailHistory] a ,[erpdata].[dbo].[tblPackTreeInf] b  , [erpdata].[dbo].[tblPackTreeInf] c,[ERPBASE].[dbo].[tblCustomerOI] d,[erpdata].[dbo].[GR_COG_LV_Data] e   " & _
'" Where a.单据编号='" + UCase(Trim(TxtCogNo.Text)) + "'" & _
' "   and b.箱号=a.CartonNO and c.上级序号=b.序号 and d.SOURCE_BATCH_ID=a.FAB_Lot_ID and d.PO_NUM=a.PO_NO and e.[QBOXNUMBER]=c.箱号 ) X"
'
 strSql = "SELECT ROW_NUMBER() OVER(ORDER BY A.LBCODE ,A.LVCODE)AS NO,A.*  FROM TBLPL_HTKS_COG_report A"
               

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
      
            
        
        For j = 0 To rs.Fields.count - 1
             
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
    
    strRecipient = "ks015918@ht-tech.com"
    strRecipientCC = "wanli.ma@htkjks.com"
        
    Call MailDetail("GC COG 发货报表", strRecipient, g_Path & "\" & "PL_HTKS_COG_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtCogNo.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub



Private Sub SaveFileSendCOG_02()
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
'查询报表名的序号

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim KK As String

'PL_HTKS_WLT_20151111001.csv
    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_COG_TRC_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
    strDatas = "BOXA_ID,BOXB_ID,CHIP_QTY,GRADE,BOX_TYPE,SERIAL_NUM" & vbCrLf
           
               
    strSql = " select TrayNo ,LotID  ,Qty  ,Grade ,BoxType ,SERIAL_NUM from  TSV_GR_COG_Tray_Data order by id "
            '   strSql = " select BOXA_ID ,BOXB_ID  ,CHIP_QTY  ,GRADE ,BOXTYPE ,SERIAL_NUM from  TSV_GR_COG_TRAY_DATA_1"
               
    ' tangwei: 20171010 上下对调

    strRowData = ""
    If rs.State = adStateOpen Then rs.Close
    If Cnn.State = 0 Then
    ConOracle
    End If
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
      
            
        
        For j = 0 To rs.Fields.count - 1
             
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
    
    strRecipient = "jiayun.zhang@htkjks.com"
    strRecipientCC = "wanli.ma@htkjks.com"
        
    Call MailDetail("GC COG 发货报表", strRecipient, g_Path & "\" & "PL_HTKS_COG_TRC_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtCogNo.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
ErrHandler:
    Set FSO = Nothing
End Sub

Private Sub SaveFileSendCOG_03()
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
'查询报表名的序号

'fileNo = GetGC_FileNoNew("GC")
'waferidMain = ""
'waferPT = ""
'waferVer = ""
'waferVerResult = ""


Dim KK As String

    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PLP_ERP_COG_HTKJ_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
               
    
    strSql = "select * from tbl_PLP_ERP_COG_HTKJ"
    If INIadoCon.State <> adStateOpen Then
    INIConnectSTART2
    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
If rs.RecordCount > 0 Then
               
   
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
      
            
        
        For j = 0 To rs.Fields.count - 1
             
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
    
    strRecipient = "jiayun.zhang@htkjks.com"
    strRecipientCC = "wanli.ma@htkjks.com"
        
    Call MailDetail("GC COG 发货报表", strRecipient, g_Path & "\" & "PLP_ERP_COG_HTKJ_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtCogNo.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
   End If
    
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


Dim KK As String

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
            
            
        
        For j = 3 To rs.Fields.count - 1
             
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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "发送成功！", vbInformation, "友情提示"
    
    
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
'查询报表名的序号

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim KK As String

    '创建文件
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_WLD_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    '写数据
    strDatas = ""
    '头数据
'    strDatas = "No.,Sub Name,Ship To,Customer Device,GC Version,CST ID,CST QTY,Bond Pro.,FAB Lot ID,Wafer ID,Wafer Mark,Gross Dies" & _
'               ",PO NO,WO,Invoice NO,FAB Device,Pack lot ID,FAB-Out Date,Sampling Qty,Pass Dies,Yield" & vbCrLf
    
    
   strDatas = "NO,Sub Name,Ship To,Fab Device,Customer Device,GC Version,PO NO,Invoice NO,Ship Out Date,FAB Lot ID," & _
               "Wafer ID,Gross Dies,Sampling Qty,Pass Dies,NG Die,Yield,Pack Lot ID,Wafer Mark,Grade,Carton NO,WO,Remark" & vbCrLf
    
    
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
            
            
        
        For j = 3 To rs.Fields.count - 1
             
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
    
    strRecipient = "jiayun.zhang@htkjks.com"
    strRecipientCC = "wanli.ma@htkjks.com"
        
    Call MailDetail("GC 发货报表", strRecipient, g_Path & "\" & "PL_HTKS_WLD_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '把发送记录保存到DB中
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
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
    xlSheet.name = "GrData"
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
tempBillNo = UCase(Trim(TxtBillNo.text))

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
        For j = 0 To rs.Fields.count - 1
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


comBo2.AddItem ("GC")
comBo2.AddItem ("SX")
comBo2.AddItem ("HD")


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


tempBillNo = UCase(Trim(TxtBillNoGC.text))
custNameTemp = UCase(Trim(comBo2.text))

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
           
'sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [Sub Name],[Ship_To]as [Ship To] ,[Fab_Device]as [Fab Device] ,[Customer_Device] as [Customer Device],[PO_NO] as [PO NO]," & _
'          " [WO],[GC_Version]as [GC Version],[Invoice_NO]as [Invoice NO] ,[PACK_Out_Date]as[PACK-Out Date],[PACK_Lot_ID]as[PACK Lot ID],[FAB_Lot_ID]as[FAB Lot ID] ," & _
'          " [Wafer_ID]as [Wafer ID],[Wafer_Mark]as [Wafer Mark],[Gross_Dies]as [Gross Dies],[Pass_Dies]as [Pass Dies],[NG_Die]as [NG Die] ,[Yield] ," & _
'          " [Remark] , [System_CartonNO]as [System CartonNO], [PACK_Device]as [PACK Device], [CartonNO]as [CartonNO], [MaskType] " & _
'          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + tempBillNo + "' order by 1  "
'
'
          
sqlTemp = "  select cast([NO] as int) as NO, " & _
" [Sub_Name],'GCSH' as [Ship To],[Fab_Device],[Customer_Device],[GC_Version], " & _
" [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [Ship Out Date],[FAB_Lot_ID],[Wafer_ID],[Gross_Dies], " & _
" '' as [Sampling Qty] ,[Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as  [Pass_Dies],[NG_Die],[Yield] ,[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark] " & _
" FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.单据编号='" + tempBillNo + "'  order by a.CartonNO "
           
          
          
          
          
          
          
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



'正常发送
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtBillNoGC.text))
custNameTemp = UCase(Trim(comBo2.text))


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
txtPackage.SetFocus
End If

End Sub



