VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmGCGrV2 
   Caption         =   "GC�ͻ�������Ϣ �°��ʽ"
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
      TabCaption(0)   =   "������Ʒ����"
      TabPicture(0)   =   "FrmGCGrV2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TxtBillNoGC"
      Tab(0).Control(1)=   "GCCmdOut"
      Tab(0).Control(2)=   "GCCmdSend"
      Tab(0).Control(3)=   "Combo2"
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(5)=   "Label5"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "WLT����"
      TabPicture(1)   =   "FrmGCGrV2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "TxtBillNoGCWlt"
      Tab(1).Control(2)=   "Label1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "WLA MES���� ����"
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
      TabCaption(3)   =   "WLA ��ERPϵͳ ����"
      TabPicture(3)   =   "FrmGCGrV2.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "TxtBillNoGCWLAErp"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "WLD����"
      TabPicture(4)   =   "FrmGCGrV2.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command4"
      Tab(4).Control(1)=   "TxtBillNoGCWLDErp"
      Tab(4).Control(2)=   "Label8"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "COG����"
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
         Caption         =   "���ͱ���"
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
         Caption         =   "�ܻ�������"
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
         Caption         =   "���ͱ���"
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
         Caption         =   "����Excel"
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
         Caption         =   "����Excel"
         Height          =   480
         Left            =   -72600
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "���ͱ���"
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
         Caption         =   "��������"
         Height          =   480
         Left            =   -69600
         TabIndex        =   3
         Top             =   1920
         Width           =   990
      End
      Begin VB.CommandButton GCCmdSend 
         Caption         =   "���ͱ���"
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
         Caption         =   "ע�����ܻ������ݣ���ϵͳ��ʾ�ɹ�����ѡ�񱨱����ͣ��ٷ��ͱ���"
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
         Caption         =   "�������ͣ�"
         Height          =   195
         Left            =   -74520
         TabIndex        =   26
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݱ�ţ�"
         Height          =   195
         Left            =   -74520
         TabIndex        =   25
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݱ�ţ�"
         Height          =   195
         Left            =   -74640
         TabIndex        =   22
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݱ�ţ�"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ����֣� "
         Height          =   195
         Left            =   -73800
         TabIndex        =   16
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ���ڣ� "
         Height          =   195
         Left            =   -73800
         TabIndex        =   15
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ڣ�"
         Height          =   195
         Left            =   -73800
         TabIndex        =   14
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݱ�ţ�"
         Height          =   195
         Left            =   -74640
         TabIndex        =   9
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݱ�ţ�"
         Height          =   195
         Left            =   -74640
         TabIndex        =   6
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ���"
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
Private Enum E_FPS0          'Detail�֭�
'    E_ID = 1                 'id��
    E_Key = 1                'Key
    E_Value                  'Value
    E_getValue               'getValue
    E_otherValue             '��ע
    E_END
    
End Enum
Dim reportRS As New ADODB.Recordset

Public g_Date           As String



Private Sub cmdADD_Click()
'����
Dim tempKey As String
Dim tempValue As String
Dim getValue As String
Dim otherValue As String

Dim sqlTemp As String

tempKey = UCase(Trim(txtdelNote.text))
tempValue = Trim(txtawb.text)
getValue = CombMo.text
otherValue = Trim(txtPackage.text)

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

Private Sub CmdOut_Click()
Dim tempBillNo As String
tempBillNo = UCase(Trim(TxtBillNo.text))

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
    


 Dim sqlTemp As String

 sqlTemp = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & _
           " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & _
           " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & _
           " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & _
           " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & _
           " Where a.���ݱ�� = b.���ݱ�� and a.���ݱ��='" + tempBillNo + "' "

  SqlServerExporToExcel (sqlTemp)


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
    MsgBox "��������������!", vbInformation, "������ʾ"
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
tempBillNo = UCase(Trim(TxtBillNo.text))

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

Private Sub CmdCOGOutRpt_Click()

Call Command6_Click

'����
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtCogNo.text))



If tempBillNo = "" Then
    MsgBox "�����뵥�ݱ�ţ��ٵ��ܱ������ݣ�����ٷ��ͱ���", vbInformation, "������ʾ"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(tempBillNo)
 If judgeEmp = False Then
 MsgBox "��ѯ�����˵��ݱ�ŵ������Ϣ����ȷ��!", vbInformation, "������ʾ"
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
'WLT ����  2015-11-11

Dim strSql As String
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtBillNoGCWlt.text))

If tempBillNo = "" Then
    MsgBox "�����뵥�ݱ��!", vbInformation, "������ʾ"
    Exit Sub
End If

Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(tempBillNo)
 If judgeEmp = False Then
 MsgBox "��ѯ�����˵��ݱ�ŵ������Ϣ����ȷ��!", vbInformation, "������ʾ"
' Exit Sub
 
End If

' strSql = "SELECT row_number() over(order by t.lot_id,t.waferid) AS NO ,t.sub_name,t.SHIP_SITE,t.FAB_CONV_ID,t.cust_device,t.gcversion,t.PO_NUM,t.invoice,t.create_date,t.lot_id,t.waferid,t.gross_die,ISNULL(t.BIN3,L.NDPW) as Sampling_Qty " & _
'",ISNULL(ISNULL(t.BIN1,t.A),K.NDPW) as Pass_Dies1,ISNULL(T.BIN2,m.NDPW) as Pass_Dies2,'' AS Pass_Dies3,ISNULL(ISNULL(T.E,n.NDPW),0) as NG_Die,CONVERT(VARCHAR(10),CONVERT(decimal(18,2), (t.gross_die - ISNULL(ISNULL(T.E,n.NDPW),0))*1.0/(t.gross_die )*100)) + '%' AS Yield   " & _
'",ISNULL(t.sfc,k.FIRSTNAME) as Pack_Lot_ID,t.PRODUCTID,'A' as Grade,rtrim(t.���) as ���,t.�󹤵� FROM ( SELECT  'HTKS' AS sub_name,d.SHIP_SITE,RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID " & _
'",a.cust_device,a.gcversion,d.PO_NUM,'' AS invoice, a.create_date,rtrim(a.lot_id) as lot_id,SUBSTRING(REPLACE(b.���̿����,'+',''),LEN(a.lot_id)+1,2) as waferid,c.FAILBINCOUNT+c.PASSBINCOUNT AS gross_die " & _
'",e.GRADES,e.QTY,c.PRODUCTID,'A' as Grade,rtrim(ay.���) as ���,b.�󹤵�,a.qbox,b.���̿����,SUBSTRING( e.SFC,12,8) AS SFC " & _
'"FROM erptemp..tblshipreport_new a INNER JOIN erpdata..tblStockNumTree ax  ON ax.��� =a.qbox  INNER JOIN erpdata..tblStockNumTree ay  ON ay.��� = ax.�ϼ���� " & _
'"INNER JOIN  erpdata..tblStocksqfhsub b  ON b.���ݱ�� = a.ship_order AND b.��� = a.qbox AND b.������ = a.lot_id " & _
'"INNER JOIN  ERPBASE..tblmappingData c ON  c.SUBSTRATEID = b.���̿���� AND c.LOTID = b.������ INNER JOIN  erpbase..tblCustomerOI d ON  CONVERT(VARCHAR(20), d.ID) = c.FILENAME AND d.SOURCE_BATCH_ID = c.LOTID  " & _
'"left JOIN  erptemp..WAFER_BIN_LIST e ON e.WAFER_ID = b.���̿���� inner JOIN erpdata..tblErpInStockRelation ee ON ee.SFC_ID = e.SFC  AND CHARINDEX(e.WAFER_ID,ee.WAFER_ID) <> 0    " & _
'"WHERE a.ship_order = '" & UCase(Trim(TxtBillNoGCWlt.Text)) & "' ) AS p PIVOT(sum(qty) FOR grades IN(A,BIN1,BIN2,BIN3, E)) AS T " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV k ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.���̿���� AND k.CONTAINERNAME LIKE '%-A' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox AND L.WAFERSCRIBENUMBER = t.���̿���� AND L.CONTAINERNAME LIKE '%-A-01' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV m ON m.QBOXNUMBER = t.qbox AND m.WAFERSCRIBENUMBER = t.���̿���� AND m.CONTAINERNAME LIKE '%-A-02' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV n ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.���̿���� AND n.CONTAINERNAME LIKE '%-E' "
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

'ERP�ĵ���
Dim billNoTemp As String
Dim sqlTemp As String

billNoTemp = Trim(TxtBillNoGCWLAErp.text)
  
If Left$(billNoTemp, 3) = "FDP" Then
'    sqlTemp = "SELECT  row_number() OVER(ORDER BY X.[CST ID],X.[Wafer ID]) AS [NO],X.* FROM ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'"[erpdata].[dbo].[Get_TSV_GCWLA_LotID_NewDB](a.�������,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST ID],[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_NewDB](a.�������,rtrim(ltrim(a.������)), " & _
'" rtrim(ltrim(a.���̿����))) as [Wafer Qty], 'SH' as [Bond Pro],B.PO_NUM AS [PO NO],'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.������ as [FAB Lot ID], " & _
'" right(rtrim(ltrim(a.���̿����)),2) as [Wafer ID], a.�ϸ��� as [Gross Dies], '0' as [Sampling Qty],a.�ϸ��� as [Pass Dies],0 as [NG Die],''as [Yield],c.FIRSTNAME as [Pack Lot ID], d.productid as [Wafer Mark], " & _
'" 'A' AS Grade,c.QBOXNUMBER as [Carton NO],f.ORDERNAME as WO,'' as [Remark] FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV c  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.������� ='" & billNoTemp & "' " & _
'" and b.SOURCE_BATCH_ID=a.������ and d.filename = cast(b.ID as nvarchar) and c.WAFERSCRIBENUMBER=a.���̿���� and c.WAFERNUMBER=a.������ and d.SUBSTRATEID=a.���̿���� and f.WAFERID=a.���̿���� )X"
    
'    sqlTemp = "SELECT  row_number() OVER(ORDER BY X.[CST ID],X.[Wafer ID]) AS [NO],X.* FROM ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_NewDB](a.�������,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST ID],[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_NewDB](a.�������,rtrim(ltrim(a.������)), " & _
'" rtrim(ltrim(a.���̿����))) as [Wafer Qty], 'SH' as [Bond Pro],B.PO_NUM AS [PO NO],'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.������ as [FAB Lot ID], " & _
'" right(rtrim(ltrim(a.���̿����)),2) as [Wafer ID], a.�ϸ��� as [Gross Dies], '0' as [Sampling Qty],a.�ϸ��� as [Pass Dies],0 as [NG Die],''as [Yield],c.FIRSTNAME as [Pack Lot ID], d.productid as [Wafer Mark], " & _
'" 'A' AS Grade,c.QBOXNUMBER as [Carton NO],f.ORDERNAME as WO,'' as [Remark] FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV c  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.������� ='" & billNoTemp & "' " & _
'" and b.SOURCE_BATCH_ID=a.������ and d.filename = cast(b.ID as nvarchar) and c.WAFERSCRIBENUMBER=a.���̿���� and c.WAFERNUMBER=a.������ and d.SUBSTRATEID=a.���̿���� and f.WAFERID=a.���̿���� )X" & _
'" union SELECT  row_number() OVER(ORDER BY Y.[CST ID],Y.[Wafer ID]) AS [NO],Y.* FROM ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_NewDB](a.�������,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST ID],[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_NewDB](a.�������,rtrim(ltrim(a.������)), " & _
'" rtrim(ltrim(a.���̿����))) as [Wafer Qty], 'SH' as [Bond Pro],B.PO_NUM AS [PO NO],'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.������ as [FAB Lot ID], " & _
'" right(rtrim(ltrim(replace(a.���̿����,'+',''))),2) as [Wafer ID], a.�ϸ��� as [Gross Dies], '0' as [Sampling Qty],a.�ϸ��� as [Pass Dies],0 as [NG Die],''as [Yield],REPLACE(BB.SFC_ID,'SFCBO:1020,','') as [Pack Lot ID], d.productid as [Wafer Mark], " & _
'" 'A' AS Grade,A.��� as [Carton NO],f.ORDERNAME as WO,'' as [Remark] FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata..tblErpInStockDetailInfo aa,erpdata..tblErpInStockRelation bb  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.������� ='" & billNoTemp & "' " & _
'" and b.SOURCE_BATCH_ID=a.������ and d.filename = cast(b.ID as nvarchar)  and d.SUBSTRATEID=a.���̿���� and f.WAFERID=a.���̿����  and a.��� = aa.KEY_VALUE and bb.BOX_ID = aa.BOX_ID and  SUBSTRING(replace(bb.WAFER_ID,'SFCBO:1020,','') " & _
'", CHARINDEX(',', replace(bb.WAFER_ID,'SFCBO:1020,',''))+1,CHARINDEX('::', replace(bb.WAFER_ID,'SFCBO:1020,',''))-CHARINDEX(',', replace(bb.WAFER_ID,'SFCBO:1020,',''))-1) = a.���̿���� )Y"
'


        sqlTemp = "SELECT  row_number() OVER(ORDER BY X.[CST ID],X.[Wafer ID]) AS [NO],X.* FROM ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_NewDB](a.�������,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST ID],[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_NewDB](a.�������,rtrim(ltrim(a.������)), " & _
" rtrim(ltrim(a.���̿����))) as [Wafer Qty], 'SH' as [Bond Pro],B.PO_NUM AS [PO NO],'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.������ as [FAB Lot ID], " & _
" right(rtrim(ltrim(a.���̿����)),2) as [Wafer ID], a.�ϸ��� as [Gross Dies], '0' as [Sampling Qty],a.�ϸ��� as [Pass Dies],0 as [NG Die],''as [Yield],c.FIRSTNAME as [Pack Lot ID], d.productid as [Wafer Mark], " & _
" 'A' AS Grade,c.QBOXNUMBER as [Carton NO],f.ORDERNAME as WO,'' as [Remark] , a.���̿����  FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV c  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.������� ='" & billNoTemp & "' " & _
" and b.SOURCE_BATCH_ID=a.������ and d.filename = cast(b.ID as nvarchar) and c.WAFERSCRIBENUMBER=a.���̿���� and c.WAFERNUMBER=a.������ and d.SUBSTRATEID=a.���̿���� and f.WAFERID=a.���̿���� )X  union " & _
"   SELECT row_number() OVER(ORDER BY Y.[FAB Lot ID], Y.[Wafer ID]) AS NO,Y.* FROM (SELECT DISTINCT 'HTKS' as  'Sub Name' , 'GC_LG' as 'Ship To' , " & _
"    B.FAB_CONV_ID AS  'Fab Device' , replace(b.MPN_DESC, '-3', '-2.5') as 'Customer Device' ,  b.IMAGER_CUSTOMER_REV as 'GC Version','' AS [CST ID], " & _
"   '' as [Wafer Qty] , 'SH' as 'Bond Pro', B.PO_NUM AS 'PO NO',''  AS 'Invoice NO', convert(varchar(100), getdate(), 111) AS 'Ship Out Date', " & _
"   A.������ as 'FAB Lot ID', right(rtrim(ltrim(replace(a.���̿����, '+', ''))), 2) as [Wafer ID], a.�ϸ��� as 'Gross Dies','0' as 'Sampling Qty', " & _
"  a.�ϸ��� AS 'Pass Dies',0 as 'NG Die','' as 'Yield',SUBSTRING( REPLACE(ab.KEYID, 'SFCBO:1020,', ''),1 " & _
"  ,CHARINDEX(rtrim(a.���̿����),REPLACE(ab.KEYID, 'SFCBO:1020,', ''))-2) as 'Pack Lot ID',  d.productid as 'Wafer Mark', 'A' AS Grade, A.��� as 'Carton NO', " & _
"   f.ORDERNAME as WO, '' as 'Remark' , a.���̿���� FROM erpdata..tblStockdbsub a, ERPBASE..tblCustomerOI  b, erpdata..tblErpInStockDetailInfo aa,erpdata..tblErpInStockDetailInfo ab, " & _
"   ERPBASE..tblmappingData  d, erpdata..tblTSVwaferlist  f  WHERE a.������� = '" & billNoTemp & "' AND b.SOURCE_BATCH_ID = a.������ and d.filename = cast(b.ID as nvarchar) " & _
"  and d.SUBSTRATEID = a.���̿���� AND f.WAFERID = a.���̿���� and SUBSTRING(a.���,1,CASE WHEN CHARINDEX('_VT',A.��� )>0 THEN  CHARINDEX('_VT',A.��� )-1 ELSE len(A.���) END )= aa.KEY_VALUE  AND ab.BOX_ID = aa.BOX_ID AND ab.KEY_TYPE = 'WAFER' AND ab.KEY_VALUE = a.���̿���� ) Y "
    
    
    
   Call SqlServerExporToExcel_WLA_new(sqlTemp, billNoTemp)

    Exit Sub
End If
  
  
Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGCWlt(billNoTemp)
 If judgeEmp = False Then
 MsgBox "��ѯ�����˵��ݱ�ŵ������Ϣ����ȷ��!", vbInformation, "������ʾ"
 Exit Sub
 
End If

      'sqlTemp = "  SELECT row_number() OVER(ORDER BY a.������,a.���̿����) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
'" replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.���ݱ��,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST ID]," & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.���ݱ��,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST QTY]," & _
'" 'SH' as [Bond Pro.],a.������ as [FAB Lot ID],right(rtrim(ltrim(a.���̿����)),2) as [Wafer ID]," & _
'" a.���� as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
'" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.���� as [Sampling Qty]," & _
'" ''as [Pass Dies],''as [Yield],'A'as [Remark] " & _
'" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c WHERE a.���ݱ��='" + billNoTemp + "'" & _
'" and b.SOURCE_BATCH_ID=a.������ and c.WAFERSCRIBENUMBER=a.���̿���� and c.WAFERNUMBER=a.������ "


  sqlTemp = " SELECT  row_number() OVER(ORDER BY X.[CST ID],X.[Wafer ID]) AS [NO],X.* FROM " & _
 " ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
" B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
"[erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.���ݱ��,rtrim(ltrim(a.������)),rtrim(ltrim(a.���̿����))) as [CST ID], " & _
"[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.���ݱ��,rtrim(ltrim(a.������)), " & _
"   rtrim(ltrim(a.���̿����))) as [Wafer Qty], 'SH' as [Bond Pro],B.PO_NUM AS [PO NO], " & _
"   '' [Invoice NO],convert(varchar, getdate(), 111) AS [Ship Out Date], A.������ as [FAB Lot ID], " & _
 "  right(rtrim(ltrim(a.���̿����)),2) as [Wafer ID], a.���� as [Gross Dies], '0' as [Sampling Qty], " & _
 "  a.���� as [Pass Dies],0 as [NG Die],''as [Yield],c.FIRSTNAME as [Pack Lot ID], d.productid as [Wafer Mark], " & _
 "  'A' AS Grade,c.QBOXNUMBER as [Carton NO],f.ORDERNAME as WO,'' as [Remark] " & _
  "  FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b , " & _
  "   erpdata.dbo.TblQBOXNUMBER_TSV c  ,[ERPBASE].[dbo].[tblmappingData] d, " & _
   "  [erpdata].[dbo].[tblTSVwaferlist] f WHERE a.���ݱ��='" + billNoTemp + "' " & _
   "  and b.SOURCE_BATCH_ID=a.������ and d.filename = cast(b.ID as nvarchar) " & _
   "  and c.WAFERSCRIBENUMBER=a.���̿���� " & _
   "  and c.WAFERNUMBER=a.������ " & _
   "  and d.SUBSTRATEID=a.���̿���� " & _
   "  and f.WAFERID=a.���̿���� )X "
        
        
        
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
            MsgBox ("��ѯ��������!")
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
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.name = "����"
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

tempBillNo = UCase(Trim(TxtBillNoGCWLDErp.text))



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

Private Sub Command5_Click()
'CoG����

Dim tempBillNo As String
Dim custNameTemp As String


tempBillNo = UCase(Trim(TxtCogNo.text))


If tempBillNo = "" Then
    MsgBox "�����뵥�ݱ��!", vbInformation, "������ʾ"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGC(tempBillNo, "GC")
 If judgeEmp = False Then
 MsgBox "��ѯ�����˵��ݱ�ŵ������Ϣ����ȷ��!", vbInformation, "������ʾ"
 Exit Sub
 
End If
    


 Dim sqlTemp As String
      
 If custNameTemp = "GC" Then
           


          
sqlTemp = "  select cast([NO] as int) as NO, " & _
" [Sub_Name],'GCSH' as [Ship To],[Fab_Device],[Customer_Device],[GC_Version], " & _
" [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [Ship Out Date],[FAB_Lot_ID],[Wafer_ID],[Gross_Dies], " & _
" '' as [Sampling Qty] ,[Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as  [Pass_Dies],[NG_Die],[Yield] ,[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark] " & _
" FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "'  order by 1 "
           
          
          
          
          
   
End If

  SqlServerExporToExcel (sqlTemp)


End Sub

Private Sub Command6_Click()
'COG ��������

'��������
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
    MsgBox "��ѡ��ͻ����룬���뵥�ݱ��!", vbInformation, "������ʾ"
    Exit Sub
End If


  Dim judgeEmp As Boolean

judgeEmp = JudgeGRBillNoGC(tempBillNo, custNameTemp)
 If judgeEmp = False Then
 MsgBox "��ѯ�����˵��ݱ�ŵ������Ϣ����ȷ��!", vbInformation, "������ʾ"
 Exit Sub
 
End If


'����ERP�������ţ��е�С��ţ��ܳ�����

Call delCOGIniData      'ɾ����ǰ���� oracle cog С���

Call delCOGRptInt01      'ɾ����ǰ���� sqlserver  �ܱ���1�е���������Ϣ

Call delCOGRptInt02      'ɾ����ǰ���� oracle �ܱ���2�е�Tray��Ϣ

Set reportRS = GetCOGBaseData2(tempBillNo)

If (reportRS.RecordCount < 0) Then
    MsgBox "��ѯ�������ݣ� ��ȷ�ϵ��ݱ�ţ�", vbInformation, "������ʾ"
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


  
  '����������Ϣ�ܳ����ŵ�ERP��
  
Set reportRS = GetCOGLVDataList()

If (reportRS.RecordCount < 0) Then
    MsgBox "��ѯ�������ݣ� ��ϵITȷ����������ǩ��Ϣ�����쳣��", vbInformation, "������ʾ"
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
  
    'end ����������Ϣ�ܳ����ŵ�ERP��


Call AddCOGRptInt02      '���� oracle �ܱ���2�е�Tray��Ϣ
Call AddCOGRptInt02_2      '���� oracle �ܱ���2�е�Tray��Ϣ 02����
Call AddCOGRptInt02_3      '���� oracle �ܱ���2�е�Tray��Ϣ 03����

'PL_HTKS_COG cog�����������ִ��SQL SEVER�洢����
  Set adoCmd = New ADODB.Command
         Set adoCmd.ActiveConnection = INIadoCon2
             adoCmd.CommandText = "PL_HTKS_COG"
             adoCmd.Parameters.Refresh
             adoCmd.CommandType = adCmdStoredProc
        
          Set adoprm1 = New ADODB.Parameter   '����Ϊ������
          adoprm1.type = adChar
          adoprm1.Size = 20
          adoprm1.Direction = adParamInput
          adoprm1.Value = tempBillNo
          adoCmd.Parameters.Append adoprm1
          adoCmd.Execute
          
'PLP_ERP_COG_HTKJ cog�����������ִ��SQL SEVER�洢����
  Set adoCmd1 = New ADODB.Command
         Set adoCmd1.ActiveConnection = INIadoCon2
             adoCmd1.CommandText = "PLP_ERP_COG_HTKJ"
             adoCmd1.Parameters.Refresh
             adoCmd1.CommandType = adCmdStoredProc
        
          Set adoPrm11 = New ADODB.Parameter   '����Ϊ������
          adoPrm11.type = adChar
          adoPrm11.Size = 20
          adoPrm11.Direction = adParamInput
          adoPrm11.Value = tempBillNo
          adoCmd1.Parameters.Append adoprm1
          adoCmd1.Execute



Cnn.Execute ("CCS_COG_SENDREPORT") '��ʼ��TRC��������
MsgBox "�ܱ������������!", vbInformation, "������ʾ"
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
           " Where a.���ݱ�� = b.���ݱ�� and a.���ݱ��='" + UCase(Trim(TxtBillNo.text)) + "' "

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

Dim KK As String

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
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.text)) + "' order by 1  "
          
          
           
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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'SX') "
    
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

Dim KK As String

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
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.text)) + "' order by 1  "
          
          
           
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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'HD') "
    
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


Dim KK As String

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
           " Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.text)) + "'  order by 4 "

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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
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


Dim KK As String

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
             " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.text)) + "'   order by 4 "
 
           
           

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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
    
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
'��ѯ�����������

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim KK As String

    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_COG_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����

    strDatas = "NO,Sub Name,Ship To,Vacuum Bag ID,Vacuum Bag Qty,Box ID,Customer Device,GC Version," & _
                "Grade,Bond Pro,Ship Out Date,PO NO,Carton NO,WO,Invoice NO,Remark" & vbCrLf
       
               
    
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
             " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.text)) + "'   order by 4 "
 
           
           

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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
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


Dim KK As String

'PL_HTKS_WLT_20151111001.csv
    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_WLT_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "NO,Sub Name,Ship To,Fab Device,Customer Device,GC Version,PO NO,Invoice NO,Ship Out Date,FAB Lot ID," & _
               "Wafer ID,Gross Dies,Sampling Qty,Pass Dies1,Pass Dies2, Pass Dies3,NG Die,Yield,Pack Lot ID,Wafer Mark,Grade,Carton NO,WO,Remark" & vbCrLf
    '��ϸ����
'    strSql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & _
'           " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & _
'           " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & _
'           " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & _
'           " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & _
'           " Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "
           
           
           
'     strSql = "  select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,left(rtrim(ltrim(GC_Version)),2)+erpdata.dbo.GET_TSV_DOUBLECODE(rtrim(FAB_Lot_ID)+rtrim(Wafer_ID)) as gcversion, cast([NO] as int), " & _
'    " [Sub_Name],'GCSH' as [Ship_To],[Fab_Device],[Customer_Device],[GC_Version], " & _
'    " [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [PACK_Out_Date],[FAB_Lot_ID],CASE WHEN CHARINDEX('-',Wafer_ID)>0 THEN RIGHT(Wafer_ID,2) WHEN CHARINDEX('+',Wafer_ID)>0 THEN RIGHT(Wafer_ID,2) ELSE Wafer_ID END AS Wafer_ID,[Gross_Dies]," & _
'    " erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as [Sampling Qty] ,[Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as  [Pass_Dies],[NG_Die],[Yield] ,[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark]" & _
'    " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGCWlt.Text)) + "'  order by 4 "
               
               
               
    
    'strsql = " select rtrim(ltrim(FAB_Lot_ID)) + rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,a.GC_Version,left(rtrim(ltrim(GC_Version)), 2) +erpdata.dbo.GET_TSV_DOUBLECODE(rtrim(FAB_Lot_ID) + rtrim(Wafer_ID)) as gcversion,cast( NO  as int)," & _
    '         " Sub_Name ,'GCSH' as  Ship_To ,Fab_Device ,Customer_Device ,GC_Version ," & _
    '         " PO_NO ,Invoice_NO ,replace(PACK_Out_Date , '/', '-') as  PACK_Out_Date ,FAB_Lot_ID ,CASE WHEN CHARINDEX('-', Wafer_ID) > 0 THEN RIGHT(Wafer_ID, 2) WHEN CHARINDEX('+', Wafer_ID) > 0 THEN RIGHT(Wafer_ID, 2) Else Wafer_id END AS Wafer_ID,Gross_Dies ," & _
    '         " Pass_Dies,NG_Die ,Yield ,PACK_Lot_ID ,Wafer_Mark ,'A' as Grade,CartonNO ,WO ,Remark FROM  erpdata..GR_GC_DetailHistory  a Where a.���ݱ�� = 'F1708030012' order by 4"
               

'       strSql = "  select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,left(rtrim(ltrim(GC_Version)),2)+erpdata.dbo.GET_TSV_DOUBLECODE(rtrim(FAB_Lot_ID)+rtrim(Wafer_ID)) as gcversion, cast([NO] as int), " & _
'    " [Sub_Name],'GCSH' as [Ship_To],[Fab_Device],[Customer_Device],[GC_Version], " & _
'    " [PO_NO] ,[Invoice_NO],replace([PACK_Out_Date],'/','-') as [PACK_Out_Date],[FAB_Lot_ID],CASE WHEN CHARINDEX('-',Wafer_ID)>0 THEN RIGHT(Wafer_ID,2) WHEN CHARINDEX('+',Wafer_ID)>0 THEN RIGHT(Wafer_ID,2) ELSE Wafer_ID END AS Wafer_ID,[Gross_Dies], erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) as [Sampling Qty], " & _
'    " [Pass_Dies]-erpdata.dbo.Get_TSV_GC_WaferGDieBin3(rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID))) - erpdata.dbo.Get_TSV_GC_WaferGDieBin2(rtrim(ltrim(FAB_Lot_ID)) + rtrim(ltrim(Wafer_ID))) as  [Pass_Dies1],erpdata.dbo.Get_TSV_GC_WaferGDieBin2(rtrim(ltrim(FAB_Lot_ID)) + rtrim(ltrim(Wafer_ID))) as [Pass_Dies2],'' as [Pass_Dies3], [NG_Die],[Yield],[PACK_Lot_ID],[Wafer_Mark],'A' as Grade,[CartonNO] ,[WO],[Remark]" & _
'    " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGCWlt.Text)) + "'  order by 4 "
'
'

'
    strSql = "SELECT row_number() over(order by t.lot_id,t.waferid) AS rounum ,t.sub_name,t.SHIP_SITE,t.FAB_CONV_ID,t.cust_device,t.gcversion,t.PO_NUM,t.invoice,t.create_date,t.lot_id,t.waferid,t.gross_die,ISNULL(t.BIN3,L.NDPW) as Sampling_Qty " & _
",ISNULL(ISNULL(t.BIN1,t.A),K.NDPW) as Pass_Dies1,ISNULL(T.BIN2,m.NDPW) as Pass_Dies2,'' AS Pass_Dies3,ISNULL(ISNULL(T.E,n.NDPW),0) as NG_Die,CONVERT(VARCHAR(10),CONVERT(decimal(18,2), (t.gross_die - ISNULL(ISNULL(T.E,n.NDPW),0))*1.0/(t.gross_die )*100)) + '%' AS Yield   " & _
",ISNULL(t.sfc,k.FIRSTNAME) as Pack_Lot_ID,t.PRODUCTID,'A' as Grade,rtrim(t.���) as ���,t.�󹤵� FROM ( SELECT  'HTKS' AS sub_name,d.SHIP_SITE,RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID " & _
",a.cust_device,a.gcversion,d.PO_NUM,'' AS invoice, a.create_date,rtrim(a.lot_id) as lot_id,SUBSTRING(REPLACE(b.���̿����,'+',''),LEN(a.lot_id)+1,2) as waferid,c.FAILBINCOUNT+c.PASSBINCOUNT AS gross_die " & _
",e.GRADES,e.QTY,c.PRODUCTID,'A' as Grade,rtrim(ay.���) as ���,b.�󹤵�,a.qbox,b.���̿����,SUBSTRING( e.SFC,12,8) AS SFC " & _
"FROM erptemp..tblshipreport_new a INNER JOIN erpdata..tblStockNumTree ax  ON ax.��� =a.qbox  INNER JOIN erpdata..tblStockNumTree ay  ON ay.��� = ax.�ϼ���� " & _
"INNER JOIN  erpdata..tblStocksqfhsub b  ON b.���ݱ�� = a.ship_order AND b.��� = a.qbox AND b.������ = a.lot_id " & _
"INNER JOIN  ERPBASE..tblmappingData c ON  c.SUBSTRATEID = b.���̿���� AND c.LOTID = b.������ INNER JOIN  erpbase..tblCustomerOI d ON  CONVERT(VARCHAR(20), d.ID) = c.FILENAME AND d.SOURCE_BATCH_ID = c.LOTID  " & _
"left JOIN  erptemp..WAFER_BIN_LIST e ON e.WAFER_ID = b.���̿���� inner JOIN erpdata..tblErpInStockRelation ee ON ee.SFC_ID = e.SFC  AND CHARINDEX(e.WAFER_ID,ee.WAFER_ID) <> 0    " & _
"WHERE a.ship_order = '" & UCase(Trim(TxtBillNoGCWlt.text)) & "' ) AS p PIVOT(sum(qty) FOR grades IN(A,BIN1,BIN2,BIN3, E)) AS T " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV k ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.���̿���� AND k.CONTAINERNAME LIKE '%-A' " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox AND L.WAFERSCRIBENUMBER = t.���̿���� AND L.CONTAINERNAME LIKE '%-A-01' " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV m ON m.QBOXNUMBER = t.qbox AND m.WAFERSCRIBENUMBER = t.���̿���� AND m.CONTAINERNAME LIKE '%-A-02' " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV n ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.���̿���� AND n.CONTAINERNAME LIKE '%-E' "



'    strSql = "SELECT t.FAB_CONV_ID,t.cust_device,t.gcversion,t.PO_NUM,t.invoice,t.create_date,t.lot_id,t.waferid,t.gross_die,ISNULL(t.BIN3,L.NDPW) as Sampling_Qty " & _
'",ISNULL(ISNULL(t.BIN1,t.A),K.NDPW) as Pass_Dies1,ISNULL(T.BIN2,m.NDPW) as Pass_Dies2,'' AS Pass_Dies3,ISNULL(ISNULL(T.E,n.NDPW),0) as NG_Die,CONVERT(VARCHAR(10),CONVERT(decimal(18,2), (t.gross_die - ISNULL(ISNULL(T.E,n.NDPW),0))*1.0/(t.gross_die )*100)) + '%' AS Yield   " & _
'",ISNULL(t.sfc,k.FIRSTNAME) as Pack_Lot_ID,t.PRODUCTID,'A' as Grade,rtrim(t.���) as ���,t.�󹤵� FROM ( SELECT  'HTKS' AS sub_name,d.SHIP_SITE,RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID " & _
'",a.cust_device,a.gcversion,d.PO_NUM,'' AS invoice, a.create_date,rtrim(a.lot_id) as lot_id,SUBSTRING(REPLACE(b.���̿����,'+',''),LEN(a.lot_id)+1,2) as waferid,c.FAILBINCOUNT+c.PASSBINCOUNT AS gross_die " & _
'",e.GRADES,e.QTY,c.PRODUCTID,'A' as Grade,rtrim(ay.���) as ���,b.�󹤵�,a.qbox,b.���̿����,SUBSTRING( e.SFC,12,8) AS SFC " & _
'"FROM erptemp..tblshipreport_new a INNER JOIN erpdata..tblStockNumTree ax  ON ax.��� =a.qbox  INNER JOIN erpdata..tblStockNumTree ay  ON ay.��� = ax.�ϼ���� " & _
'"INNER JOIN  erpdata..tblStocksqfhsub b  ON b.���ݱ�� = a.ship_order AND b.��� = a.qbox AND b.������ = a.lot_id " & _
'"INNER JOIN  ERPBASE..tblmappingData c ON  c.SUBSTRATEID = b.���̿���� AND c.LOTID = b.������ INNER JOIN  erpbase..tblCustomerOI d ON  CONVERT(VARCHAR(20), d.ID) = c.FILENAME AND d.SOURCE_BATCH_ID = c.LOTID  " & _
'"left JOIN  erptemp..WAFER_BIN_LIST e ON e.WAFER_ID = b.���̿���� inner JOIN erpdata..tblErpInStockRelation ee ON ee.SFC_ID = e.SFC  AND CHARINDEX(e.WAFER_ID,ee.WAFER_ID) <> 0    " & _
'"WHERE a.ship_order = '" & UCase(Trim(TxtBillNoGCWlt.Text)) & "' ) AS p PIVOT(sum(qty) FOR grades IN(A,BIN1,BIN2,BIN3, E)) AS T " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV k ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.���̿���� AND k.CONTAINERNAME LIKE '%-A' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox AND L.WAFERSCRIBENUMBER = t.���̿���� AND L.CONTAINERNAME LIKE '%-A-01' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV m ON m.QBOXNUMBER = t.qbox AND m.WAFERSCRIBENUMBER = t.���̿���� AND m.CONTAINERNAME LIKE '%-A-02' " & _
'"LEFT JOIN erpdata..TblQBOXNUMBER_TSV n ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.���̿���� AND n.CONTAINERNAME LIKE '%-E' "
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
'                MsgBox waferidMain & " ��Ƭ�������볤�Ȳ�����3����ȷ�Ϻú���ܵ�����", vbInformation, "������ʾ"
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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
    
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
'��ѯ�����������

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim KK As String

'PL_HTKS_WLT_20151111001.csv
    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_COG_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "NO,Sub Name,Ship To,Vacuum Bag ID,Vacuum Bag Qty,Box ID,Customer Device," & _
               "GC Version,Grade,Bond Pro,Ship Out Date,PO NO,Carton NO,WO,Invoice NO,Remark" & vbCrLf
               
'strsql = " select ROW_NUMBER() OVER(ORDER BY X.Qboxnumber ,X.lvNo) as No,Sub_Name,ShipTo,lvNo,VacuumBagQty,Qboxnumber,Customer_Device,GCVersion, " & _
'" Grade,BondPro,ShipOutDate,PO_NO,CartonNO,WOnumber,InvoiceNO,Remark from ( " & _
'" select  distinct  Sub_Name,'GC_SH' as ShipTo,e.[LVBARCODEQBOX] as lvNo,CASE WHEN dbo.GET_TSV_COG_LOTQTY(A.FAB_Lot_ID,E.LVBARCODEQBOX)=0 THEN 1500 ELSE dbo.GET_TSV_COG_LOTQTY(A.FAB_Lot_ID,E.LVBARCODEQBOX)END as VacuumBagQty,c.��� as Qboxnumber, " & _
'" Customer_Device,GC_Version+'D' as GCVersion,'A' as Grade,'SH' as BondPro,REPLACE(PACK_Out_Date,'/','-') as ShipOutDate,PO_NO," & _
'" '1' as CartonNO,d.MTRL_NUM as WOnumber,'' as InvoiceNO,''as Remark " & _
'" from  [erpdata].[dbo].[GR_GC_DetailHistory] a ,[erpdata].[dbo].[tblPackTreeInf] b  , [erpdata].[dbo].[tblPackTreeInf] c,[ERPBASE].[dbo].[tblCustomerOI] d,[erpdata].[dbo].[GR_COG_LV_Data] e   " & _
'" Where a.���ݱ��='" + UCase(Trim(TxtCogNo.Text)) + "'" & _
' "   and b.���=a.CartonNO and c.�ϼ����=b.��� and d.SOURCE_BATCH_ID=a.FAB_Lot_ID and d.PO_NUM=a.PO_NO and e.[QBOXNUMBER]=c.��� ) X"
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
    strDatas = strDatas + strRowData '��������
    'д���ļ�
    LogFile.WriteLine (strDatas)
    
    LogFile.Close
    Set LogFile = Nothing
    
    
    '���ʼ�
    Dim strRecipient    As String
    Dim strRecipientCC  As String
    
    strRecipient = "ks015918@ht-tech.com"
    strRecipientCC = "wanli.ma@htkjks.com"
        
    Call MailDetail("GC COG ��������", strRecipient, g_Path & "\" & "PL_HTKS_COG_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtCogNo.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
    
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
'��ѯ�����������

fileNo = GetGC_FileNoNew("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim KK As String

'PL_HTKS_WLT_20151111001.csv
    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "PL_HTKS_COG_TRC_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "BOXA_ID,BOXB_ID,CHIP_QTY,GRADE,BOX_TYPE,SERIAL_NUM" & vbCrLf
           
               
    strSql = " select TrayNo ,LotID  ,Qty  ,Grade ,BoxType ,SERIAL_NUM from  TSV_GR_COG_Tray_Data order by id "
            '   strSql = " select BOXA_ID ,BOXB_ID  ,CHIP_QTY  ,GRADE ,BOXTYPE ,SERIAL_NUM from  TSV_GR_COG_TRAY_DATA_1"
               
    ' tangwei: 20171010 ���¶Ե�

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
        
    Call MailDetail("GC COG ��������", strRecipient, g_Path & "\" & "PL_HTKS_COG_TRC_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtCogNo.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
    
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
'��ѯ�����������

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
        
    Call MailDetail("GC COG ��������", strRecipient, g_Path & "\" & "PLP_ERP_COG_HTKJ_" & Format(g_Date, "YYYYMMDD") & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
   Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtCogNo.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
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
'��ѯ�����������

fileNo = GetGC_FileNo("GC")
waferidMain = ""
waferPT = ""
waferVer = ""
waferVerResult = ""


Dim KK As String

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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
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


Dim KK As String

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

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.text)) + "',getdate(),'Y','Auto',getdate(),'GC') "
    
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
    xlSheet.name = "GrData"
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
tempBillNo = UCase(Trim(TxtBillNo.text))

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
    Set xlApp = Nothing  '"���ٱ��Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Sub



Private Sub Form_Load()

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
" FROM [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "'  order by a.CartonNO "
           
          
          
          
          
          
          
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
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by a.CartonNO  "
          
          
ElseIf custNameTemp = "HD" Then

sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Fab_Device] as [�汾],[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & _
          " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[Gross_Dies]as [NGDieQty],[NG_Die]as [ShipmentGoodDie] ,[Yield] ,[PACK_Out_Date]as[��������], " & _
          "  [Remark] as [��ע] " & _
          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by a.CartonNO  "
End If

  SqlServerExporToExcel (sqlTemp)

End Sub

Private Sub GCCmdSend_Click()



'��������
Dim tempBillNo As String
Dim custNameTemp As String

tempBillNo = UCase(Trim(TxtBillNoGC.text))
custNameTemp = UCase(Trim(comBo2.text))


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
txtPackage.SetFocus
End If

End Sub



