VERSION 5.00
Begin VB.Form FrmAAGr 
   Caption         =   "AA��SX �ͻ�������Ϣ"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13440
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
   ScaleHeight     =   7830
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "�����ͻ���������"
      Height          =   3975
      Left            =   360
      TabIndex        =   14
      Top             =   3600
      Width           =   12735
      Begin VB.CommandButton cmdSX_NEW 
         BackColor       =   &H00FF80FF&
         Caption         =   "SX_NEW��������"
         Height          =   480
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmAAGr.frx":0000
         Left            =   1320
         List            =   "FrmAAGr.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton GCCmdSend 
         Caption         =   "���ͱ���"
         Height          =   480
         Left            =   7320
         TabIndex        =   18
         Top             =   1320
         Width           =   990
      End
      Begin VB.CommandButton GCCmdOut 
         Caption         =   "��������"
         Height          =   480
         Left            =   5280
         TabIndex        =   17
         Top             =   1320
         Width           =   990
      End
      Begin VB.TextBox TxtBillNoGC 
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ���"
         Height          =   195
         Left            =   600
         TabIndex        =   20
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݱ�ţ�"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AA GR"
      Height          =   3375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.TextBox txtdelNote 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtawb 
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox TxtPackage 
         Height          =   375
         Left            =   6360
         TabIndex        =   6
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CommandButton CmdSaver 
         Caption         =   "������Ϣ"
         Height          =   480
         Left            =   2880
         TabIndex        =   5
         Top             =   2400
         Width           =   990
      End
      Begin VB.CommandButton CmdSend 
         Caption         =   "����GR"
         Height          =   480
         Left            =   7200
         TabIndex        =   4
         Top             =   2400
         Width           =   990
      End
      Begin VB.TextBox TxtWeight 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox TxtBillNo 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton CmdOut 
         Caption         =   "����GR"
         Height          =   480
         Left            =   5160
         TabIndex        =   1
         Top             =   2400
         Width           =   990
      End
      Begin VB.Label LblKey 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del_Note��"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label LblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AWB��"
         Height          =   195
         Left            =   5640
         TabIndex        =   12
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weight(kgs)��"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Package��"
         Height          =   195
         Left            =   5400
         TabIndex        =   10
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݱ�ţ�"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmAAGr"
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

Dim reportRS  As New ADODB.Recordset

Public g_Date As String

Private Sub CmdAdd_Click()

    '����
    Dim tempKey    As String

    Dim tempValue  As String

    Dim getValue   As String

    Dim otherValue As String

    Dim sqlTemp    As String

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

Private Sub CmdOut_Click()

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

    Dim sqlTemp As String

    sqlTemp = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & " Where a.���ݱ�� = b.���ݱ�� and a.���ݱ��='" + tempBillNo + "' "

    SqlServerExporToExcel (sqlTemp)

End Sub

Private Sub CmdSaver_Click()
    '���浽SqlServer��

    Dim tempBillNo  As String

    Dim tempdelNote As String

    Dim tempAwb     As String

    Dim tempWeight  As Single

    Dim tempPackage As Integer

    Dim cmdStrSql   As String

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

    If tempBillNo = "" Or tempdelNote = "" Or tempAwb = "" Or Trim(txtWeight.Text) = "" Or Trim(TxtPackage.Text) = "" Then
        MsgBox "��������������!", vbInformation, "������ʾ"

        Exit Sub

    End If

    tempWeight = CSng(Trim(txtWeight.Text))
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

    cmdStrSql = " insert into [erpdata].[dbo].[GRdetailSetUp](���ݱ��,Del_Note,AWB,[Weight],Package) values('" & tempBillNo & "'," & " '" & tempdelNote & "','" & tempAwb & "'," & tempWeight & "," & tempPackage & " )"

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

Private Sub cmdSX_NEW_Click()

    Dim tempBillNo   As String

    Dim custNameTemp As String

    tempBillNo = UCase(Trim(TxtBillNoGC.Text))
    custNameTemp = UCase(Trim(comBo2.Text))

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
           
        sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [Sub Name],[Ship_To]as [Ship To] ,[Fab_Device]as [Fab Device] ,[Customer_Device] as [Customer Device],[PO_NO] as [PO NO]," & " [WO],[GC_Version]as [GC Version],[Invoice_NO]as [Invoice NO] ,[PACK_Out_Date]as[PACK-Out Date],[PACK_Lot_ID]as[PACK Lot ID],[FAB_Lot_ID]as[FAB Lot ID] ," & " [Wafer_ID]as [Wafer ID],[Wafer_Mark]as [Wafer Mark],[Gross_Dies]as [Gross Dies],[Pass_Dies]as [Pass Dies],[NG_Die]as [NG Die] ,[Yield] ," & " [Remark] , [System_CartonNO]as [System CartonNO], [PACK_Device]as [PACK Device], [CartonNO]as [CartonNO], [MaskType] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
                 
    ElseIf custNameTemp = "SX" Or custNameTemp = "HJ" Or custNameTemp = "TJ003" Or custNameTemp = "JS140" Then

        sqlTemp = "select ROW_NUMBER() over(order by c.SUBSTRATEID) as 'NO','HTKS' as '������' ,d.SHIP_SITE as '�ͻ�',d.MPN_DESC as '��Ʒ����',d.PO_NUM as '�ͻ�����' ,d.SOURCE_BATCH_ID as '�ͻ�Lot',SUBSTRING(replace(c.SUBSTRATEID,'+',''),LEN(replace(c.SUBSTRATEID,'+',''))-1,2) as 'WaferNo' " & _
           ",e.BIN1 as GoodDieQty,e.BIN2,e.BIN3,e.BIN4,e.BIN5,e.BIN6,e.BIN7,e.BIN8,e.BIN9,e.BIN10,e.BIN11,e.BIN12,e.BIN13,e.BIN14,e.BIN15,e.BIN16,e.BIN17,e.BIN18,[erpdata].[dbo].[Get_TSV_WaferNGDie](c.SUBSTRATEID,b.���) as 'BadDieQty', convert(varchar(100),Convert(decimal(18,2),[erpdata].[dbo].[Get_TSV_WaferGDie](c.SUBSTRATEID,b.���)/(c.PASSBINCOUNT + c.FAILBINCOUNT)*100))+ '%' as 'Yield' " & _
           ",CONVERT(varchar(100), a.��������,23 ) as '��������',c.PRODUCTID as 'LaserMark',g.��� from erpdata..tblStockMove a,erpdata..tblStockmovesub b,ERPBASE..tblmappingData c,ERPBASE..tblCustomerOI d,(select * from  OPENQUERY(ORACLEDB, 'SELECT * from SX_BINCODE_LIST' )) e,erpdata..tblstocknumtree f ,erpdata..tblstocknumtree g where a.�ͻ����� = 'SX' and a.���ݱ�� = '" + tempBillNo + "' and b.���ݱ�� = a.���ݱ�� and b.������� = a.��� and c.SUBSTRATEID = b.���̿���� and c.LOTID = d.SOURCE_BATCH_ID " & _
           "and CONVERT(varchar(100), d.ID) = c.FILENAME and e.waferid = b.���̿���� and f.��� = b.��� and g.��� = f.�ϼ����"

        'css add 20160707
    ElseIf custNameTemp = "56" Then

        sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & " [Wafer_Mark]as [Laser Mark],CartonNO as [���], [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
          
    ElseIf custNameTemp = "BD" Then

        sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & " [Wafer_Mark]as [Laser Mark],CartonNO as [���], [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
          
    ElseIf custNameTemp = "HD" Then

        sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Fab_Device] as [�汾],[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[Gross_Dies]as [NGDieQty],[NG_Die]as [ShipmentGoodDie] ,[Yield] ,[PACK_Out_Date]as[��������], " & "  [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
                    
    ElseIf custNameTemp = "45" Then

        sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Fab_Device] as [�汾],[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " rtrim(ltrim([FAB_Lot_ID]))as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & "  [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
          
    End If

    SqlServerExporToExcel (sqlTemp)
End Sub

Private Sub Combo2_Change()
    TxtBillNoGC.SetFocus

End Sub

Private Sub Combo2_Click()
    TxtBillNoGC.SetFocus
End Sub

Private Sub Form_Activate()
    TxtBillNo.SetFocus
    g_Date = Format(Now, "YYYY-MM-DD hh:mm:ss")
End Sub

Private Sub SaveFileSendTest()

    Dim FSO        As New FileSystemObject

    Dim LogFile    As TextStream

    Dim strDatas   As String

    Dim strRowData As String

    Dim strColData As String

    Dim strsql     As String

    Dim i          As Integer, J           As Integer

    Dim rs         As New ADODB.Recordset

    On Error GoTo ErrHandler

    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path_GR & "\" & "QT_FG_CSP_" & Format(g_Date, "YYYYMMDD") & "_" & Format(g_Date, "hhmmss") & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "PO_num,PO_Item,Previous_Batch_ID,Previous_Mtrl_Num,Batch_ID,mtrl_num,mtrl_desc,Mtrl_num_Mtrlgrp,Output_Qty,Consumed_Qty,Reject_Qty,Current_Wafer_Qty" & ",Film_Frame_Qty,Optical_Quality,Country_of_Assembly,Offshore_ASM_Company,Asm_Containment_type,Date_code,asm_conv_id,asm_excr_id,assembly_facility,Country_of_Test,Offshore_TEST_Company" & ",Tst_Containment_type,Tst_Program_rev,Created_date,Created_time,Del_Note,AWB,weight(kgs),package" & vbCrLf
    '��ϸ����
    strsql = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & " Where a.���ݱ�� = b.���ݱ�� and a.���ݱ��='" + UCase(Trim(TxtBillNo.Text)) + "' "

    strRowData = ""

    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If

    rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then Exit Sub

    For i = 1 To rs.RecordCount
        strColData = ""

        For J = 0 To rs.Fields.Count - 1

            If J = 26 Then
                strColData = strColData + Format(g_Date, "hh:mm:ss") + ","
            Else
                strColData = strColData + Trim("" & rs.Fields(J).Value) + ","
            
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

    Dim FSO        As New FileSystemObject

    Dim LogFile    As TextStream

    Dim strDatas   As String

    Dim strRowData As String

    Dim strColData As String

    Dim strsql     As String

    Dim i          As Integer, J           As Integer

    Dim maxRow     As Integer

    Dim rs         As New ADODB.Recordset

    Dim fileNo     As String

    On Error GoTo ErrHandler

    '��ѯ�����������

    fileNo = GetGC_FileNo("SX")

    Dim KK As String

    '�����ļ�
    If comBo2.Text = "TJ003" Then
        Set LogFile = FSO.CreateTextFile(g_Path & "\" & "TJ003_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    ElseIf comBo2.Text = "JS140" Then
        Set LogFile = FSO.CreateTextFile(g_Path & "\" & "JS140_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    Else
        Set LogFile = FSO.CreateTextFile(g_Path & "\" & "SX_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    End If

    'д����
    strDatas = ""
    'ͷ����
    strDatas = "NO,������,�ͻ�,��Ʒ����,�ͻ�������,�ͻ�Lot,WaferNo,GoodDieQty,BadDieQty,Yield,��������,LaserMark,���,��ע" & vbCrLf
    '��ϸ����
  
    strsql = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & " [Wafer_Mark]as [Laser Mark],CartonNO as [���], [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
           
    strRowData = ""

    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If

    rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""

        For J = 0 To rs.Fields.Count - 1

            strColData = strColData + Trim("" & rs.Fields(J).Value) + ","
           
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
    Dim strRecipient   As String

    Dim strRecipientCC As String

    Dim filePre        As String

    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    filePre = "SX_HTKS_CSP_"
    
    If comBo2.Text = "JS140" Then
        filePre = "JS140_HTKS_CSP_"
    End If
    
    Call MailDetailSX("SX ��������", strRecipient, g_Path & "\" & filePre & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
    Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'SX') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
ErrHandler:
    Set FSO = Nothing
End Sub

'Private Sub SaveFileSendSX()
'Dim FSO         As New FileSystemObject
'Dim LogFile     As TextStream
'Dim strDatas    As String
'Dim strRowData  As String
'Dim strColData  As String
'Dim strSql      As String
'Dim i           As Integer, j           As Integer
'
'Dim maxRow As Integer
'
'Dim Rs          As New ADODB.Recordset
'
'Dim fileNo As String
'
'On Error GoTo ErrHandler
''��ѯ�����������
'
'fileNo = GetGC_FileNo("SX")
'
'Dim kk As String
'
'    '�����ļ�
'    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "SX_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
'    'д����
'    strDatas = ""
'    'ͷ����
'    strDatas = "NO,������,�ͻ�,��Ʒ����,�ͻ�������,�ͻ�Lot,WaferNo,GoodDieQty,BadDieQty,Yield,��������,LaserMark,���,��ע" & vbCrLf
'    '��ϸ����
'
'  strSql = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & _
'          " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & _
'          " [Wafer_Mark]as [Laser Mark],CartonNO as [���], [Remark] as [��ע] " & _
'          " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "

'
'
'
'    strRowData = ""
'    If Rs.State = adStateOpen Then Rs.Close
'    If INIadoCon.State <> adStateOpen Then
'        INIConnectSTART
'    End If
'    Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
'    If Rs.EOF Then Exit Sub
'
'    maxRow = Rs.RecordCount
'
'    For i = 1 To Rs.RecordCount
'        strColData = ""
'        For j = 0 To Rs.fields.Count - 1
'
'             strColData = strColData + Trim("" & Rs.fields(j).Value) + ","
'
'        Next
'
'        If i = maxRow Then
'          strRowData = strRowData + Left(strColData, Len(strColData) - 1)
'
'        Else
'
'        strRowData = strRowData + Left(strColData, Len(strColData) - 1) + vbCrLf
'
'        End If
'
'        Rs.MoveNext
'    Next
'    strDatas = strDatas + strRowData '��������
'    'д���ļ�
'    LogFile.WriteLine (strDatas)
'
'    LogFile.Close
'    Set LogFile = Nothing
'
'
'    '���ʼ�
'    Dim strRecipient    As String
'    Dim strRecipientCC  As String
'
'    strRecipient = "jiayun.zhang@qtechglobal.com"
'    strRecipientCC = "wanli.ma@qtechglobal.com"
'
'    Call MailDetailSX("SX ��������", strRecipient, g_Path & "\" & "SX_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
'
'    '�ѷ��ͼ�¼���浽DB��
'
'    Dim sqltemp2 As String
'
'    sqltemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'SX') "
'
'    Call AddSql2(sqltemp2)
'
'    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
'
'
'ErrHandler:
'    Set FSO = Nothing
'End Sub
Private Sub SaveFileSend56()

    Dim FSO        As New FileSystemObject

    Dim LogFile    As TextStream

    Dim strDatas   As String

    Dim strRowData As String

    Dim strColData As String

    Dim strsql     As String

    Dim i          As Integer, J           As Integer

    Dim maxRow     As Integer

    Dim rs         As New ADODB.Recordset

    Dim fileNo     As String

    On Error GoTo ErrHandler

    '��ѯ�����������

    fileNo = GetGC_FileNo("56")

    Dim KK As String

    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "56_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "NO,������,�ͻ�,��Ʒ����,�ͻ�������,�ͻ�Lot,WaferNo,GoodDieQty,BadDieQty,Yield,��������,LaserMark,���,��ע" & vbCrLf
    '��ϸ����
  
    strsql = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & " [Wafer_Mark]as [Laser Mark],CartonNO as [���], [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
           
    strRowData = ""

    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If

    rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""

        For J = 0 To rs.Fields.Count - 1

            strColData = strColData + Trim("" & rs.Fields(J).Value) + ","
           
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
    Dim strRecipient   As String

    Dim strRecipientCC As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    'Call MailDetailSX("56 ��������", strRecipient, g_Path & "\" & "56_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
    Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'56') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
ErrHandler:
    Set FSO = Nothing
End Sub

Private Sub SaveFileSendBD()

    Dim FSO        As New FileSystemObject

    Dim LogFile    As TextStream

    Dim strDatas   As String

    Dim strRowData As String

    Dim strColData As String

    Dim strsql     As String

    Dim i          As Integer, J           As Integer

    Dim maxRow     As Integer

    Dim rs         As New ADODB.Recordset

    Dim fileNo     As String

    On Error GoTo ErrHandler

    '��ѯ�����������

    fileNo = GetGC_FileNo("BD")

    Dim KK As String

    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "BD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "NO,������,�ͻ�,��Ʒ����,�ͻ�������,�ͻ�Lot,WaferNo,GoodDieQty,BadDieQty,Yield,��������,LaserMark,���,��ע" & vbCrLf
    '��ϸ����
  
    strsql = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & " [Wafer_Mark]as [Laser Mark],CartonNO as [���], [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
           
    strRowData = ""

    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If

    rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""

        For J = 0 To rs.Fields.Count - 1

            strColData = strColData + Trim("" & rs.Fields(J).Value) + ","
           
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
    Dim strRecipient   As String

    Dim strRecipientCC As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    Call MailDetailSX("BD ��������", strRecipient, g_Path & "\" & "BD_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
    Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'BD') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
ErrHandler:
    Set FSO = Nothing
End Sub

Private Sub SaveFileSendHD()

    Dim FSO        As New FileSystemObject

    Dim LogFile    As TextStream

    Dim strDatas   As String

    Dim strRowData As String

    Dim strColData As String

    Dim strsql     As String

    Dim i          As Integer, J           As Integer

    Dim maxRow     As Integer

    Dim rs         As New ADODB.Recordset

    Dim fileNo     As String

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
  
    strsql = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Fab_Device] as [�汾],[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[Gross_Dies]as [NGDieQty],[NG_Die]as [ShipmentGoodDie] ,[Yield] ,[PACK_Out_Date]as[��������], " & " [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
           
    strRowData = ""

    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If

    rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""

        For J = 0 To rs.Fields.Count - 1

            strColData = strColData + Trim("" & rs.Fields(J).Value) + ","
           
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
    Dim strRecipient   As String

    Dim strRecipientCC As String
    
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

    Dim FSO            As New FileSystemObject

    Dim LogFile        As TextStream

    Dim strDatas       As String

    Dim strRowData     As String

    Dim strColData     As String

    Dim strsql         As String

    Dim i              As Integer, J           As Integer

    Dim waferidMain    As String

    Dim waferPT        As String

    Dim waferVer       As String

    Dim waferVerResult As String

    Dim maxRow         As Integer

    Dim rs             As New ADODB.Recordset

    Dim fileNo         As String

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
    strDatas = "NO,Sub Name,Ship To,Fab Device,Customer Device,PO NO,WO,GC Version,Invoice NO,PACK-Out Date,PACK Lot ID,FAB Lot ID" & ",Wafer ID,Wafer Mark,Gross Dies,Pass Dies,NG Die,Yield,Remark,System CartonNO,PACK Device,CartonNO,MaskType" & vbCrLf
    '��ϸ����
    strsql = "select rtrim(ltrim(FAB_Lot_ID))+rtrim(ltrim(Wafer_ID)) as waferidMain,rtrim(ltrim(Customer_Device)) as device,rtrim(ltrim(GC_Version)) as gcversion, cast([NO] as int),[Sub_Name],[Ship_To],[Fab_Device],[Customer_Device],[PO_NO] " & " ,[WO],[GC_Version],[Invoice_NO],[PACK_Out_Date],[PACK_Lot_ID],[FAB_Lot_ID] " & " ,[Wafer_ID],[Wafer_Mark],[Gross_Dies],[Pass_Dies],[NG_Die],[Yield] " & " ,[Remark],[System_CartonNO],[PACK_Device],[CartonNO],[MaskType] " & " FROM [erpdata].[dbo].[GR_GC_DetailHistory] a  " & " Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "'  order by 4 "

    strRowData = ""

    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If

    rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""
        
        waferidMain = Trim("" & rs.Fields(0).Value) & "-A"
            
        waferPT = Trim("" & rs.Fields(1).Value)
            
        waferVer = Trim("" & rs.Fields(2).Value)
            
        waferVerResult = GetGCOutRpt_Ver(waferidMain, waferPT, waferVer)
        
        For J = 3 To rs.Fields.Count - 1
             
            If J = 10 Then
             
                strColData = strColData + waferVerResult + ","
             
            Else
             
                strColData = strColData + Trim("" & rs.Fields(J).Value) + ","
             
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
    Dim strRecipient   As String

    Dim strRecipientCC As String
    
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

Private Sub SaveFileSend()
    'Excel����

    Dim strsql          As String

    Dim i               As Integer, J           As Integer

    Dim rs              As New ADODB.Recordset

    Dim RsD             As New ADODB.Recordset

    Dim xlApp           As New Excel.Application

    Dim xlBook          As Excel.Workbook

    Dim xlSheet         As Excel.Worksheet

    Dim currentSheetRow As Long

    Dim txtHeaderTemp   As String

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
    
    txtHeaderTemp = "PO_num,PO_Item,Previous_Batch_ID,Previous_Mtrl_Num,Batch_ID,mtrl_num,mtrl_desc,Mtrl_num_Mtrlgrp,Output_Qty,Consumed_Qty,Reject_Qty,Current_Wafer_Qty" & " Film_Frame_Qty,Optical_Quality,Country_of_Assembly,Asm_Containment_type,Date_code,asm_conv_id,asm_excr_id,assembly_facility,Country_of_Test,Offshore_TEST_Company" & " Tst_Containment_type,Tst_Program_rev,Created_date,Created_time,Del_Note,AWB,weight(kgs),package" & vbCrLf
    xlSheet.Cells(1, 1) = txtHeaderTemp
    
    Dim tempBillNo As String

    tempBillNo = UCase(Trim(TxtBillNo.Text))

    Dim sqlTemp As String

    strsql = "SELECT [PO_num] ,[PO_Item] ,[Previous_Batch_ID] ,[Previous_Mtrl_Num],[Batch_ID] ,[Mtrl_num] ,[Mtrl_desc] ,[Mtrl_num_Mtrlgrp] " & " ,[Output_Qty] ,[Consumed_Qty] ,[Reject_Qty] ,[Current_Wafer_Qty] ,[Film_Frame_Qty] ,[Optical_Quality] ,[Country_of_Assembly] " & " ,[Offshore_ASM_Company] ,[Asm_Containment_type] ,[Date_code] ,[asm_conv_id] ,[asm_excr_id] ,[assembly_facility] " & " ,[Country_of_Test],[Offshore_TEST_Company] ,[Tst_Containment_type] ,[Tst_Program_rev] ,[Created_date] ,[Created_time],b.Del_Note,b.AWB,b.Weight,b.Package " & " FROM [erpdata].[dbo].[GRdetailHistory] a,[erpdata].[dbo].[GRdetailSetUp] b " & " Where a.���ݱ�� = b.���ݱ�� and a.���ݱ��='" + tempBillNo + "' "

    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If

    rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then Exit Sub
    '     xlSheet.Range("a2:K" & Rs.RecordCount + 1).NumberFormatLocal = "@"
    currentSheetRow = rs.RecordCount + 1

    For i = 2 To rs.RecordCount + 1
        For J = 0 To rs.Fields.Count - 1
            xlSheet.Cells(i, J + 1) = Trim("" & rs.Fields(J).Value)
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
    comBo2.AddItem ("HJ")
    comBo2.AddItem ("TJ003")
    comBo2.AddItem ("HD")
    comBo2.AddItem ("BD")
    comBo2.AddItem ("45")
    comBo2.AddItem ("56")
    comBo2.AddItem ("JS140")

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

    Dim tempBillNo   As String

    Dim custNameTemp As String

    tempBillNo = UCase(Trim(TxtBillNoGC.Text))
    custNameTemp = UCase(Trim(comBo2.Text))

    If tempBillNo = "" Or custNameTemp = "" Then
        MsgBox "��ѡ��ͻ����룬���뵥�ݱ��!", vbInformation, "������ʾ"

        Exit Sub

    End If

    If Left(tempBillNo, 3) = "FDP" Then
    
        SqlServerExporToExcel ("SELECT row_number() OVER(ORDER BY X.[�ͻ�Lot], X.[WaferNo]) AS [ NO ],X.* FROM (SELECT DISTINCT 'HTKS' as [������], 'super pix' as [�ͻ�], B.FAB_CONV_ID AS [��Ʒ����], " & " B.PO_NUM AS [�ͻ�������],rtrim(A.������) as [�ͻ�Lot],right(rtrim(ltrim(a.���̿����)), 2) as [WaferNo],a.�ϸ��� as [GoodDieQty],0 as [BadDieQty],'100%' as [Yield],convert(varchar(100), getdate(), 111) AS [��������], " & "d.productid as [LaserMark],c.QBOXNUMBER as [���],'' as [��ע] FROM erpdata.dbo.tblStockdbsub a,[ERPBASE].[dbo].[tblCustomerOI] b, erpdata.dbo.TblQBOXNUMBER_TSV c,[ERPBASE].[dbo].[tblmappingData] d, " & "[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.������� = '" & tempBillNo & "' and b.SOURCE_BATCH_ID = a.������ and d.filename = cast(b.ID as nvarchar)and c.WAFERSCRIBENUMBER = a.���̿���� " & " and c.WAFERNUMBER = a.������ and d.SUBSTRATEID = a.���̿���� and f.WAFERID = a.���̿����) X")

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
           
        sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [Sub Name],[Ship_To]as [Ship To] ,[Fab_Device]as [Fab Device] ,[Customer_Device] as [Customer Device],[PO_NO] as [PO NO]," & " [WO],[GC_Version]as [GC Version],[Invoice_NO]as [Invoice NO] ,[PACK_Out_Date]as[PACK-Out Date],[PACK_Lot_ID]as[PACK Lot ID],[FAB_Lot_ID]as[FAB Lot ID] ," & " [Wafer_ID]as [Wafer ID],[Wafer_Mark]as [Wafer Mark],[Gross_Dies]as [Gross Dies],[Pass_Dies]as [Pass Dies],[NG_Die]as [NG Die] ,[Yield] ," & " [Remark] , [System_CartonNO]as [System CartonNO], [PACK_Device]as [PACK Device], [CartonNO]as [CartonNO], [MaskType] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
                 
    ElseIf custNameTemp = "SX" Or custNameTemp = "HJ" Or custNameTemp = "TJ003" Or custNameTemp = "JS140" Then

        sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & " [Wafer_Mark]as [Laser Mark],CartonNO as [���], [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
          
        'css add 20160707
    ElseIf custNameTemp = "56" Then

        sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & " [Wafer_Mark]as [Laser Mark],CartonNO as [���], [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
          
    ElseIf custNameTemp = "BD" Then

        sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & " [Wafer_Mark]as [Laser Mark],CartonNO as [���], [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
          
    ElseIf custNameTemp = "HD" Then

        sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Fab_Device] as [�汾],[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " [FAB_Lot_ID]as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[Gross_Dies]as [NGDieQty],[NG_Die]as [ShipmentGoodDie] ,[Yield] ,[PACK_Out_Date]as[��������], " & "  [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
                    
    ElseIf custNameTemp = "45" Then

        sqlTemp = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Fab_Device] as [�汾],[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " rtrim(ltrim([FAB_Lot_ID]))as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & "  [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + tempBillNo + "' order by 1  "
          
    End If

    SqlServerExporToExcel (sqlTemp)

End Sub

Private Sub GCCmdSend_Click()

    '����
    Dim tempBillNo   As String

    Dim custNameTemp As String

    tempBillNo = UCase(Trim(TxtBillNoGC.Text))
    custNameTemp = UCase(Trim(comBo2.Text))

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

        SaveFileSendGC

    ElseIf custNameTemp = "SX" Or custNameTemp = "HJ" Or custNameTemp = "TJ003" Or custNameTemp = "JS140" Then
        SaveFileSendSX

    ElseIf custNameTemp = "56" Then
        SaveFileSend56

    ElseIf custNameTemp = "BD" Then
        SaveFileSendBD

    ElseIf custNameTemp = "HD" Then
        SaveFileSendHD

    ElseIf custNameTemp = "45" Then
        SaveFileSend45

    End If
    
End Sub

Private Sub SaveFileSend45()

    Dim FSO        As New FileSystemObject

    Dim LogFile    As TextStream

    Dim strDatas   As String

    Dim strRowData As String

    Dim strColData As String

    Dim strsql     As String

    Dim i          As Integer, J           As Integer

    Dim maxRow     As Integer

    Dim rs         As New ADODB.Recordset

    Dim fileNo     As String

    On Error GoTo ErrHandler

    '��ѯ�����������

    fileNo = GetGC_FileNo("45")

    Dim KK As String

    '�����ļ�
    Set LogFile = FSO.CreateTextFile(g_Path & "\" & "45_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv")
    'д����
    strDatas = ""
    'ͷ����
    strDatas = "NO,������,�ͻ�,�汾,��Ʒ����,�ͻ�������,�ͻ�Lot,WaferNo,GoodDieQty,NGDieQty,Yield,��������,��ע" & vbCrLf
    '��ϸ����
  
    strsql = " select cast([NO] as int) as NO,[Sub_Name] as [������],[Ship_To]as [�ͻ�] ,[Fab_Device] as [�汾],[Customer_Device] as [��Ʒ����],[PO_NO] as [�ͻ�������]," & " rtrim(ltrim([FAB_Lot_ID])) as[�ͻ�Lot] ,[Wafer_ID]as [WaferNo],[Pass_Dies]as [GoodDieQty],[NG_Die]as [BadDieQty] ,[Yield] ,[PACK_Out_Date]as[��������], " & " [Remark] as [��ע] " & " From [erpdata].[dbo].[GR_GC_DetailHistory] a Where a.���ݱ��='" + UCase(Trim(TxtBillNoGC.Text)) + "' order by 1  "
           
    strRowData = ""

    If rs.State = adStateOpen Then rs.Close
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART
    End If

    rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then Exit Sub
    
    maxRow = rs.RecordCount
    
    For i = 1 To rs.RecordCount
        strColData = ""

        For J = 0 To rs.Fields.Count - 1

            strColData = strColData + Trim("" & rs.Fields(J).Value) + ","
           
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
    Dim strRecipient   As String

    Dim strRecipientCC As String
    
    strRecipient = "jiayun.zhang@qtechglobal.com"
    strRecipientCC = "wanli.ma@qtechglobal.com"
        
    'Call MailDetail45("45 ��������", strRecipient, g_Path & "\" & "45_HTKS_CSP_" & Format(g_Date, "YYYYMMDD") & "-" & fileNo & ".csv", strRecipientCC)
    
    '�ѷ��ͼ�¼���浽DB��
    
    Dim sqlTemp2 As String

    sqlTemp2 = " insert into [erpdata].[dbo].[GR_GC_SendHistory](���ݱ��,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(TxtBillNoGC.Text)) + "',getdate(),'Y','Auto',getdate(),'45') "
    
    Call AddSql2(sqlTemp2)
    
    MsgBox "���ͳɹ���", vbInformation, "������ʾ"
    
ErrHandler:
    Set FSO = Nothing
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
