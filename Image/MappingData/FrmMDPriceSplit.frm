VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmMDPriceSplit 
   Caption         =   "�г��������۸����PO"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
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
   ScaleHeight     =   9885
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   360
      Left            =   3480
      TabIndex        =   7
      Top             =   2640
      Width           =   990
   End
   Begin VB.TextBox Txtno 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox TxtPO 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "��ѯ"
      Height          =   360
      Left            =   5760
      TabIndex        =   1
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H000000FF&
      Caption         =   "����"
      Height          =   360
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   990
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   1215
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   17055
      _Version        =   196613
      _ExtentX        =   30083
      _ExtentY        =   2143
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
      SpreadDesigner  =   "FrmMDPriceSplit.frx":0000
      TextTip         =   2
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   2295
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   17055
      _Version        =   196613
      _ExtentX        =   30083
      _ExtentY        =   4048
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
      SpreadDesigner  =   "FrmMDPriceSplit.frx":4474
      TextTip         =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ҫ��ֳɼ��"
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ţ�"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "FrmMDPriceSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail�֭��
    E_SeqId = 1                '���
    E_CustID               '�ͻ�����
    E_CustName             '�ͻ�����
    E_PO                   'PO��
    E_PODate               '�ϴ�����
    E_POType               'PO����
    E_PT                   '����
    E_Qty                   '����
    E_Price                 '����
    E_PUnit                '���۵�Ԫ
    E_File                 '�ļ���
    E_OK                     'ѡ���
    

    
    E_End
    
End Enum



Private Enum E_FPS1          'Detail�֭��
    E_SeqId = 1                '���
    E_CustID               '�ͻ�����
    E_CustName             '�ͻ�����
    E_PO                   'PO��
    E_PODate               '�ϴ�����
    E_POType               'PO����
    E_PT                   '����
    E_Qty                   '����
    E_Price                 '����
    E_PUnit                '���۵�Ԫ
    E_File                 '�ļ���
    E_OK                     'ѡ���
    

    
    E_End
    
End Enum




Dim reportRS As New ADODB.Recordset
Dim mainItemRS As New ADODB.Recordset
Dim bomRS2        As New ADODB.Recordset

'Private Sub CmbCustomer_Change()
'TxtQtechPT.SetFocus
'End Sub

Private Sub CmdAdd_Click()
Dim nPIProductTemp As NpiProduct
Dim userId As String

'У���Ƿ��ظ�

If UCase(Trim(CmbCustomer.Text)) = "" Or UCase(Trim(TxtQtechPT.Text)) = "" Then
     MsgBox "�ͻ����������Ŀ���Ʋ�����Ϊ�գ�"
     Exit Sub
End If

If UCase(Trim(TxtCustPT1.Text)) = "" And UCase(Trim(TxtCustPT2.Text)) = "" Then
     MsgBox "�ͻ����ֲ�����Ϊ�գ�"
     Exit Sub
End If


 Set bomRS2 = GetNpiProductCheck(UCase(Trim(CmbCustomer.Text)), UCase(Trim(TxtQtechPT.Text)), UCase(Trim(TxtCustPT1.Text)), UCase(Trim(TxtCustPT2.Text)), UCase(Trim(TxtQtechPT2.Text)))
If bomRS2.RecordCount > 0 Then
    MsgBox "ϵͳ���Ѵ���������ݣ�������ȷ�������Ƿ���ȷ ��"
    Exit Sub
End If

userId = UCase(gUserName)

nPIProductTemp.CreateBy = UCase(gUserName)
nPIProductTemp.CustomershortName = UCase(Trim(CmbCustomer.Text))
nPIProductTemp.qtechPTNo = UCase(Trim(TxtQtechPT.Text))
nPIProductTemp.QtechPTNo2 = UCase(Trim(TxtQtechPT2.Text))
nPIProductTemp.CustomerPTNo1 = UCase(Trim(TxtCustPT1.Text))
nPIProductTemp.CustomerPTNo2 = UCase(Trim(TxtCustPT2.Text))
nPIProductTemp.CustomerDieQty = UCase(Trim(TxtTestu.Text))
nPIProductTemp.QtechDieQty = UCase(Trim(Txtfzr.Text))
nPIProductTemp.XiangSu = UCase(Trim(TxtTestR.Text))
nPIProductTemp.UsedArea = UCase(Trim(TxtNreF.Text))
nPIProductTemp.StruckStr1 = UCase(Trim(TxtNreW.Text))
nPIProductTemp.StruckStr2 = UCase(Trim(TxtHis2.Text))
nPIProductTemp.StruckStr3 = UCase(Trim(TxtHis1.Text))
nPIProductTemp.STDate = IIf(IsNull(DTPicker1.Value), "", DTPicker1.Value)
nPIProductTemp.TTDate = IIf(IsNull(DTPicker2.Value), "", DTPicker2.Value)
nPIProductTemp.PTDate = IIf(IsNull(DTPicker3.Value), "", DTPicker3.Value)


Call AddNpiProduct(nPIProductTemp)

 MsgBox "�����ɹ�!", vbInformation, "������ʾ"
 
 ShowData_Where

End Sub

Private Sub ChkAll_Click()
Dim i As Integer
    If ChkAll.Value = 1 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .Text = 1
            End With
        Next i
        
    ElseIf ChkAll.Value = 0 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .Text = 0
            End With
        Next i
        
    End If

End Sub

'Private Sub CmbCustomer_Change()
'Dim customerTemp As String
'
'customerTemp = CmbCustomer.Text
'
'TxtCusName.Text = GetCustomerNameSqlServer(customerTemp)
'
'
''���ݿͻ����룬��ʼ������
'
'IniCustomerPOName (customerTemp)
'
'
'
'End Sub

Private Sub CmbPONum_Change()
 Dim poTemp As String
Dim custTemp As String

TxtCreDate.Text = ""
TxtPT.Text = ""
TxtQty.Text = ""


poTemp = UCase(Trim(CmbPONum.Text))
custTemp = UCase(Trim(CmbCustomer.Text))

 Set oiRS = GetOIDataPONum(custTemp, poTemp)
    If (oiRS.RecordCount > 0) Then

        TxtCreDate.Text = oiRS.fields("qtech_created_date").Value
        TxtPT.Text = oiRS.fields("mpn_desc").Value
        TxtQty.Text = oiRS.fields("qty").Value
        
     End If
     


End Sub



Private Sub CmdDel_Click()
'��ѯ




ShowData_WhereCus (UCase(Trim(TxtPO.Text)))


End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdModify_Click()
'�޸�

Dim nPOTemp As POPrice
Dim userId As String
userId = UCase(gUserName)

nPOTemp.CreateBy = UCase(gUserName)




'-----------
With fps(0)

For i = 1 To .MaxRows

    .Row = i
    .Col = 12
    If .Text = 1 Then
    
    'Ҫ�޸�

    .Row = i
    .Col = 1
    nPOTemp.id = .Text
    
    
    .Row = i
    .Col = 4
    nPOTemp.PONo = .Text
    
    .Row = i
    .Col = 5
    nPOTemp.PODate = .Text
    
     .Row = i
    .Col = 6
    nPOTemp.POType = .Text
    
    .Row = i
    .Col = 7
    nPOTemp.PT = .Text
    
        
    .Row = i
    .Col = 8
    nPOTemp.qty = .Text
    
    
    .Row = i
    .Col = 9
    nPOTemp.Price = .Text
    
    .Row = i
    .Col = 10
    nPOTemp.unit = .Text
    
    .Row = i
    .Col = 11
    nPOTemp.File = .Text
    
    
    

    Call ModifyPOPrice(nPOTemp)

    
    
    End If
    

Next i


End With

'--------------------


 MsgBox "�޸ĳɹ�!", vbInformation, "������ʾ"
 
 ShowData_Where
 
End Sub

Private Sub CmdOutReport_Click()
Dim sqlTemp As String

sqlTemp = "select  id  , CUSTOMERSHORTNAME as �ͻ����� , QtechPTNo as ������Ŀ���� ,QtechPTNo2 as ��Ʒ�Ϻ�, CUSTOMERPTNo1  as �ͻ�������1, CUSTOMERPTNo2 as �ͻ�������2 , " & _
         " CUSTOMERDieQty as �ͻ����die��, XiangSu  as ����,  fzFreeUSD as ��װ��USD,testFreeUSD as ���Է�USD,fzFreeRMB as ��װ��RMB,testFreeRMB as ���Է�RMB,nreFree as NRE����YN��Ʊ����,nreMethod as NRE������ʽ,updatePrice2 as ������ʷ2,updatePrice1 as ������ʷ1 " & _
         " From TBLTsvNpiProduct where flag='Y' order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
         
  ExporToExcel (sqlTemp)

End Sub

Private Sub CmdSave_Click()
'У�������Ƿ��������
Dim qtySum As Long
Dim qtySumBefor As Long
Dim nPOTemp As POPrice
Dim oldRecord As Long

qtySum = 0
qtySumBefor = 0

nPOTemp.CreateBy = UCase(gUserName)


With fps(0)

    .Row = 1
    .Col = 12
    If .Text = 1 Then
    
        .Row = 1
        .Col = 1
        oldRecord = CLng(.Text)

        .Row = 1
        .Col = 8
        qtySumBefor = CInt(.Text)
        
    End If
       
End With


With fps(1)
    For i = 1 To .MaxRows
    
        .Row = i
        .Col = 8
        qtySum = qtySum + CInt(.Text)
       
    Next i
End With

If qtySum = qtySumBefor Then

'�����ݲ������

'�Ѳ���ǰ�ģ�״̬�ص�


Call UpdatePOPriceStatus(oldRecord, nPOTemp.CreateBy)




With fps(1)

For i = 1 To .MaxRows



 
    
    'Ҫ�޸�
    nPOTemp.id = GetPOPriceID()
    
      .Row = i
    .Col = 2
    nPOTemp.CustomershortName = .Text
    
     .Row = i
    .Col = 3
    nPOTemp.CustomerName = .Text
    
    
    
    .Row = i
    .Col = 4
    nPOTemp.PONo = .Text
    
    .Row = i
    .Col = 5
    nPOTemp.PODate = .Text
    
     .Row = i
    .Col = 6
    nPOTemp.POType = .Text
    
    .Row = i
    .Col = 7
    nPOTemp.PT = .Text
    
        
    .Row = i
    .Col = 8
    nPOTemp.qty = .Text
    
    
    .Row = i
    .Col = 9
    nPOTemp.Price = .Text
    
    .Row = i
    .Col = 10
    nPOTemp.unit = .Text
    
    .Row = i
    .Col = 11
    nPOTemp.File = .Text
    
    Call AddPOPrice(nPOTemp)

Next i


End With


 MsgBox "��PO�ɹ�!", vbInformation, "������ʾ"



Else

 MsgBox "��������ȷ�������º˶�!", vbInformation, "������ʾ"

End If





End Sub

Private Sub Command1_Click()
 With fps(1)

    .MaxRows = CInt(Trim(Txtno.Text))
    

End With
End Sub

Private Sub Form_Activate()
TxtPO.SetFocus

End Sub

Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CmbCustomer.RowSource = mainItemRS
CmbCustomer.ListField = mainItemRS("productname").Name
CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub


Private Sub IniCustomerPOName(customerTemp As String)
Set mainItemRS = GetCustomerPONum(customerTemp)
Set CmbPONum.RowSource = mainItemRS
CmbPONum.ListField = mainItemRS("productname").Name
CmbPONum.BoundColumn = mainItemRS("PID").Name



End Sub



Private Sub Form_Load()

'IniCustomerName

'IniCustomerName
IniFpsHeader

IniFpsHeader1

'DTPicker1.Value = DateTime.Date
'DTPicker2.Value = DateTime.Date
'DTPicker3.Value = DateTime.Date

'DTPicker1.MultiSelect = True
'DTPicker2.MultiSelect = True
'DTPicker3.MultiSelect = True


'DTPicker1.Value = Null
'DTPicker2.Value = Null
'DTPicker3.Value = Null

'ShowData_Where

'�����û���,���Ƿ����޸ĵ�Ȩ��

'Call UserType(UCase(gUserName))



End Sub

Private Sub UserType(nametemp As String)

If nametemp = "11040" Then
CmdAdd.Enabled = True
CmdModify.Enabled = True

Else

CmdAdd.Enabled = False

CmdModify.Enabled = False

End If



End Sub

'Private Sub IniCustomerName()
'Set mainItemRS = GetJDCustomerName()
'Set CmbCustomer.RowSource = mainItemRS
'CmbCustomer.ListField = mainItemRS("productname").Name
'CmbCustomer.BoundColumn = mainItemRS("PID").Name
'
'End Sub


Private Sub ShowData_Where()
Set reportRS = GetPOPriceModify()

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub

Private Sub ShowData_WhereCus(poNumTemp As String)
Set reportRS = GetPOPriceModify3(poNumTemp)

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub




Private Sub IniFpsHeader()
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
        
             .Col = E_FPS0.E_OK
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
    

        
        
        .SetText E_FPS0.E_SeqId, 0, "��¼��"
        .SetText E_FPS0.E_CustID, 0, "�ͻ�����"
        .SetText E_FPS0.E_CustName, 0, "�ͻ�ȫ��"
        .SetText E_FPS0.E_PO, 0, "������"
        .SetText E_FPS0.E_PODate, 0, "�ṩ��������"
        .SetText E_FPS0.E_POType, 0, "��������"
        .SetText E_FPS0.E_PT, 0, "����"
        .SetText E_FPS0.E_Qty, 0, "��������"
        
        .SetText E_FPS0.E_Price, 0, "����"
        .SetText E_FPS0.E_PUnit, 0, "���۵�λ"
        .SetText E_FPS0.E_File, 0, "�����ļ���"
         .SetText E_FPS0.E_OK, 0, "ѡ��"
        

          
       
        .ColWidth(E_FPS0.E_SeqId) = 5
        .ColWidth(E_FPS0.E_CustID) = 8
        .ColWidth(E_FPS0.E_CustName) = 25
        .ColWidth(E_FPS0.E_PO) = 15
        .ColWidth(E_FPS0.E_PODate) = 10
        .ColWidth(E_FPS0.E_POType) = 8
        .ColWidth(E_FPS0.E_PT) = 10
        .ColWidth(E_FPS0.E_Qty) = 10
        
        
        .ColWidth(E_FPS0.E_Price) = 5
        .ColWidth(E_FPS0.E_PUnit) = 5
        .ColWidth(E_FPS0.E_File) = 15
        
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .Col = E_FPS0.E_PO
                .Lock = False
        .Col = E_FPS0.E_PODate
                .Lock = False
        .Col = E_FPS0.E_POType
                .Lock = False
        .Col = E_FPS0.E_PT
                .Lock = False
        .Col = E_FPS0.E_Qty
                .Lock = False
        .Col = E_FPS0.E_Price
                .Lock = False
        .Col = E_FPS0.E_PUnit
                .Lock = False
        .Col = E_FPS0.E_File
                .Lock = False
                                             
         
      .Col = E_FPS0.E_OK
        .Lock = False
        
        
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub IniFpsHeader1()
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
        
             .Col = E_FPS1.E_OK
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
    

        
        
        .SetText E_FPS1.E_SeqId, 0, "��¼��"
        .SetText E_FPS1.E_CustID, 0, "�ͻ�����"
        .SetText E_FPS1.E_CustName, 0, "�ͻ�ȫ��"
        .SetText E_FPS1.E_PO, 0, "������"
        .SetText E_FPS1.E_PODate, 0, "�ṩ��������"
        .SetText E_FPS1.E_POType, 0, "��������"
        .SetText E_FPS1.E_PT, 0, "����"
        .SetText E_FPS1.E_Qty, 0, "��������"
        
        .SetText E_FPS1.E_Price, 0, "����"
        .SetText E_FPS1.E_PUnit, 0, "���۵�λ"
        .SetText E_FPS1.E_File, 0, "�����ļ���"
         .SetText E_FPS1.E_OK, 0, "ѡ��"
        

          
       
        .ColWidth(E_FPS1.E_SeqId) = 5
        .ColWidth(E_FPS1.E_CustID) = 8
        .ColWidth(E_FPS1.E_CustName) = 25
        .ColWidth(E_FPS1.E_PO) = 15
        .ColWidth(E_FPS1.E_PODate) = 10
        .ColWidth(E_FPS1.E_POType) = 8
        .ColWidth(E_FPS1.E_PT) = 10
        .ColWidth(E_FPS1.E_Qty) = 10
        
        
        .ColWidth(E_FPS1.E_Price) = 5
        .ColWidth(E_FPS1.E_PUnit) = 5
        .ColWidth(E_FPS1.E_File) = 15
        
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
                
         .Col = E_FPS1.E_SeqId
                .Lock = False
        
         .Col = E_FPS1.E_CustID
                .Lock = False
        .Col = E_FPS1.E_CustName
                .Lock = False
                
        
        
        .Col = E_FPS1.E_PO
                .Lock = False
        .Col = E_FPS1.E_PODate
                .Lock = False
        .Col = E_FPS1.E_POType
                .Lock = False
        .Col = E_FPS1.E_PT
                .Lock = False
        .Col = E_FPS1.E_Qty
                .Lock = False
        .Col = E_FPS1.E_Price
                .Lock = False
        .Col = E_FPS1.E_PUnit
                .Lock = False
        .Col = E_FPS1.E_File
                .Lock = False
                                             
         
      .Col = E_FPS1.E_OK
        .Lock = False
        
        
        .ReDraw = True
    End With
    
    
    

End Sub




'Private Sub fps_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
'Dim i As Long
'
'With fps(0)
'            .Row = Row
'            .Col = 1
'       i = .Text
'
'End With
'
'showData (i)
'
'Txtfzu.SetFocus
'
'End Sub

Private Sub showData(i As Long)

Set reportRS = GetNPIDataIDPrice(i)


 If reportRS.RecordCount > 0 Then
 
 
'    CmbCustomer.Text = reportRS.fields("CustomershortName").Value & ""
'    TxtQtechPT.Text = reportRS.fields("QtechPTNo").Value & ""
'    TxtQtechPT2.Text = reportRS.fields("QtechPTNo2").Value & ""
'    TxtCustPT1.Text = reportRS.fields("CustomerPTNo1").Value & ""
'    TxtCustPT2.Text = reportRS.fields("CustomerPTNo2").Value & ""
'    TxtTestu.Text = reportRS.fields("CustomerDieQty").Value & ""
'    Txtfzr.Text = reportRS.fields("QtechDieQty").Value & ""
'    TxtTestR.Text = reportRS.fields("XiangSu").Value & ""
'    TxtNreF.Text = reportRS.fields("UsedArea").Value & ""
'    TxtNreW.Text = reportRS.fields("StruckStr1").Value & ""
'    TxtHis2.Text = reportRS.fields("StruckStr2").Value & ""
'    TxtHis1.Text = reportRS.fields("StruckStr3").Value & ""
'    DTPicker1.Value = reportRS.fields("ST_DATE").Value
'    DTPicker2.Value = reportRS.fields("TT_DATE").Value
'    DTPicker3.Value = reportRS.fields("PT_DATE").Value
    
    
    Txtfzu.Text = reportRS.fields("fzFreeUSD").Value & ""
    TxtTestu.Text = reportRS.fields("testFreeUSD").Value & ""
    Txtfzr.Text = reportRS.fields("fzFreeRMB").Value & ""
    TxtTestR.Text = reportRS.fields("testFreeRMB").Value & ""
    TxtNreF.Text = reportRS.fields("nreFree").Value & ""
    TxtNreW.Text = reportRS.fields("nreMethod").Value & ""
    TxtHis2.Text = reportRS.fields("updatePrice2").Value & ""
    TxtHis1.Text = reportRS.fields("updatePrice1").Value & ""


    
    TxtIDTemp.Caption = reportRS.fields("ID").Value
 End If



End Sub

Private Sub TxtQtechPT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtQtechPT2.SetFocus
End If

End Sub

Private Sub TxtQtechPT2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtCustPT1.SetFocus
End If

End Sub

Private Sub TxtCustPT1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtCustPT2.SetFocus
End If

End Sub

Private Sub TxtCustPT2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtTestu.SetFocus
End If

End Sub

Private Sub Txtfzu_KeyPress(KeyAscii As Integer)


Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46) + Chr(13)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
TxtTestu.SetFocus
End If


End Sub

Private Sub TxtTestu_KeyPress(KeyAscii As Integer)
Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46) + Chr(13)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If


If KeyAscii = 13 Then
Txtfzr.SetFocus
End If

End Sub

Private Sub Txtfzr_KeyPress(KeyAscii As Integer)

Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46) + Chr(13)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If


If KeyAscii = 13 Then
TxtTestR.SetFocus
End If

End Sub


Private Sub TxtTestR_KeyPress(KeyAscii As Integer)

Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46) + Chr(13)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
TxtNreF.SetFocus
End If

End Sub

Private Sub TxtNreF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtNreW.SetFocus
End If

End Sub

Private Sub TxtNreW_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtHis2.SetFocus
End If

End Sub

Private Sub TxtHis2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtHis1.SetFocus
End If

End Sub



Private Sub TxtPrice_KeyPress(KeyAscii As Integer)
Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46) + Chr(13)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
CmbPUnit.SetFocus
End If
End Sub
