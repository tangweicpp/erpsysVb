VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmMDPriceQuery 
   Caption         =   "�г��������۸��ѯ"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20250
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
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   360
      Left            =   8160
      TabIndex        =   14
      Top             =   1680
      Width           =   990
   End
   Begin VB.TextBox TxtPT 
      Height          =   375
      Left            =   10320
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      ItemData        =   "FrmMDPriceQuery.frx":0000
      Left            =   6120
      List            =   "FrmMDPriceQuery.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox TxtPO 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "��ѯ"
      Height          =   360
      Left            =   5040
      TabIndex        =   0
      Top             =   1680
      Width           =   990
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   7455
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   17055
      _Version        =   524288
      _ExtentX        =   30083
      _ExtentY        =   13150
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
      SpreadDesigner  =   "FrmMDPriceQuery.frx":003C
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin MSDataListLib.DataCombo CmbCustomer 
      Height          =   315
      Left            =   10320
      TabIndex        =   5
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   107020289
      CurrentDate     =   41424
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   107020289
      CurrentDate     =   41424
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������ʱ�䣺"
      Height          =   195
      Left            =   4800
      TabIndex        =   13
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ʼʱ�䣺"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������ͣ�"
      Height          =   195
      Left            =   5160
      TabIndex        =   9
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���֣�"
      Height          =   195
      Left            =   9600
      TabIndex        =   8
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ����룺"
      Height          =   195
      Left            =   9360
      TabIndex        =   7
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ţ�"
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Width           =   720
   End
End
Attribute VB_Name = "FrmMDPriceQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail�֭��

    E_SeqId = 1                '���
    E_CustID               '�ͻ�����
    E_CUSTNAME             '�ͻ�����
    e_PO                   'PO��
    E_PODATE               '�ϴ�����
    E_CreateDate           '��������
    E_POTYPE               'PO����
    E_PT                   '����
    E_PeaceQty                   '����
    E_ShengYuQty                   '����
    
    E_QTY                   '����
    e_Price                 '����
    E_PUnit                '���۵�Ԫ
    E_File                 '�ļ���
    E_CustAA                 '�ͻ���д
    E_BJ                '���۵���
    
    E_END
    
End Enum

Private Enum E_FPS1          'Detail�֭��

    E_SeqId = 1                '���
    E_CustID               '�ͻ�����
    E_CUSTNAME             '�ͻ�����
    e_PO                   'PO��
    E_PODATE               '�ϴ�����
    E_POTYPE               'PO����
    E_PT                   '����
    E_PeaceQty                   '����
    E_QTY                   '����
    e_Price                 '����
    E_PUnit                '���۵�Ԫ
    E_File                 '�ļ���
    E_OK                     'ѡ���
    E_CustAA                 '�ͻ���д
    E_BJ                '���۵���
    
    E_END
    
End Enum

Dim reportRS   As New ADODB.Recordset

Dim mainItemRS As New ADODB.Recordset

Dim bomRS2     As New ADODB.Recordset

'Private Sub CmbCustomer_Change()
'TxtQtechPT.SetFocus
'End Sub

Private Sub CmdAdd_Click()

    Dim nPIProductTemp As NpiProduct

    Dim userid         As String

    'У���Ƿ��ظ�

    If UCase(Trim(CmbCustomer.Text)) = "" Or UCase(Trim(TxtQtechPT.Text)) = "" Then
        MsgBox "�ͻ����������Ŀ���Ʋ�����Ϊ�գ�"

        Exit Sub

    End If

    If UCase(Trim(TxtCustPT1.Text)) = "" And UCase(Trim(TxtCustPT2.Text)) = "" Then
        MsgBox "�ͻ����ֲ�����Ϊ�գ�"

        Exit Sub

    End If

   ' Set bomRS2 = GetNpiProductCheck(UCase(Trim(CmbCustomer.Text)), UCase(Trim(TxtQtechPT.Text)), UCase(Trim(TxtCustPT1.Text)), UCase(Trim(TxtCustPT2.Text)), UCase(Trim(TxtQtechPT2.Text)))

    'If bomRS2.RecordCount > 0 Then
     '   MsgBox "ϵͳ���Ѵ���������ݣ�������ȷ�������Ƿ���ȷ ��"

      '  Exit Sub

    'End If

    userid = UCase(gUserName)

    nPIProductTemp.CreateBy = UCase(gUserName)
    nPIProductTemp.CUSTOMERSHORTNAME = UCase(Trim(CmbCustomer.Text))
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

    Dim potemp   As String

    Dim custTemp As String

    TxtCreDate.Text = ""
    TxtPT.Text = ""
    txtQty.Text = ""

    potemp = UCase(Trim(CmbPONum.Text))
    custTemp = UCase(Trim(CmbCustomer.Text))

    Set oiRS = GetOIDataPONum(custTemp, potemp)

    If (oiRS.RecordCount > 0) Then

        TxtCreDate.Text = oiRS.Fields("qtech_created_date").Value
        TxtPT.Text = oiRS.Fields("mpn_desc").Value
        txtQty.Text = oiRS.Fields("qty").Value
        
    End If

End Sub

Private Sub cmdDel_Click()
    '��ѯ

    Dim beginTime    As String

    Dim endTime      As String

    Dim woTemp       As String

    Dim sqlTemp      As String

    Dim sql1         As String

    Dim billNoTemp   As String

    Dim lotIdTemp    As String

    Dim bigQboxTemp  As String

    Dim productTemp  As String

    Dim date1Temp    As String

    Dim date2Temp    As String

    Dim customerTemp As String

    Dim sql2         As String

    Dim sql3         As String

    If Trim(txtPO.Text) <> "" Then   '��PO�Ų�ѯ

        billNoTemp = UCase(Trim(txtPO.Text))

        sqlTemp = " select   ID ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE ,qtech_created_date, PO_TYPE ,PT , PeaceQty,'' as shengyuqty, QTY ,PRICE ,UNIT,  " & "  FILENAME,CUSTAA,BJ   from  TSV_MD_POPrice where flag='Y'  and  PO_NUM ='" & billNoTemp & "'order by id desc "

        Set mainItemRS = GetAAMPNDataSQL(sqlTemp)

    Else
        ' ���ǰ�PO�Ų�ѯ
        'productTemp = UCase(Trim(CmbType.Text))
     
        date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
        date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
     
        sqlTemp = " select   ID ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE ,qtech_created_date, PO_TYPE ,PT ,PeaceQty, '' as shengyuqty,QTY ,PRICE ,UNIT , " & "  FILENAME,CUSTAA,BJ   from  TSV_MD_POPrice where flag='Y'  and  PO_DATE >=to_date('" + date1Temp + "','YYYY-MM-DD')  and  PO_DATE <to_date('" + date2Temp + "','YYYY-MM-DD')  "
     
        If CmbCustomer.Text <> "" Then

            sqlTemp = sqlTemp & " and  CUSTOMERSHORTNAME = '" + UCase(Trim(CmbCustomer.Text)) + "' "

        End If
     
        If CmbType.Text <> "" Then

            sqlTemp = sqlTemp & " and  PO_TYPE = '" + UCase(Trim(CmbType.Text)) + "' "

        End If
     
        If TxtPT.Text <> "" Then

            sqlTemp = sqlTemp & " and  PT = '" + UCase(Trim(TxtPT.Text)) + "' "

        End If
     
        Set mainItemRS = GetAAMPNData(sqlTemp)
           
    End If

    With fpS(0)
        .MaxRows = 0

        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
       
        End If

    End With

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdModify_Click()
    '�޸�

    Dim nPOTemp As POPrice

    Dim userid  As String

    userid = UCase(gUserName)

    nPOTemp.CreateBy = UCase(gUserName)

    '-----------
    With fpS(0)

        For I = 1 To .MaxRows

            .Row = I
            .Col = 12

            If .Text = 1 Then
    
                'Ҫ�޸�

                .Row = I
                .Col = 1
                nPOTemp.ID = .Text
    
                .Row = I
                .Col = 4
                nPOTemp.PONo = .Text
    
                .Row = I
                .Col = 5
                nPOTemp.PODATE = .Text
    
                .Row = I
                .Col = 6
                nPOTemp.POType = .Text
    
                .Row = I
                .Col = 7
                nPOTemp.pt = .Text
        
                .Row = I
                .Col = 8
                nPOTemp.QTY = .Text
    
                .Row = I
                .Col = 9
                nPOTemp.Price = .Text
    
                .Row = I
                .Col = 10
                nPOTemp.unit = .Text
    
                .Row = I
                .Col = 11
                nPOTemp.File = .Text

                Call ModifyPOPrice(nPOTemp)
    
            End If

        Next I

    End With

    '--------------------

    MsgBox "�޸ĳɹ�!", vbInformation, "������ʾ"
 
    ShowData_Where
 
End Sub

Private Sub CmdOutReport_Click()

    Dim sqlTemp As String

    sqlTemp = "select  id  , CUSTOMERSHORTNAME as �ͻ����� , QtechPTNo as ������Ŀ���� ,QtechPTNo2 as ��Ʒ�Ϻ�, CUSTOMERPTNo1  as �ͻ�������1, CUSTOMERPTNo2 as �ͻ�������2 , " & " CUSTOMERDieQty as �ͻ����die��, XiangSu  as ����,  fzFreeUSD as ��װ��USD,testFreeUSD as ���Է�USD,fzFreeRMB as ��װ��RMB,testFreeRMB as ���Է�RMB,nreFree as NRE����YN��Ʊ����,nreMethod as NRE������ʽ,updatePrice2 as ������ʷ2,updatePrice1 as ������ʷ1 " & " From TBLTsvNpiProduct where flag='Y' order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
         
    ExporToExcel (sqlTemp)

End Sub

Private Sub CmdSave_Click()

    'У�������Ƿ��������
    Dim qtySum      As Long

    Dim qtySumBefor As Long

    Dim nPOTemp     As POPrice

    Dim oldRecord   As Long

    qtySum = 0
    qtySumBefor = 0

    nPOTemp.CreateBy = UCase(gUserName)

    With fpS(0)

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

    With fpS(1)

        For I = 1 To .MaxRows
    
            .Row = I
            .Col = 8
            qtySum = qtySum + CInt(.Text)
       
        Next I

    End With

    If qtySum = qtySumBefor Then

        '�����ݲ������

        '�Ѳ���ǰ�ģ�״̬�ص�

        Call UpdatePOPriceStatus(oldRecord, nPOTemp.CreateBy)

        With fpS(1)

            For I = 1 To .MaxRows
    
                'Ҫ�޸�
                nPOTemp.ID = GetPOPriceID()
    
                .Row = I
                .Col = 2
                nPOTemp.CUSTOMERSHORTNAME = .Text
    
                .Row = I
                .Col = 3
                nPOTemp.customerName = .Text
    
                .Row = I
                .Col = 4
                nPOTemp.PONo = .Text
    
                .Row = I
                .Col = 5
                nPOTemp.PODATE = .Text
    
                .Row = I
                .Col = 6
                nPOTemp.POType = .Text
    
                .Row = I
                .Col = 7
                nPOTemp.pt = .Text
        
                .Row = I
                .Col = 8
                nPOTemp.QTY = .Text
    
                .Row = I
                .Col = 9
                nPOTemp.Price = .Text
    
                .Row = I
                .Col = 10
                nPOTemp.unit = .Text
    
                .Row = I
                .Col = 11
                nPOTemp.File = .Text
    
                Call AddPOPrice(nPOTemp)

            Next I

        End With

        MsgBox "��PO�ɹ�!", vbInformation, "������ʾ"

    Else

        MsgBox "��������ȷ�������º˶�!", vbInformation, "������ʾ"

    End If

End Sub

Private Sub Command1_Click()

    Dim beginTime    As String

    Dim endTime      As String

    Dim woTemp       As String

    Dim sqlTemp      As String

    Dim sql1         As String

    Dim billNoTemp   As String

    Dim lotIdTemp    As String

    Dim bigQboxTemp  As String

    Dim productTemp  As String

    Dim date1Temp    As String

    Dim date2Temp    As String

    Dim customerTemp As String

    Dim sql2         As String

    Dim sql3         As String

    If Trim(txtPO.Text) <> "" Then   '��PO�Ų�ѯ

        billNoTemp = UCase(Trim(txtPO.Text))

        sqlTemp = " select   ID ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE ,qtech_created_date, PO_TYPE ,PT , PeaceQty ,'' as shengyuqty,QTY ,PRICE ,UNIT , " & "  FILENAME   from  TSV_MD_POPrice where flag='Y'  and  PO_NUM ='" & billNoTemp & "'order by id desc "

        ExporToExcel (sqlTemp)

    Else
        ' ���ǰ�PO�Ų�ѯ
        'productTemp = UCase(Trim(CmbType.Text))
     
        date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
        date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
     
        sqlTemp = " select   ID ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE , qtech_created_date,PO_TYPE ,PT ,PeaceQty, '' as shengyuqty,QTY ,PRICE ,UNIT , " & "  FILENAME   from  TSV_MD_POPrice where flag='Y'  and  PO_DATE >=to_date('" + date1Temp + "','YYYY-MM-DD')  and  PO_DATE <to_date('" + date2Temp + "','YYYY-MM-DD')  "
     
        If CmbCustomer.Text <> "" Then

            sqlTemp = sqlTemp & " and  CUSTOMERSHORTNAME = '" + UCase(Trim(CmbCustomer.Text)) + "' "

        End If
     
        If CmbType.Text <> "" Then

            sqlTemp = sqlTemp & " and  PO_TYPE = '" + UCase(Trim(CmbType.Text)) + "' "

        End If
     
        If TxtPT.Text <> "" Then

            sqlTemp = sqlTemp & " and  PT = '" + UCase(Trim(TxtPT.Text)) + "' "

        End If
     
        ExporToExcel (sqlTemp)
           
    End If

End Sub

Private Sub Form_Activate()
    txtPO.SetFocus

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
    IniCustomerName
    'IniCustomerName

    'IniCustomerName
    IniFpsHeader

    DTP1.Value = Now - 1

    DTP2.Value = Now

    'IniFpsHeader1

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
    Set reportRS = GetPOPriceModify(UCase(Trim(CmbCustomer.Text)), UCase(Trim(Comflag.Text)))

    With fpS(0)
        .MaxRows = 0

        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If

    End With

End Sub

Private Sub ShowData_WhereCus(ponumTemp As String)
    Set reportRS = GetPOPriceModify3(ponumTemp)

    With fpS(0)
        .MaxRows = 0

        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If

    End With

End Sub

Private Sub IniFpsHeader()

    With fpS(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_END - 1
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
        
        .SetText E_FPS0.E_SeqId, 0, "��¼��"
        .SetText E_FPS0.E_CustID, 0, "�ͻ�����"
        .SetText E_FPS0.E_CUSTNAME, 0, "�ͻ�ȫ��"
        .SetText E_FPS0.e_PO, 0, "������"
        .SetText E_FPS0.E_PODATE, 0, "�ṩ��������"
        .SetText E_FPS0.E_CreateDate, 0, "¼������"
        
        .SetText E_FPS0.E_POTYPE, 0, "��������"
        .SetText E_FPS0.E_PT, 0, "����"
        .SetText E_FPS0.E_PeaceQty, 0, "����Ƭ��"
        .SetText E_FPS0.E_ShengYuQty, 0, "ʣ��Ƭ��"
            
        .SetText E_FPS0.E_QTY, 0, "��������"
        
        .SetText E_FPS0.e_Price, 0, "����"
        .SetText E_FPS0.E_PUnit, 0, "���۵�λ"
        .SetText E_FPS0.E_File, 0, "�����ļ���"

        .SetText E_FPS0.E_CustAA, 0, "�ͻ���д"
        .SetText E_FPS0.E_BJ, 0, "���۵���"
       
        .ColWidth(E_FPS0.E_SeqId) = 5
        .ColWidth(E_FPS0.E_CustID) = 8
        .ColWidth(E_FPS0.E_CUSTNAME) = 25
        .ColWidth(E_FPS0.e_PO) = 15
        .ColWidth(E_FPS0.E_PODATE) = 10
        .ColWidth(E_FPS0.E_CreateDate) = 10
        .ColWidth(E_FPS0.E_POTYPE) = 8
        .ColWidth(E_FPS0.E_PT) = 10
        .ColWidth(E_FPS0.E_PeaceQty) = 9
        
        .ColWidth(E_FPS0.E_ShengYuQty) = 9
        
        .ColWidth(E_FPS0.E_QTY) = 10
        
        .ColWidth(E_FPS0.e_Price) = 5
        .ColWidth(E_FPS0.E_PUnit) = 5
        .ColWidth(E_FPS0.E_File) = 15
        .ColWidth(E_FPS0.E_CustAA) = 5
        .ColWidth(E_FPS0.E_BJ) = 15

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .ReDraw = True
    End With

End Sub

Private Sub IniFpsHeader1()

    With fpS(1)
        .ReDraw = False
        .MaxCols = E_FPS1.E_END - 1
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
        .SetText E_FPS1.E_CUSTNAME, 0, "�ͻ�ȫ��"
        .SetText E_FPS1.e_PO, 0, "������"
        .SetText E_FPS1.E_PODATE, 0, "�ṩ��������"
        .SetText E_FPS1.E_POTYPE, 0, "��������"
        .SetText E_FPS1.E_PT, 0, "����"
        .SetText E_FPS1.E_PeaceQty, 0, "����Ƭ��"
           
        .SetText E_FPS1.E_QTY, 0, "��������"
        
        .SetText E_FPS1.e_Price, 0, "����"
        .SetText E_FPS1.E_PUnit, 0, "���۵�λ"
        .SetText E_FPS1.E_File, 0, "�����ļ���"
        .SetText E_FPS1.E_CustAA, 0, "�ͻ����"
        .SetText E_FPS1.E_BJ, 0, "���۵���"
        .SetText E_FPS1.E_OK, 0, "ѡ��"
       
        .ColWidth(E_FPS1.E_SeqId) = 5
        .ColWidth(E_FPS1.E_CustID) = 8
        .ColWidth(E_FPS1.E_CUSTNAME) = 25
        .ColWidth(E_FPS1.e_PO) = 15
        .ColWidth(E_FPS1.E_PODATE) = 10
        .ColWidth(E_FPS1.E_POTYPE) = 8
        .ColWidth(E_FPS1.E_PT) = 10
        .ColWidth(E_FPS1.E_PeaceQty) = 10
        
        .ColWidth(E_FPS1.E_QTY) = 10
        
        .ColWidth(E_FPS1.e_Price) = 5
        .ColWidth(E_FPS1.E_PUnit) = 5
        .ColWidth(E_FPS1.E_File) = 15
        .ColWidth(E_FPS1.E_CustAA) = 5
        .ColWidth(E_FPS1.E_BJ) = 15

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
                
        .Col = E_FPS1.E_SeqId
        .Lock = False
        
        .Col = E_FPS1.E_CustID
        .Lock = False
        .Col = E_FPS1.E_CUSTNAME
        .Lock = False
        
        .Col = E_FPS1.e_PO
        .Lock = False
        .Col = E_FPS1.E_PODATE
        .Lock = False
        .Col = E_FPS1.E_POTYPE
        .Lock = False
        .Col = E_FPS1.E_PT
        .Lock = False
                
        .Col = E_FPS1.E_PeaceQty
                   
        .Lock = False
                
        .Col = E_FPS1.E_QTY
        .Lock = False
        .Col = E_FPS1.e_Price
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

Private Sub ShowData(I As Long)

    Set reportRS = GetNPIDataIDPrice(I)

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
    
        Txtfzu.Text = reportRS.Fields("fzFreeUSD").Value & ""
        TxtTestu.Text = reportRS.Fields("testFreeUSD").Value & ""
        Txtfzr.Text = reportRS.Fields("fzFreeRMB").Value & ""
        TxtTestR.Text = reportRS.Fields("testFreeRMB").Value & ""
        TxtNreF.Text = reportRS.Fields("nreFree").Value & ""
        TxtNreW.Text = reportRS.Fields("nreMethod").Value & ""
        TxtHis2.Text = reportRS.Fields("updatePrice2").Value & ""
        TxtHis1.Text = reportRS.Fields("updatePrice1").Value & ""
    
        TxtIDTemp.Caption = reportRS.Fields("ID").Value
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

