VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmMDPriceCreate 
   Caption         =   "�г��������۸�ά��"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20280
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
   ScaleWidth      =   20280
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtPease 
      Height          =   375
      Left            =   9480
      TabIndex        =   24
      Top             =   780
      Width           =   2535
   End
   Begin VB.TextBox TxtFile 
      Height          =   375
      Left            =   9480
      TabIndex        =   22
      Top             =   1380
      Width           =   2535
   End
   Begin VB.ComboBox CmbPUnit 
      Height          =   315
      ItemData        =   "FrmMDPriceCreate.frx":0000
      Left            =   5520
      List            =   "FrmMDPriceCreate.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1440
      Width           =   2535
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      ItemData        =   "FrmMDPriceCreate.frx":001C
      Left            =   1680
      List            =   "FrmMDPriceCreate.frx":002C
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "�˳�"
      Height          =   360
      Left            =   9360
      TabIndex        =   12
      Top             =   2160
      Width           =   990
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "���"
      Height          =   360
      Left            =   6840
      TabIndex        =   11
      Top             =   2160
      Width           =   990
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "����"
      Height          =   360
      Left            =   4320
      TabIndex        =   10
      Top             =   2160
      Width           =   990
   End
   Begin VB.TextBox TxtPrice 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   1380
      Width           =   2535
   End
   Begin VB.TextBox TxtQty 
      Height          =   375
      Left            =   13680
      TabIndex        =   7
      Top             =   780
      Width           =   2535
   End
   Begin VB.TextBox TxtPT 
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   780
      Width           =   2535
   End
   Begin VB.TextBox TxtCreDate 
      Height          =   375
      Left            =   13680
      TabIndex        =   3
      Top             =   180
      Width           =   2535
   End
   Begin VB.TextBox TxtCusName 
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   180
      Width           =   2535
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   6855
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   2880
      Width           =   17055
      _Version        =   196613
      _ExtentX        =   30083
      _ExtentY        =   12091
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
      SpreadDesigner  =   "FrmMDPriceCreate.frx":0058
      TextTip         =   2
   End
   Begin MSDataListLib.DataCombo CmbCustomer 
      Height          =   315
      Left            =   1680
      TabIndex        =   14
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo CmbPONum 
      Height          =   315
      Left            =   9480
      TabIndex        =   16
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������(Ƭ)��"
      Height          =   195
      Left            =   8280
      TabIndex        =   25
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ļ�����"
      Height          =   195
      Left            =   8400
      TabIndex        =   23
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���۵�λ��"
      Height          =   195
      Left            =   4560
      TabIndex        =   20
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ۣ�"
      Height          =   195
      Left            =   960
      TabIndex        =   19
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ţ�"
      Height          =   195
      Left            =   8640
      TabIndex        =   17
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ����룺"
      Height          =   195
      Left            =   600
      TabIndex        =   15
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������(Die)��"
      Height          =   195
      Left            =   12360
      TabIndex        =   8
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���֣�"
      Height          =   195
      Left            =   4800
      TabIndex        =   6
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������ͣ�"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ṩ�������ڣ�"
      Height          =   195
      Left            =   12360
      TabIndex        =   2
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ�ȫ�ƣ�"
      Height          =   195
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "FrmMDPriceCreate"
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
    E_PeaceQty                   '����
    E_Qty                   '����
    E_Price                 '����
    E_PUnit                '���۵�Ԫ
    E_File                 '�ļ���

    
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
Dim userid As String

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

userid = UCase(gUserName)

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

Private Sub CmbCustomer_Change()
Dim customerTemp As String

customerTemp = CmbCustomer.Text

TxtCusName.Text = GetCustomerNameSqlServer(customerTemp)


'���ݿͻ����룬��ʼ������

IniCustomerPOName (customerTemp)



End Sub

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
        TxtPease.Text = oiRS.fields("qty2").Value
        
        
     End If
     


End Sub



Private Sub CmdDel_Click()
TxtCusName.Text = ""
TxtCreDate.Text = ""
TxtPT.Text = ""
TxtQty.Text = ""
TxtPease.Text = ""

TxtPrice.Text = ""
TxtFile.Text = ""



End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdModify_Click()
'���

Dim nPOTemp As POPrice
Dim userid As String
userid = UCase(gUserName)

nPOTemp.CreateBy = UCase(gUserName)


If UCase(Trim(CmbCustomer.Text)) = "" Or UCase(Trim(TxtCusName.Text)) = "" Then
     MsgBox "�ͻ������ͻ����Ʋ�����Ϊ�գ�"
     Exit Sub
End If


If UCase(Trim(TxtQty.Text)) = "" Or UCase(Trim(TxtPrice.Text)) = "" Or UCase(Trim(CmbPUnit.Text)) = "" Then
     MsgBox "���������ۣ���λ ������Ϊ�գ�"
     Exit Sub
End If

nPOTemp.CreateBy = UCase(gUserName)
nPOTemp.id = GetPOPriceID()
nPOTemp.customerName = UCase(Trim(TxtCusName.Text))
nPOTemp.CustomershortName = UCase(Trim(CmbCustomer.Text))
nPOTemp.PONo = UCase(Trim(CmbPONum.Text))
nPOTemp.PODate = UCase(Trim(TxtCreDate.Text))
nPOTemp.POType = UCase(Trim(CmbType.Text))
nPOTemp.PT = UCase(Trim(TxtPT.Text))
nPOTemp.qty = UCase(Trim(TxtQty.Text))
nPOTemp.peaseQty = UCase(Trim(TxtPease.Text))

nPOTemp.Price = UCase(Trim(TxtPrice.Text))
nPOTemp.unit = UCase(Trim(CmbPUnit.Text))
nPOTemp.File = UCase(Trim(TxtFile.Text))





Call AddPOPrice(nPOTemp)

 MsgBox "�����ɹ�!", vbInformation, "������ʾ"
 
 ShowData_Where
 
End Sub

Private Sub CmdOutReport_Click()
Dim sqltemp As String

sqltemp = "select  id  , CUSTOMERSHORTNAME as �ͻ����� , QtechPTNo as ������Ŀ���� ,QtechPTNo2 as ��Ʒ�Ϻ�, CUSTOMERPTNo1  as �ͻ�������1, CUSTOMERPTNo2 as �ͻ�������2 , " & _
         " CUSTOMERDieQty as �ͻ����die��, XiangSu  as ����,  fzFreeUSD as ��װ��USD,testFreeUSD as ���Է�USD,fzFreeRMB as ��װ��RMB,testFreeRMB as ���Է�RMB,nreFree as NRE����YN��Ʊ����,nreMethod as NRE������ʽ,updatePrice2 as ������ʷ2,updatePrice1 as ������ʷ1 " & _
         " From TBLTsvNpiProduct where flag='Y' order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
         
  ExporToExcel (sqltemp)

End Sub

Private Sub Form_Activate()
'CmbCustomer.SetFocus

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
IniFpsHeader

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
Set reportRS = GetPOPrice()

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
        
    
        
        .SetText E_FPS0.E_SeqId, 0, "��¼��"
        .SetText E_FPS0.E_CustID, 0, "�ͻ�����"
        .SetText E_FPS0.E_CustName, 0, "�ͻ�ȫ��"
        .SetText E_FPS0.E_PO, 0, "������"
        .SetText E_FPS0.E_PODate, 0, "�ṩ��������"
        .SetText E_FPS0.E_POType, 0, "��������"
        .SetText E_FPS0.E_PT, 0, "����"
        
        .SetText E_FPS0.E_PeaceQty, 0, "����Ƭ��"
        .SetText E_FPS0.E_Qty, 0, "��������"
        
        .SetText E_FPS0.E_Price, 0, "����"
        .SetText E_FPS0.E_PUnit, 0, "���۵�λ"
        .SetText E_FPS0.E_File, 0, "�����ļ���"
       
        .ColWidth(E_FPS0.E_SeqId) = 5
        .ColWidth(E_FPS0.E_CustID) = 8
        .ColWidth(E_FPS0.E_CustName) = 25
        .ColWidth(E_FPS0.E_PO) = 15
        .ColWidth(E_FPS0.E_PODate) = 10
        .ColWidth(E_FPS0.E_POType) = 8
         
        .ColWidth(E_FPS0.E_PT) = 10
        
         .ColWidth(E_FPS0.E_PeaceQty) = 9
        
        .ColWidth(E_FPS0.E_Qty) = 10
        
        
        .ColWidth(E_FPS0.E_Price) = 5
        .ColWidth(E_FPS0.E_PUnit) = 5
        .ColWidth(E_FPS0.E_File) = 15
        
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub



Private Sub fps_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i As Long

With fps(0)
            .Row = Row
            .Col = 1
       i = .Text

End With

showData (i)

Txtfzu.SetFocus

End Sub

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
