VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmOnMarkingCode 
   Caption         =   "MPS,չѶ�ͻ���Ʊ���ݵ��ŵ���Ϣά��"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15960
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
   ScaleHeight     =   10155
   ScaleWidth      =   15960
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   3
      Left            =   11400
      TabIndex        =   19
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   2
      Left            =   11400
      TabIndex        =   18
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      ItemData        =   "FrmOnMarkingCode.frx":0000
      Left            =   7800
      List            =   "FrmOnMarkingCode.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   1
      Left            =   7800
      TabIndex        =   16
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   15
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���"
      Height          =   360
      Left            =   6120
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox TxtShiping 
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox TxtInvoice 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   360
      Left            =   4200
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox TxtCode 
      Height          =   375
      Left            =   15720
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtWsgPT 
      Height          =   375
      Left            =   15720
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox TxtOpn 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   360
      Left            =   17400
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ѯ"
      Height          =   360
      Left            =   15840
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   990
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   6495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   18735
      _Version        =   524288
      _ExtentX        =   33046
      _ExtentY        =   11456
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
      SpreadDesigner  =   "FrmOnMarkingCode.frx":0004
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Frame fra 
      Caption         =   "չѶ������Ϣά��"
      Height          =   1935
      Left            =   5880
      TabIndex        =   20
      Top             =   120
      Width           =   8055
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "specialMark(����������ʶ)��"
         Height          =   585
         Index           =   4
         Left            =   4200
         TabIndex        =   25
         Top             =   720
         Width           =   1305
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "trackingNo(����������)��"
         Height          =   390
         Index           =   3
         Left            =   4560
         TabIndex        =   24
         Top             =   240
         Width           =   765
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "selfPickup(�Ƿ�����)��"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1725
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "shipper(������)��"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "requestNumber��"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1530
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ShippingNumber��"
      Height          =   195
      Left            =   840
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "InvoiceNumber��"
      Height          =   195
      Left            =   840
      TabIndex        =   11
      Top             =   960
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MarkingCode:"
      Height          =   195
      Left            =   14640
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WSGPT:"
      Height          =   195
      Left            =   14760
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������ţ�"
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "FrmOnMarkingCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail�֭��
    E_SeqId = 1                '���
    E_PART
    E_WSGPART
    E_CODE
    E_REQUESTNUMBER
    E_SHIPPER
    E_SELFPICKUP
    E_TRACKINGNO
    E_SPECIALMARK
    E_END
    
End Enum

Dim reportRS As New ADODB.Recordset
Dim mainItemRS As New ADODB.Recordset
Dim bomRS2        As New ADODB.Recordset



Private Sub CmbCustomer_Change()
TxtQtechPT.SetFocus
End Sub

Private Sub CmbDept_Click()
txtWaferID.SetFocus
End Sub

Private Sub CmdAdd_Click()
Dim baoFeiTemp As Baofei
Dim userid As String

'У���Ƿ��ظ�

If UCase(Trim(CmbDept.Text)) = "" Then
     MsgBox "��ѡ���ű�"
     Exit Sub
End If

If UCase(Trim(txtWaferID.Text)) = "" And UCase(Trim(txtLotID.Text)) = "" Then
     MsgBox "��ȷ��WaferID,LotID!"
     Exit Sub
End If


userid = UCase(gUserName)

baoFeiTemp.CreateBy = userid

baoFeiTemp.putInDept = CmbDept.Text
baoFeiTemp.err_date = DTPicker1.Value
baoFeiTemp.putIn_date = DTPicker2.Value
baoFeiTemp.WaferID = txtWaferID.Text
baoFeiTemp.LOTID = txtLotID.Text
baoFeiTemp.gDDie = CInt(TxtGDDie.Text)
baoFeiTemp.ngdie = CInt(TxtNGDie.Text)
baoFeiTemp.customerPTNo = TxtCustPT.Text
baoFeiTemp.qtechPTNo = TxtHTPT.Text
baoFeiTemp.productName = TxtProduct.Text


Call AddBaoFei(baoFeiTemp)

 MsgBox "�����ɹ�!", vbInformation, "������ʾ"
 
 ShowData_Where

End Sub

Private Sub cmdDel_Click()
txtWaferID.Text = ""
txtLotID.Text = ""
TxtGDDie.Text = ""
TxtNGDie.Text = ""
TxtCustPT.Text = ""
TxtHTPT.Text = ""
TxtProduct.Text = ""

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdModify_Click()
'�޸�

Dim nPIProductTemp As NpiProduct
Dim userid As String
userid = UCase(gUserName)

nPIProductTemp.CreateBy = UCase(gUserName)
nPIProductTemp.CUSTOMERSHORTNAME = UCase(Trim(CmbCustomer.Text))
nPIProductTemp.qtechPTNo = UCase(Trim(TxtQtechPT.Text))
nPIProductTemp.QtechPTNo2 = UCase(Trim(TxtQtechPT2.Text))
nPIProductTemp.CustomerPTNo1 = UCase(Trim(TxtCustPT1.Text))
nPIProductTemp.CustomerPTNo2 = UCase(Trim(TxtCustPT2.Text))
nPIProductTemp.CustomerPTNo3 = UCase(Trim(TxtCustPT3.Text))
nPIProductTemp.CustomerPTNo4 = UCase(Trim(TxtCustPT4.Text))

nPIProductTemp.CustomerDieQty = UCase(Trim(TxtCustDie.Text))
nPIProductTemp.QtechDieQty = UCase(Trim(TxtQtechDie.Text))
nPIProductTemp.XiangSu = UCase(Trim(TxtXS.Text))
nPIProductTemp.UsedArea = UCase(Trim(TxtArea.Text))
nPIProductTemp.StruckStr1 = UCase(Trim(TxtStr1.Text))
nPIProductTemp.StruckStr2 = UCase(Trim(TxtStr2.Text))
nPIProductTemp.StruckStr3 = UCase(Trim(TxtStr3.Text))
nPIProductTemp.STDate = IIf(IsNull(DTPicker1.Value), "", DTPicker1.Value)
nPIProductTemp.TTDate = IIf(IsNull(DTPicker2.Value), "", DTPicker2.Value)
nPIProductTemp.PTDate = IIf(IsNull(DTPicker3.Value), "", DTPicker3.Value)


Call ModifyNpiProduct(nPIProductTemp, CLng(TxtIDTemp.Text))

 MsgBox "�޸ĳɹ�!", vbInformation, "������ʾ"

ShowData_Where

End Sub

Private Sub CmdOutReport_Click()
Dim sqlTemp As String

sqlTemp = "select  id ,PutInDept as �Ǽǲ��ű�, WaferID ,LotID ,GDDie,NGDie ,CUSTOMERPTNO as �ͻ����ֺ�, QTECHPTNO as ���ڻ��ֺ�,ProductName,ERR_DATE as �쳣���� ,PutIn_DATE as ¼������ from TBLTSVBaoFei where flag='Y'  order by PutIn_DATE desc,lotid,waferid "
  
         
  ExporToExcel (sqlTemp)

End Sub

Private Sub ChkAll_Click()
'Dim i As Integer
'    If ChkAll.Value = 1 Then
'        For i = 1 To fps(0).MaxRows
'            With fps(0)
'                .Row = i
'                .Col = E_FPS0.E_OK
'                .Text = 1
'            End With
'        Next i
'
'    ElseIf ChkAll.Value = 0 Then
'        For i = 1 To fps(0).MaxRows
'            With fps(0)
'                .Row = i
'                .Col = E_FPS0.E_OK
'                .Text = 0
'            End With
'        Next i
'
'    End If

End Sub

Private Sub CmbType_Click()

'
'If CmbType.Text = "���1" Then
'typeId = 2
'ElseIf CmbType.Text = "�Ƿ�֪ͨ�ͻ�" Then
'typeId = 3
'
'ElseIf CmbType.Text = "�ͻ�ͬ�ⱨ���г���ȷ��" Then
'typeId = 4
'
'
'End If
'
'Call ShowData_Where(typeId)

End Sub

Private Sub Command1_Click()
Dim userid As String

ShowData_Where


End Sub

Private Sub Command2_Click()
Dim temp As String

  temp = "  select  ID,MPNPART,WSGPART,MarkCodeFirst from  CUSTOMERMarkingCode  where flag='Y' order by id "
  
 ExporToExcel (temp)
 
End Sub

Private Sub Command3_Click()
Dim cmdStr      As String
Dim cmdStr2     As String
Dim idTemp      As Long
Dim mpnPTTemp   As String
Dim wsgPTTemp   As String
Dim mCodetemp   As String
Dim username    As String
Dim invTemp     As String
Dim shipTemp    As String
Dim strTmp(5)   As String
Dim billNoTemp  As String
Dim dateTemp    As String
Dim notemp      As Long
Dim noStrTemp   As String
Dim invoiceTemp As String
Dim strSql      As String

'�ж���ʷ������ţ��Ƿ���ȷ
If Trim(TxtOpn.Text) = "" Then
    MsgBox "�����뷢��������Ϣ���ٱ��棡", vbInformation, "������ʾ"
    Exit Sub

End If

strSql = "SELECT * FROM erpdata..tblStockSQfh where ���ݱ�� = '" & Trim(TxtOpn.Text) & "'"
If Get_SqlserverCnt(strSql) = 0 Then
    MsgBox "������ķ�Ʊ���ݺŲ�����, ��ȷ���Ƿ���ȷ", vbInformation, "��ʾ"
    Exit Sub

End If

username = gUserName
idTemp = GetMaxIDONMarkCode()
mpnPTTemp = UCase(Trim(TxtOpn.Text))
wsgPTTemp = UCase(Trim(TxtWsgPT.Text))
mCodetemp = UCase(Trim(txtCode.Text))
billNoTemp = Replace(Replace(UCase(Trim(TxtOpn.Text)), Chr(10), ""), vbCrLf, "")
invTemp = Replace(Replace(UCase(Trim(TxtInvoice.Text)), Chr(10), ""), vbCrLf, "")
shipTemp = Replace(Replace(UCase(Trim(TxtShiping.Text)), Chr(10), ""), vbCrLf, "")
strTmp(0) = Replace(Replace(UCase(Trim(txt(0).Text)), Chr(10), ""), vbCrLf, "")
strTmp(1) = Replace(Replace(UCase(Trim(txt(1).Text)), Chr(10), ""), vbCrLf, "")
strTmp(2) = Replace(Replace(UCase(Trim(cmb.Text)), Chr(10), ""), vbCrLf, "")
strTmp(3) = Replace(Replace(UCase(Trim(txt(2).Text)), Chr(10), ""), vbCrLf, "")
strTmp(4) = Replace(Replace(UCase(Trim(txt(3).Text)), Chr(10), ""), vbCrLf, "")

'
'�ж���û��ά����
If JudgeMPSInvStatus(billNoTemp) Then
    MsgBox "�˷������������ɹ���Ʊ��!", vbInformation, "������ʾ"
    Exit Sub

End If

cmdStr = " insert into CUSTOMERMarkingCode (id,mpnpart,wsgpart,flag,Qtech_Created_By,Qtech_Created_Date,markcodefirst,ZX_REQUESTNUMBER,ZX_SHIPPER,ZX_SELFPICKUP,ZX_TRACKINGNO,ZX_SPECIALMARK) values " & " (" & idTemp & ",'" & billNoTemp & "','" & invTemp & "','Y','" & username & "',sysdate,'" & shipTemp & "','" & strTmp(0) & "','" & strTmp(1) & "','" & strTmp(2) & "','" & strTmp(3) & "','" & strTmp(4) & "') "
cmdStr2 = " insert into [erpdata].[dbo].[TSVtblMPSInvoice] (id,���ݱ��,InvoiceNo,CreateDate,shipno,ZX_REQUESTNUMBER,ZX_SHIPPER,ZX_SELFPICKUP,ZX_TRACKINGNO,ZX_SPECIALMARK) values " & " (" & idTemp & ",'" & billNoTemp & "','" & invTemp & "',getdate(),'" & shipTemp & "','" & strTmp(0) & "','" & strTmp(1) & "','" & strTmp(2) & "','" & strTmp(3) & "','" & strTmp(4) & "') "
AddSql (cmdStr)
AddSql2 (cmdStr2)
'Cnn.CommitTrans
MsgBox "����ɹ�!", vbInformation, "������ʾ"
ShowData_Where

End Sub

Private Sub Command4_Click()
TxtOpn.Text = ""
TxtInvoice.Text = ""
TxtShiping.Text = ""

TxtOpn.SetFocus

End Sub

Private Sub Form_Activate()
TxtOpn.SetFocus

End Sub

Private Sub Form_Load()


'  Get_ON_MShortName  ���̿��ϱ�ʶ��


'DTPicker1.Value = DateTime.Date - 7
'DTPicker2.Value = DateTime.Date

'IniCustomerName
IniFpsHeader



'DTPicker1.MultiSelect = True
'DTPicker2.MultiSelect = True
'DTPicker3.MultiSelect = True


'DTPicker1.Value = Null
'DTPicker2.Value = Null
'DTPicker3.Value = Null

'ShowData_Where

'�����û���,���Ƿ����޸ĵ�Ȩ��

'Call UserType(UCase(gUserName))

ShowData_Where

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

Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CmbCustomer.RowSource = mainItemRS
CmbCustomer.ListField = mainItemRS("productname").Name
CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub


Private Sub ShowData_Where()

Dim DT1 As Date
Dim DT2 As Date
Dim lotIdTemp As String

Set reportRS = GetONMarkingCode()

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
        .SetText E_FPS0.E_PART, 0, "��������"
        .SetText E_FPS0.E_WSGPART, 0, "Invoice Number"
        .SetText E_FPS0.E_CODE, 0, "Shipping Number"
        .SetText E_FPS0.E_REQUESTNUMBER, 0, "ZX_REQUESTNUMBER"
        .SetText E_FPS0.E_SHIPPER, 0, "ZX_SHIPPER"
        .SetText E_FPS0.E_SELFPICKUP, 0, "ZX_SELFPICKUP"
        .SetText E_FPS0.E_TRACKINGNO, 0, "ZX_TRACKINGNO"
        .SetText E_FPS0.E_SPECIALMARK, 0, "ZX_SPECIALMARK"
  
        
    
        .ColWidth(E_FPS0.E_SeqId) = 6
        .ColWidth(E_FPS0.E_PART) = 20
        .ColWidth(E_FPS0.E_WSGPART) = 20
        .ColWidth(E_FPS0.E_CODE) = 20
        .ColWidth(E_FPS0.E_REQUESTNUMBER) = 15
        .ColWidth(E_FPS0.E_SHIPPER) = 15
        .ColWidth(E_FPS0.E_SELFPICKUP) = 15
        .ColWidth(E_FPS0.E_TRACKINGNO) = 15
        .ColWidth(E_FPS0.E_SPECIALMARK) = 15
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15

        
        .ReDraw = True
    End With
    
    With cmb
        .Clear
        .AddItem ""
        .AddItem "Y"
        .AddItem "N"
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
'ShowData (i)
'
'End Sub

'Private Sub ShowData(i As Long)
'
'Set reportRS = GetNPIDataID(i)
'
'
' If reportRS.RecordCount > 0 Then
'
'
'    CmbCustomer.Text = reportRS.fields("CustomershortName").Value & ""
'    TxtQtechPT.Text = reportRS.fields("QtechPTNo").Value & ""
'    TxtQtechPT2.Text = reportRS.fields("QtechPTNo2").Value & ""
'    TxtCustPT1.Text = reportRS.fields("CustomerPTNo1").Value & ""
'    TxtCustPT2.Text = reportRS.fields("CustomerPTNo2").Value & ""
'    TxtCustDie.Text = reportRS.fields("CustomerDieQty").Value & ""
'    TxtQtechDie.Text = reportRS.fields("QtechDieQty").Value & ""
'    TxtXS.Text = reportRS.fields("XiangSu").Value & ""
'    TxtArea.Text = reportRS.fields("UsedArea").Value & ""
'    TxtStr1.Text = reportRS.fields("StruckStr1").Value & ""
'    TxtStr2.Text = reportRS.fields("StruckStr2").Value & ""
'    TxtStr3.Text = reportRS.fields("StruckStr3").Value & ""
'    DTPicker1.Value = reportRS.fields("ST_DATE").Value
'    DTPicker2.Value = reportRS.fields("TT_DATE").Value
'    DTPicker3.Value = reportRS.fields("PT_DATE").Value
'
'    TxtIDTemp.Text = reportRS.fields("ID").Value
' End If
'
'
'
'End Sub

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
TxtCustDie.SetFocus
End If

End Sub

Private Sub TxtCustDie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtQtechDie.SetFocus
End If

End Sub

Private Sub TxtQtechDie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtXS.SetFocus
End If

End Sub


Private Sub TxtXS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtArea.SetFocus
End If

End Sub

Private Sub TxtArea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtStr1.SetFocus
End If

End Sub

Private Sub TxtStr1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtStr2.SetFocus
End If

End Sub

Private Sub TxtStr2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtStr3.SetFocus
End If

End Sub



Private Sub TxtWaferID_KeyPress(KeyAscii As Integer)
Dim WaferID As String

If KeyAscii = 13 Then
WaferID = UCase(Trim(txtWaferID.Text))


 Set reportRS = GetBaoFeiOIData(WaferID)
    If (reportRS.RecordCount > 0) Then
    
    txtLotID.Text = reportRS.Fields("lotid").Value
    TxtGDDie.Text = reportRS.Fields("passbincount").Value
    TxtNGDie.Text = reportRS.Fields("failbincount").Value
    
    TxtCustPT.Text = reportRS.Fields("design_id").Value
    TxtHTPT.Text = reportRS.Fields("alternatename").Value
    TxtProduct.Text = reportRS.Fields("product").Value
    
    Else
    
      MsgBox "��ȷ��WaferID�Ƿ���ȷ��"
    End If




End If

End Sub

Private Sub TxtOpn_KeyPress(KeyAscii As Integer)
'�س���

Dim billNoTemp As String
Dim dateTemp As String
Dim notemp As Long
Dim noStrTemp As String


If KeyAscii = 13 Then



billNoTemp = UCase(Trim(TxtOpn.Text))


'�ж���ʷ������ţ��Ƿ���ȷ
If Not JudgeMPSBillNo(billNoTemp) Then

 MsgBox "��ȷ�Ϸ��������Ƿ���ȷ��ֿ���û�е㷢��!", vbInformation, "������ʾ"
 
 Exit Sub

End If




dateTemp = GetMPS_OutDate(billNoTemp)

notemp = GetMPSBillID()
noStrTemp = Right("00" & notemp, 3)

invoiceTemp = "HT-68-" & Format(dateTemp, "YYMMDD") & noStrTemp & "A"


TxtInvoice.Text = invoiceTemp

TxtShiping.SetFocus


End If


End Sub
