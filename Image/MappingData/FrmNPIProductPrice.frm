VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmNPIProductPrice 
   Caption         =   "�г���NPI��Ʒ�۸�ά��"
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
   Begin VB.CommandButton CmdOutReport 
      Caption         =   "��������"
      Height          =   360
      Left            =   11760
      TabIndex        =   19
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "�˳�"
      Height          =   360
      Left            =   9180
      TabIndex        =   18
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "���"
      Height          =   360
      Left            =   6600
      TabIndex        =   17
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "�޸�"
      Height          =   360
      Left            =   4020
      TabIndex        =   16
      Top             =   1560
      Width           =   990
   End
   Begin VB.TextBox TxtHis1 
      Height          =   375
      Left            =   15120
      TabIndex        =   15
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox TxtHis2 
      Height          =   375
      Left            =   10920
      TabIndex        =   13
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox TxtNreW 
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox TxtNreF 
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox TxtTestR 
      Height          =   375
      Left            =   15120
      TabIndex        =   7
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox Txtfzr 
      Height          =   375
      Left            =   10920
      TabIndex        =   5
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox TxtTestu 
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox Txtfzu 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   7455
      Index           =   0
      Left            =   360
      TabIndex        =   20
      Top             =   2160
      Width           =   19815
      _Version        =   196613
      _ExtentX        =   34951
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
      SpreadDesigner  =   "FrmNPIProductPrice.frx":0000
      TextTip         =   2
   End
   Begin VB.Label TxtIDTemp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1200
      TabIndex        =   22
      Top             =   1440
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��¼��:"
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ʷ1��"
      Height          =   195
      Left            =   14040
      TabIndex        =   14
      Top             =   960
      Width           =   990
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ʷ2��"
      Height          =   195
      Left            =   9840
      TabIndex        =   12
      Top             =   960
      Width           =   990
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NRE������ʽ��"
      Height          =   195
      Left            =   5160
      TabIndex        =   10
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NRE����(Y/N)&��Ʊ���ڣ�"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���Է�(RMB)��"
      Height          =   195
      Left            =   13920
      TabIndex        =   6
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��װ��(RMB)��"
      Height          =   195
      Left            =   9720
      TabIndex        =   4
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���Է�(USD)��"
      Height          =   195
      Left            =   5280
      TabIndex        =   2
      Top             =   360
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��װ��(USD)��"
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   1140
   End
End
Attribute VB_Name = "FrmNPIProductPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail�֭��
    E_SeqId = 1                '���
    E_CustName               '�ͻ�����
    E_QtechPT                '������Ŀ����
    E_QtechPT2                '��Ʒ�Ϻ�
    E_CustPT1                '�ͻ�������1
    E_CustPT2                '�ͻ�������2
    E_CustDie                '�ͻ����die��
    E_XS                   '����
    
    E_FZU                   '��װ��USD
    E_TestU                '���Է�USD
    E_FZR                  '��װ��RMB
    E_TestR                '���Է�RMB
    
    E_NreF                 'NRE����
    E_NreW                 'NRE������ʽ
    E_UP2                  '������ʷ2
    E_UP1                  '������ʷ1
    
    E_End
    
End Enum

Dim reportRS As New ADODB.Recordset
Dim mainItemRS As New ADODB.Recordset
Dim bomRS2        As New ADODB.Recordset

Private Sub CmbCustomer_Change()
TxtQtechPT.SetFocus
End Sub

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

Private Sub CmdDel_Click()
Txtfzu.Text = ""
TxtTestu.Text = ""
Txtfzr.Text = ""
TxtTestR.Text = ""
TxtNreF.Text = ""
TxtNreW.Text = ""
TxtHis2.Text = ""
TxtHis1.Text = ""


End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdModify_Click()
'�޸�

Dim nPIProductTemp As NpiProduct
Dim userId As String
userId = UCase(gUserName)

nPIProductTemp.CreateBy = UCase(gUserName)
nPIProductTemp.FzFreeUSD = UCase(Trim(Txtfzu.Text))
nPIProductTemp.TestFreeUSD = UCase(Trim(TxtTestu.Text))
nPIProductTemp.FzFreeRMB = UCase(Trim(Txtfzr.Text))
nPIProductTemp.TestFreeRMB = UCase(Trim(TxtTestR.Text))
nPIProductTemp.NreFree = UCase(Trim(TxtNreF.Text))
nPIProductTemp.NreMethod = UCase(Trim(TxtNreW.Text))
nPIProductTemp.UpdatePrice2 = UCase(Trim(TxtHis2.Text))
nPIProductTemp.UpdatePrice1 = UCase(Trim(TxtHis1.Text))
    


Call ModifyNpiProductPrice(nPIProductTemp, CLng(TxtIDTemp.Caption))

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

Private Sub Form_Activate()
'CmbCustomer.SetFocus

End Sub

Private Sub Form_Load()

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

ShowData_Where

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

Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CmbCustomer.RowSource = mainItemRS
CmbCustomer.ListField = mainItemRS("productname").Name
CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub


Private Sub ShowData_Where()
Set reportRS = GetNPIDataPrice()

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
        .SetText E_FPS0.E_CustName, 0, "�ͻ�����"
        .SetText E_FPS0.E_QtechPT, 0, "������Ŀ����"
        .SetText E_FPS0.E_QtechPT2, 0, "��Ʒ�Ϻ�"
        .SetText E_FPS0.E_CustPT1, 0, "�ͻ�������1"
        .SetText E_FPS0.E_CustPT2, 0, "�ͻ�������2"
        .SetText E_FPS0.E_CustDie, 0, "�ͻ����die��"
        .SetText E_FPS0.E_XS, 0, "����"
        
        .SetText E_FPS0.E_FZU, 0, "��װ��(USD)"
        .SetText E_FPS0.E_TestU, 0, "���Է�(USD)"
        .SetText E_FPS0.E_FZR, 0, "��װ��(RMB)"
        .SetText E_FPS0.E_TestR, 0, "���Է�(RMB)"
        
        .SetText E_FPS0.E_NreF, 0, "NRE����(Y/N)&��Ʊ����"
        .SetText E_FPS0.E_NreW, 0, "NRE������ʽ"
        .SetText E_FPS0.E_UP2, 0, "������ʷ2"
        .SetText E_FPS0.E_UP1, 0, "������ʷ1"
        
        .ColWidth(E_FPS0.E_SeqId) = 5
        .ColWidth(E_FPS0.E_CustName) = 6
        .ColWidth(E_FPS0.E_QtechPT) = 10
        .ColWidth(E_FPS0.E_QtechPT2) = 12
        .ColWidth(E_FPS0.E_CustPT1) = 10
        .ColWidth(E_FPS0.E_CustPT2) = 10
        .ColWidth(E_FPS0.E_CustDie) = 10
        .ColWidth(E_FPS0.E_XS) = 10
        
        
        .ColWidth(E_FPS0.E_FZU) = 10
        .ColWidth(E_FPS0.E_TestU) = 12
        .ColWidth(E_FPS0.E_FZR) = 12
        .ColWidth(E_FPS0.E_TestR) = 12
        
        .ColWidth(E_FPS0.E_NreF) = 10
        .ColWidth(E_FPS0.E_NreW) = 10
        .ColWidth(E_FPS0.E_UP2) = 10
        .ColWidth(E_FPS0.E_UP1) = 10
        

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



