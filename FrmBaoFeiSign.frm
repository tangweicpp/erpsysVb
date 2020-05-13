VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmBaoFeiSign 
   Caption         =   "审核报废信息"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
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
   ScaleWidth      =   19080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtReject 
      Height          =   615
      Left            =   1080
      TabIndex        =   8
      Top             =   720
      Width           =   6255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmBaoFeiSign.frx":0000
      Left            =   1080
      List            =   "FrmBaoFeiSign.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "驳回"
      Height          =   360
      Left            =   7800
      TabIndex        =   5
      Top             =   960
      Width           =   990
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      ItemData        =   "FrmBaoFeiSign.frx":001A
      Left            =   5040
      List            =   "FrmBaoFeiSign.frx":0027
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.CheckBox ChkAll 
      Height          =   375
      Left            =   15480
      TabIndex        =   2
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "审核"
      Height          =   360
      Left            =   7800
      TabIndex        =   1
      Top             =   240
      Width           =   990
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   8175
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   19815
      _Version        =   196613
      _ExtentX        =   34951
      _ExtentY        =   14420
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
      SpreadDesigner  =   "FrmBaoFeiSign.frx":0058
      TextTip         =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "驳回原因："
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "类型别："
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "审核类型："
      Height          =   195
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "FrmBaoFeiSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
    E_SeqId = 1                '序号
    E_Type                     '类型
    E_PutInDept               '录入部门
    E_WaferId                  'WaferID
    E_LotId                    'LotID
    E_GDDie                    'Good Die
    E_NGDie                    'NG Die
    E_CustPT                '客户机种名
    E_HTPT                  '厂内机种名
    E_Product                '产品名
   
    E_Time1                '异常日期
    E_Time2                '录入日期
    E_OK                '
    
    
'      ID                    NUMBER,
'  PutInDept             VARCHAR2(10),
'  WaferID               VARCHAR2(20),
'  LotID                 VARCHAR2(20),
'  GDDie                 NUMBER,
'  NGDie                 NUMBER,
'  CUSTOMERPTNO        VARCHAR2(20),
'  QTECHPTNO           VARCHAR2(10),
'  ProductName         VARCHAR2(20),
'  ERR_DATE               VARCHAR2(50),
'  PutIn_DATE             VARCHAR2(50),
'
    
    
    
    E_End
    
End Enum

Dim reportRS As New ADODB.Recordset
Dim mainItemRS As New ADODB.Recordset
Dim bomRS2        As New ADODB.Recordset
Dim typeId As Integer


Private Sub CmbCustomer_Change()
TxtQtechPT.SetFocus
End Sub

Private Sub CmbDept_Click()
TxtWaferID.SetFocus
End Sub

Private Sub CmdAdd_Click()
Dim baoFeiTemp As Baofei
Dim userId As String

'校验是否重复

If UCase(Trim(CmbDept.Text)) = "" Then
     MsgBox "请选择部门别！"
     Exit Sub
End If

If UCase(Trim(TxtWaferID.Text)) = "" And UCase(Trim(TxtLotId.Text)) = "" Then
     MsgBox "请确认WaferID,LotID!"
     Exit Sub
End If


userId = UCase(gUserName)

baoFeiTemp.CreateBy = userId

baoFeiTemp.putInDept = CmbDept.Text
baoFeiTemp.err_date = DTPicker1.Value
baoFeiTemp.putIn_date = DTPicker2.Value
baoFeiTemp.waferid = TxtWaferID.Text
baoFeiTemp.lotid = TxtLotId.Text
baoFeiTemp.gDDie = CInt(TxtGDDie.Text)
baoFeiTemp.nGDie = CInt(TxtNGDie.Text)
baoFeiTemp.customerPTNo = TxtCustPT.Text
baoFeiTemp.qtechPTNo = TxtHTPT.Text
baoFeiTemp.productName = TxtProduct.Text


Call AddBaoFei(baoFeiTemp)

 MsgBox "新增成功!", vbInformation, "友情提示"
 
 'ShowData_Where

End Sub

Private Sub CmdDel_Click()
TxtWaferID.Text = ""
TxtLotId.Text = ""
TxtGDDie.Text = ""
TxtNGDie.Text = ""
TxtCustPT.Text = ""
TxtHTPT.Text = ""
TxtProduct.Text = ""

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdModify_Click()
'修改

Dim nPIProductTemp As NpiProduct
Dim userId As String
userId = UCase(gUserName)

nPIProductTemp.CreateBy = UCase(gUserName)
nPIProductTemp.CustomershortName = UCase(Trim(CmbCustomer.Text))
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

 MsgBox "修改成功!", vbInformation, "友情提示"

'ShowData_Where

End Sub

Private Sub CmdOutReport_Click()
Dim sqlTemp As String

sqlTemp = "select  id ,PutInDept as 登记部门别, WaferID ,LotID ,GDDie,NGDie ,CUSTOMERPTNO as 客户机种号, QTECHPTNO as 厂内机种号,ProductName,ERR_DATE as 异常日期 ,PutIn_DATE as 录入日期 from TBLTSVBaoFei where flag='Y'  order by PutIn_DATE desc,lotid,waferid "
  
         
  ExporToExcel (sqlTemp)

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

Private Sub CmbType_Click()
If UCase(Trim(Combo1.Text)) = "" Then
     MsgBox "请先选择类型！"
     Exit Sub
End If

If CmbType.Text = "审核1" Then
typeId = 2
ElseIf CmbType.Text = "是否通知客户" Then
typeId = 3

ElseIf CmbType.Text = "客户同意报废市场部确认" Then
typeId = 4


End If

Call ShowData_Where(typeId)

End Sub

Private Sub Command1_Click()
Dim userId As String


userId = UCase(gUserName)



If UCase(Trim(CmbType.Text)) = "" Then
     MsgBox "请选择审核类别！"
     Exit Sub
End If

If CmbType.Text = "审核1" Then
typeId = 2
ElseIf CmbType.Text = "是否通知客户" Then
typeId = 3
End If

Dim i As Integer
Dim signCount As Integer
Dim EmpIDTemp As Long

signCount = 0

    For i = 1 To fps(0).MaxRows
        With fps(0)
            .Row = i
            .Col = E_FPS0.E_OK
            If .Text = "1" Then
                .Row = i
                .Col = E_FPS0.E_SeqId
                EmpIDTemp = CLng(.Text)
                SignBaoFeiBill EmpIDTemp, userId, typeId, Combo1.Text
                signCount = signCount + 1
            End If
        End With
    Next i
    
   MsgBox "审核成功!", vbInformation, "友情提示"
    
ShowData_Where typeId


End Sub

Private Sub Command2_Click()

Dim userId As String


userId = UCase(gUserName)



If UCase(Trim(CmbType.Text)) = "" Then
     MsgBox "请选择审核类别！"
     Exit Sub
End If

'If CmbType.Text = "审核1" Then
'typeId = 2
'ElseIf CmbType.Text = "是否通知客户" Then
'typeId = 3
'End If

Dim i As Integer
Dim signCount As Integer
Dim EmpIDTemp As Long

signCount = 0

    For i = 1 To fps(0).MaxRows
        With fps(0)
            .Row = i
            .Col = E_FPS0.E_OK
            If .Text = "1" Then
                .Row = i
                .Col = E_FPS0.E_SeqId
                EmpIDTemp = CLng(.Text)
                RejectBaoFeiBill EmpIDTemp, userId, typeId, Combo1.Text, TxtReject.Text
                signCount = signCount + 1
            End If
        End With
    Next i

   MsgBox "退回成功!", vbInformation, "友情提示"

ShowData_Where typeId



End Sub

Private Sub Form_Activate()
'CmbCustomer.SetFocus

End Sub

Private Sub Form_Load()

'IniCustomerName
IniFpsHeader



'DTPicker1.MultiSelect = True
'DTPicker2.MultiSelect = True
'DTPicker3.MultiSelect = True


'DTPicker1.Value = Null
'DTPicker2.Value = Null
'DTPicker3.Value = Null

'ShowData_Where

'根据用户名,看是否有修改的权限

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


Private Sub ShowData_Where(typeId As Integer)
Set reportRS = GetBaoFeiDataSign(typeId, Combo1.Text)

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
        
        '砞竚Α
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
        
        
        .SetText E_FPS0.E_SeqId, 0, "记录号"
        .SetText E_FPS0.E_Type, 0, "登记类型"
        .SetText E_FPS0.E_PutInDept, 0, "登记部门别"
        .SetText E_FPS0.E_WaferId, 0, "WaferID"
        .SetText E_FPS0.E_LotId, 0, "LotID"
        .SetText E_FPS0.E_GDDie, 0, "GoodDie"
        .SetText E_FPS0.E_NGDie, 0, "NGDie"
        .SetText E_FPS0.E_CustPT, 0, "客户机种名"
        .SetText E_FPS0.E_HTPT, 0, "厂内机种名"
        .SetText E_FPS0.E_Product, 0, "成品料号"
        
        .SetText E_FPS0.E_Time1, 0, "异常日期"
        .SetText E_FPS0.E_Time2, 0, "录入日期"
        .SetText E_FPS0.E_Time2, 0, "录入日期"
        
        .SetText E_FPS0.E_OK, 0, ""
  
        
        .ColWidth(E_FPS0.E_SeqId) = 10
        .ColWidth(E_FPS0.E_Type) = 10
        .ColWidth(E_FPS0.E_PutInDept) = 10
        .ColWidth(E_FPS0.E_WaferId) = 10
        .ColWidth(E_FPS0.E_LotId) = 10
        .ColWidth(E_FPS0.E_GDDie) = 10
        .ColWidth(E_FPS0.E_NGDie) = 10
        .ColWidth(E_FPS0.E_CustPT) = 10
        .ColWidth(E_FPS0.E_HTPT) = 10
        .ColWidth(E_FPS0.E_Time1) = 10
        .ColWidth(E_FPS0.E_Time1) = 10
        .ColWidth(E_FPS0.E_Time2) = 10
         .ColWidth(E_FPS0.E_OK) = 10
       

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        .Col = E_FPS0.E_OK
        .Lock = False
        
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

ShowData (i)

End Sub

Private Sub ShowData(i As Long)

Set reportRS = GetNPIDataID(i)


 If reportRS.RecordCount > 0 Then
 
 
    CmbCustomer.Text = reportRS.fields("CustomershortName").Value & ""
    TxtQtechPT.Text = reportRS.fields("QtechPTNo").Value & ""
    TxtQtechPT2.Text = reportRS.fields("QtechPTNo2").Value & ""
    TxtCustPT1.Text = reportRS.fields("CustomerPTNo1").Value & ""
    TxtCustPT2.Text = reportRS.fields("CustomerPTNo2").Value & ""
    TxtCustDie.Text = reportRS.fields("CustomerDieQty").Value & ""
    TxtQtechDie.Text = reportRS.fields("QtechDieQty").Value & ""
    TxtXS.Text = reportRS.fields("XiangSu").Value & ""
    TxtArea.Text = reportRS.fields("UsedArea").Value & ""
    TxtStr1.Text = reportRS.fields("StruckStr1").Value & ""
    TxtStr2.Text = reportRS.fields("StruckStr2").Value & ""
    TxtStr3.Text = reportRS.fields("StruckStr3").Value & ""
    DTPicker1.Value = reportRS.fields("ST_DATE").Value
    DTPicker2.Value = reportRS.fields("TT_DATE").Value
    DTPicker3.Value = reportRS.fields("PT_DATE").Value
    
    TxtIDTemp.Text = reportRS.fields("ID").Value
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
Dim waferid As String

If KeyAscii = 13 Then
waferid = UCase(Trim(TxtWaferID.Text))


 Set reportRS = GetBaoFeiOIData(waferid)
    If (reportRS.RecordCount > 0) Then
    
    TxtLotId.Text = reportRS.fields("lotid").Value
    TxtGDDie.Text = reportRS.fields("passbincount").Value
    TxtNGDie.Text = reportRS.fields("failbincount").Value
    
    TxtCustPT.Text = reportRS.fields("design_id").Value
    TxtHTPT.Text = reportRS.fields("alternatename").Value
    TxtProduct.Text = reportRS.fields("product").Value
    
    Else
    
      MsgBox "请确认WaferID是否正确！"
    End If




End If

End Sub
