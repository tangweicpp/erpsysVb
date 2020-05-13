VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmOnMSLevel 
   Caption         =   "MSLevel"
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
   Begin VB.CommandButton Command2 
      Caption         =   "导出"
      Height          =   360
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   360
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   990
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   8655
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   18735
      _Version        =   196613
      _ExtentX        =   33046
      _ExtentY        =   15266
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
      SpreadDesigner  =   "FrmOnMSLevel.frx":0000
      TextTip         =   2
   End
End
Attribute VB_Name = "FrmOnMSLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
    E_SeqId = 1                '序号
    E_MSLevel
    E_Number              '录入部门
   
    E_End
    
End Enum

Dim reportRS As New ADODB.Recordset
Dim mainItemRS As New ADODB.Recordset
Dim bomRS2        As New ADODB.Recordset



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
 
 ShowData_Where

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

ShowData_Where

End Sub

Private Sub CmdOutReport_Click()
Dim sqlTemp As String

sqlTemp = "select  id ,PutInDept as 登记部门别, WaferID ,LotID ,GDDie,NGDie ,CUSTOMERPTNO as 客户机种号, QTECHPTNO as 厂内机种号,ProductName,ERR_DATE as 异常日期 ,PutIn_DATE as 录入日期 from TBLTSVBaoFei where flag='Y'  order by PutIn_DATE desc,lotid,waferid "
  
         
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
'If CmbType.Text = "审核1" Then
'typeId = 2
'ElseIf CmbType.Text = "是否通知客户" Then
'typeId = 3
'
'ElseIf CmbType.Text = "客户同意报废市场部确认" Then
'typeId = 4
'
'
'End If
'
'Call ShowData_Where(typeId)

End Sub

Private Sub Command1_Click()
Dim userId As String

ShowData_Where


End Sub

Private Sub Command2_Click()
Dim temp As String


  temp = "select id,ms_level,typename||numberofhours as name from  CUSTOMERMSLevelTBL where flag='Y' order by id"
      
      
 ExporToExcel (temp)
 
End Sub

Private Sub Form_Activate()
'CmbCustomer.SetFocus

End Sub

Private Sub Form_Load()

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

'根据用户名,看是否有修改的权限

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

Dim dt1 As Date
Dim dt2 As Date
Dim lotidtemp As String

Set reportRS = GetMSLevel()

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
        

        
        .SetText E_FPS0.E_SeqId, 0, "记录号"
        .SetText E_FPS0.E_MSLevel, 0, "MS-Level"
        .SetText E_FPS0.E_Number, 0, "Number of Hours"
       
         

        

        
        .ColWidth(E_FPS0.E_SeqId) = 6
        .ColWidth(E_FPS0.E_MSLevel) = 10
        .ColWidth(E_FPS0.E_Number) = 20
     
     
       

        .RowHeight(0) = 20
        .RowHeight(-1) = 15

        
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
