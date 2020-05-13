VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmMDPriceCreate 
   Caption         =   "市场部订单价格维护"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17535
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
   ScaleWidth      =   17535
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRate 
      Height          =   375
      Left            =   10080
      TabIndex        =   33
      Top             =   2070
      Width           =   2535
   End
   Begin VB.TextBox txtSingleDie 
      Height          =   375
      Left            =   6360
      TabIndex        =   30
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtBJ 
      Height          =   405
      Left            =   13680
      TabIndex        =   29
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtCustAA 
      Height          =   405
      Left            =   9480
      TabIndex        =   27
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox TxtPease 
      Height          =   375
      Left            =   9480
      TabIndex        =   24
      Top             =   780
      Width           =   2535
   End
   Begin VB.TextBox TxtFile 
      Height          =   375
      Left            =   5520
      TabIndex        =   22
      Top             =   1440
      Width           =   2535
   End
   Begin VB.ComboBox CmbPUnit 
      Height          =   315
      ItemData        =   "FrmMDPriceCreate.frx":0000
      Left            =   9480
      List            =   "FrmMDPriceCreate.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1560
      Width           =   2535
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      ItemData        =   "FrmMDPriceCreate.frx":001C
      Left            =   5520
      List            =   "FrmMDPriceCreate.frx":002C
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "退出"
      Height          =   360
      Left            =   9240
      TabIndex        =   12
      Top             =   3000
      Width           =   990
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "清空"
      Height          =   360
      Left            =   6840
      TabIndex        =   11
      Top             =   3000
      Width           =   990
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "保存"
      Height          =   360
      Left            =   4320
      TabIndex        =   10
      Top             =   3000
      Width           =   990
   End
   Begin VB.TextBox TxtPrice 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2040
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
      Left            =   1680
      TabIndex        =   5
      Top             =   1440
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
      Height          =   5775
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   3960
      Width           =   17055
      _Version        =   524288
      _ExtentX        =   30083
      _ExtentY        =   10186
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
      AppearanceStyle =   0
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
      Left            =   1680
      TabIndex        =   16
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "汇率:"
      Height          =   195
      Left            =   9600
      TabIndex        =   32
      Top             =   2160
      Width           =   420
   End
   Begin VB.Label lblDIE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单DIE价格："
      Height          =   195
      Left            =   5280
      TabIndex        =   31
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblBJ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "报价单号："
      Height          =   195
      Left            =   12600
      TabIndex        =   28
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label lblcustAA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户简写："
      Height          =   195
      Left            =   8520
      TabIndex        =   26
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "订单数量(片)："
      Height          =   195
      Left            =   8280
      TabIndex        =   25
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "返点文件名："
      Height          =   195
      Left            =   4440
      TabIndex        =   23
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单价单位："
      Height          =   195
      Left            =   8520
      TabIndex        =   20
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单片价格："
      Height          =   195
      Left            =   720
      TabIndex        =   19
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "订单号："
      Height          =   195
      Left            =   720
      TabIndex        =   17
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户代码："
      Height          =   195
      Left            =   600
      TabIndex        =   15
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "订单数量(Die)："
      Height          =   195
      Left            =   12360
      TabIndex        =   8
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "机种："
      Height          =   195
      Left            =   960
      TabIndex        =   6
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "订单类型："
      Height          =   195
      Left            =   4560
      TabIndex        =   4
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "提供订单日期："
      Height          =   195
      Left            =   12360
      TabIndex        =   2
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户全称："
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
Private Enum E_FPS0          'Detail
    E_SeqId = 1                '序号
    E_CustID               '客户代码
    E_CUSTNAME             '客户名称
    e_PO                   'PO号
    E_PODATE               '上传日期
    E_POTYPE               'PO类型
    E_PT                   '机种
    E_PeaceQty                   '数量
    E_QTY                   '数量
    e_Price                 '单价
    E_PUnit                '单价单元
    E_File                 '文件明
    E_BJ                    '报价单号
    E_CustAA                '客户简称

    
    E_END
    
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

'校验是否重复

If UCase(Trim(CmbCustomer.Text)) = "" Or UCase(Trim(TxtQtechPT.Text)) = "" Then
     MsgBox "客户代码或厂内项目名称不可以为空！"
     Exit Sub
End If

If UCase(Trim(TxtCustPT1.Text)) = "" And UCase(Trim(TxtCustPT2.Text)) = "" Then
     MsgBox "客户机种不可以为空！"
     Exit Sub
End If


' Set bomRS2 = GetNpiProductCheck(UCase(Trim(CmbCustomer.Text)), UCase(Trim(TxtQtechPT.Text)), UCase(Trim(TxtCustPT1.Text)), UCase(Trim(TxtCustPT2.Text)), UCase(Trim(TxtQtechPT2.Text)))
'If bomRS2.RecordCount > 0 Then
'    MsgBox "系统中已存在这笔数据，请重新确认输入是否正确 ！"
'    Exit Sub
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

 MsgBox "新增成功!", vbInformation, "友情提示"
 
 ShowData_Where

End Sub

Private Sub CmbCustomer_Change()
Dim customerTemp As String
Dim unit As String

customerTemp = CmbCustomer.Text

txtCusName.Text = GetCustomerNameSqlServer(customerTemp)
txtCustAA.Text = GetCustomerNameSqlServer1(customerTemp)
unit = GetCustomerNameSqlServer2(customerTemp)
If unit = "01" Then
CmbPUnit.Text = "人民币"
Else
CmbPUnit.Text = "美元"
End If

'根据客户代码，初始化订单

IniCustomerPOName (customerTemp)



End Sub

Private Sub CmbPONum_Change()
 Dim potemp As String
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
        TxtPease.Text = oiRS.Fields("qty2").Value
        
        
     End If
     


End Sub



Private Sub cmdDel_Click()
txtCusName.Text = ""
TxtCreDate.Text = ""
TxtPT.Text = ""
txtQty.Text = ""
TxtPease.Text = ""

TxtPrice.Text = ""
TxtFile.Text = ""



End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdModify_Click()
'添加
Dim nPOTemp As POPrice
Dim userid As String

userid = UCase(gUserName)

nPOTemp.CreateBy = UCase(gUserName)

If UCase(Trim(CmbCustomer.Text)) = "" Or UCase(Trim(txtCusName.Text)) = "" Then
     MsgBox "客户代码或客户名称不可以为空！"
     Exit Sub
End If


If UCase(Trim(txtQty.Text)) = "" Or (UCase(Trim(TxtPrice.Text)) = "" And UCase(Trim(txtSingleDie.Text)) = "") Or UCase(Trim(CmbPUnit.Text)) = "" Then
     MsgBox "数量，单价，单位 不可以为空！"
     Exit Sub
End If


If UCase(Trim(CmbCustomer.Text)) = "KR001" Or UCase(Trim(CmbCustomer.Text)) = "81" Then
    If UCase(Trim(txtBJ.Text)) = "" Then
        MsgBox "报价单号不能为空！"
        Exit Sub
 
    End If
End If

nPOTemp.CreateBy = UCase(gUserName)
nPOTemp.ID = GetPOPriceID()
nPOTemp.customerName = UCase(Trim(txtCusName.Text))
nPOTemp.CUSTOMERSHORTNAME = UCase(Trim(CmbCustomer.Text))
nPOTemp.PONo = UCase(Trim(CmbPONum.Text))
nPOTemp.PODATE = UCase(Trim(TxtCreDate.Text))
nPOTemp.POType = UCase(Trim(CmbType.Text))
nPOTemp.pt = UCase(Trim(TxtPT.Text))
nPOTemp.QTY = UCase(Trim(txtQty.Text))
nPOTemp.peaseQty = UCase(Trim(TxtPease.Text))

nPOTemp.Price = UCase(Trim(TxtPrice.Text))
nPOTemp.unit = UCase(Trim(CmbPUnit.Text))
nPOTemp.File = UCase(Trim(TxtFile.Text))
nPOTemp.bj = UCase(Trim(txtBJ.Text))
nPOTemp.custAA = UCase(Trim(txtCustAA))
nPOTemp.SingDie = Trim$(txtSingleDie.Text)
nPOTemp.SingWafer = Trim$(txtRate.Text)

Set PJPO = GetPJPOData(UCase(Trim(CmbPONum.Text)), UCase(Trim(TxtPT.Text)))
If PJPO.RecordCount > 0 Then
    MsgBox "Mes系统中已存在此采购单单号，请确认采购单 ！"
    ComSave.Enabled = True
    Exit Sub
End If

Call AddPOPrice(nPOTemp)

MsgBox "新增成功!", vbInformation, "友情提示"
 
 ShowData_Where
 
End Sub

Private Sub CmdOutReport_Click()
Dim sqlTemp As String

sqlTemp = "select  id  , CUSTOMERSHORTNAME as 客户代码 , QtechPTNo as 厂内项目名称 ,QtechPTNo2 as 成品料号, CUSTOMERPTNo1  as 客户机种名1, CUSTOMERPTNo2 as 客户机种名2 , " & _
         " CUSTOMERDieQty as 客户设计die数, XiangSu  as 像素,  fzFreeUSD as 封装费USD,testFreeUSD as 测试费USD,fzFreeRMB as 封装费RMB,testFreeRMB as 测试费RMB,nreFree as NRE费用YN开票日期,nreMethod as NRE返还方式,updatePrice2 as 调价历史2,updatePrice1 as 调价历史1 " & _
         " From TBLTsvNpiProduct where flag='Y' order by CUSTOMERSHORTNAME,QtechPTNo,CUSTOMERPTNo1,CUSTOMERPTNo2 "
         
  ExporToExcel (sqlTemp)

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

'Private Sub IniCustomerName()
'Set mainItemRS = GetJDCustomerName()
'Set CmbCustomer.RowSource = mainItemRS
'CmbCustomer.ListField = mainItemRS("productname").Name
'CmbCustomer.BoundColumn = mainItemRS("PID").Name
'
'End Sub


Private Sub ShowData_Where()
Set reportRS = GetPOPrice()

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
        .SetText E_FPS0.E_CustID, 0, "客户代码"
        .SetText E_FPS0.E_CUSTNAME, 0, "客户全称"
        .SetText E_FPS0.e_PO, 0, "订单号"
        .SetText E_FPS0.E_PODATE, 0, "提供订单日期"
        .SetText E_FPS0.E_POTYPE, 0, "订单类型"
        .SetText E_FPS0.E_PT, 0, "机种"
        
        .SetText E_FPS0.E_PeaceQty, 0, "订单片数"
        .SetText E_FPS0.E_QTY, 0, "订单数量"
        
        .SetText E_FPS0.e_Price, 0, "单价"
        .SetText E_FPS0.E_PUnit, 0, "单价单位"
        .SetText E_FPS0.E_File, 0, "返点文件名"
        
        .SetText E_FPS0.E_CustAA, 0, "客户简称"
        .SetText E_FPS0.E_BJ, 0, "报价单号"
       
        .ColWidth(E_FPS0.E_SeqId) = 5
        .ColWidth(E_FPS0.E_CustID) = 8
        .ColWidth(E_FPS0.E_CUSTNAME) = 25
        .ColWidth(E_FPS0.e_PO) = 15
        .ColWidth(E_FPS0.E_PODATE) = 10
        .ColWidth(E_FPS0.E_POTYPE) = 8
        .ColWidth(E_FPS0.E_PT) = 10
        .ColWidth(E_FPS0.E_PeaceQty) = 9
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



Private Sub Fps_DBLClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim I As Long

With fpS(0)
            .Row = Row
            .Col = 1
       I = .Text

End With

ShowData (I)

Txtfzu.SetFocus

End Sub

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

'
'Private Sub txtBJ_LostFocus()
'
'Dim cmdSql As String
'Dim BJNO As String
'
'BJNO = Trim(txtBJ.Text)
'
'If JudgeBJNO(BJNO) = False Then
'    MsgBox "您输入的报价单号有误， 请确认"
'    Exit Sub
'End If
'
'Set mainItemRS = GetSingelPrice(BJNO)
'
'txtTextSinglePrice.Text = mainItemRS("WAFER_PRICE").Value
'txtSingleDie.Text = mainItemRS("DIE_PRICE").Value
'
'End Sub

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


