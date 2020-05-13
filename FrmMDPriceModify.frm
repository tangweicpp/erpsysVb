VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmMDPriceModify 
   Caption         =   "市场部订单价格修改"
   ClientHeight    =   9885
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9885
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSH 
      Caption         =   "审核"
      Height          =   360
      Left            =   4680
      TabIndex        =   8
      Top             =   960
      Width           =   990
   End
   Begin VB.ComboBox Comflag 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   2175
   End
   Begin VB.CheckBox ChkAll 
      Height          =   255
      Left            =   15120
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "查询"
      Height          =   360
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton CmdModify 
      BackColor       =   &H000000FF&
      Caption         =   "修改"
      Height          =   360
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   990
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   7695
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   17055
      _Version        =   524288
      _ExtentX        =   30083
      _ExtentY        =   13573
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
      SpreadDesigner  =   "FrmMDPriceModify.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin MSDataListLib.DataCombo CmbCustomer 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label lbl11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "审核状态："
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   900
   End
   Begin VB.Label lbl11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户代码："
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "FrmMDPriceModify"
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
    E_PeaceQty             '片数
    E_QTY                   '数量
    E_BJ                   '报价单号 By Tony 20180814
    e_Price                 '片单价
    E_DIEPrice             'DIE单价
    E_PUnit                '单价单元
    E_File                 '文件明
    E_OK                     '选择

    
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

Private Sub ChkAll_Click()
Dim I As Integer
    If ChkAll.Value = 1 Then
        For I = 1 To fpS(0).MaxRows
            With fpS(0)
                .Row = I
                .Col = E_FPS0.E_OK
                .Text = 1
            End With
        Next I
        
    ElseIf ChkAll.Value = 0 Then
        For I = 1 To fpS(0).MaxRows
            With fpS(0)
                .Row = I
                .Col = E_FPS0.E_OK
                .Text = 0
            End With
        Next I
        
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
''根据客户代码，初始化订单
'
'IniCustomerPOName (customerTemp)
'
'
'
'End Sub

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
        
     End If
     


End Sub



Private Sub cmdDel_Click()
'查询


Call ShowData_WhereCus(UCase(Trim(CmbCustomer.Text)), UCase(Trim(Comflag.Text)))

If UCase(Trim(Comflag.Text)) = "Y" Then

cmdSH.Enabled = False

Else

cmdSH.Enabled = True
    
End If


End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdModify_Click()
'修改

Dim nPOTemp As POPrice
Dim userid As String
userid = UCase(gUserName)

nPOTemp.CreateBy = UCase(gUserName)




'-----------
With fpS(0)

For I = 1 To .MaxRows

    .Row = I
    .Col = 15
    If .Text = 1 Then
    
    '要修改

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
    nPOTemp.peaseQty = .Text
    
        
    .Row = I
    .Col = 9
    nPOTemp.QTY = .Text
    
    .Row = I
    .Col = 10
    nPOTemp.bj = .Text
    
    .Row = I
    .Col = 11
    nPOTemp.Price = .Text
    
     .Row = I
    .Col = 12
    nPOTemp.DIE_PRICE = .Text
    
    .Row = I
    .Col = 13
    nPOTemp.unit = .Text
    
    .Row = I
    .Col = 14
    nPOTemp.File = .Text
    
    
'        If nPOTemp.bj = "" Then
'
'            MsgBox "报价单不能为空！"
'            Exit Sub
'
'        End If

    Call ModifyPOPrice(nPOTemp)

    
    
    End If
    

Next I


End With

'--------------------


 MsgBox "修改成功!", vbInformation, "友情提示"
 
 ShowData_Where
 
End Sub


Private Sub cmdSH_Click()

Dim nPOTemp As POPrice
Dim userid As String
userid = UCase(gUserName)

nPOTemp.CreateBy = UCase(gUserName)




'-----------
With fpS(0)

For I = 1 To .MaxRows

    .Row = I
    .Col = 15
    If .Text = 1 Then
    
    '要修改

    .Row = I
    .Col = 1
    nPOTemp.ID = .Text
    
     
    .Row = I
    .Col = 2
    nPOTemp.CUSTOMERSHORTNAME = .Text
    
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
    nPOTemp.peaseQty = .Text
    
        
    .Row = I
    .Col = 9
    nPOTemp.QTY = .Text
    
    .Row = I
    .Col = 10
    nPOTemp.bj = .Text
    
    .Row = I
    .Col = 11
    nPOTemp.Price = .Text
    
     .Row = I
    .Col = 12
    nPOTemp.DIE_PRICE = .Text
    
    .Row = I
    .Col = 13
    nPOTemp.unit = .Text
    
    .Row = I
    .Col = 14
    nPOTemp.File = .Text
    
    
'        If nPOTemp.bj = "" Then
'
'            MsgBox "报价单不能为空！"
'            Exit Sub
'
'        End If

    Call ApprovalPO(nPOTemp)

    
    
    End If
    

Next I


End With

'--------------------


 MsgBox "审核成功!", vbInformation, "友情提示"
 
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

Comflag.AddItem ("Y")
Comflag.AddItem ("N")

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
Set reportRS = GetPOPriceModify(UCase(Trim(CmbCustomer.Text)), UCase(Trim(Comflag.Text)))

With fpS(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub

Private Sub ShowData_WhereCus(customerTemp As String, flagTemp As String)
Set reportRS = GetPOPriceModify2(customerTemp, flagTemp)

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
        
    
    
         .Col = E_FPS0.E_OK
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        
        .SetText E_FPS0.E_SeqId, 0, "记录号"
        .SetText E_FPS0.E_CustID, 0, "客户代码"
        .SetText E_FPS0.E_CUSTNAME, 0, "客户全称"
        .SetText E_FPS0.e_PO, 0, "订单号"
  
        .SetText E_FPS0.E_PODATE, 0, "提供订单日期"
        .SetText E_FPS0.E_POTYPE, 0, "订单类型"
        .SetText E_FPS0.E_PT, 0, "机种"
        .SetText E_FPS0.E_PeaceQty, 0, "订单片数"
        .SetText E_FPS0.E_QTY, 0, "订单数量"
        .SetText E_FPS0.E_BJ, 0, "报价单号"
        .SetText E_FPS0.e_Price, 0, "单价"
        .SetText E_FPS0.E_DIEPrice, 0, "DIE单价"
        .SetText E_FPS0.E_PUnit, 0, "单价单位"
        .SetText E_FPS0.E_File, 0, "返点文件名"
        
        .SetText E_FPS0.E_OK, 0, "选择"
          
       
        .ColWidth(E_FPS0.E_SeqId) = 5
        .ColWidth(E_FPS0.E_CustID) = 8
        .ColWidth(E_FPS0.E_CUSTNAME) = 25
        .ColWidth(E_FPS0.e_PO) = 15
    
        .ColWidth(E_FPS0.E_PODATE) = 10
        .ColWidth(E_FPS0.E_POTYPE) = 8
        .ColWidth(E_FPS0.E_PT) = 10
        .ColWidth(E_FPS0.E_PeaceQty) = 10
        .ColWidth(E_FPS0.E_QTY) = 10
        .ColWidth(E_FPS0.E_BJ) = 15
        
        
        .ColWidth(E_FPS0.e_Price) = 8
        .ColWidth(E_FPS0.E_DIEPrice) = 8
        
        .ColWidth(E_FPS0.E_PUnit) = 5
        .ColWidth(E_FPS0.E_File) = 15
        
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .Col = E_FPS0.e_PO
                .Lock = True
        .Col = E_FPS0.E_PODATE
                .Lock = False
        .Col = E_FPS0.E_POTYPE
                .Lock = False
        .Col = E_FPS0.E_PT
                .Lock = True
        .Col = E_FPS0.E_QTY
                .Lock = False
        .Col = E_FPS0.e_Price
                .Lock = False
        .Col = E_FPS0.E_PUnit
                .Lock = False
        .Col = E_FPS0.E_File
                .Lock = False
        .Col = E_FPS0.E_BJ
                .Lock = False
        .Col = E_FPS0.E_OK
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
