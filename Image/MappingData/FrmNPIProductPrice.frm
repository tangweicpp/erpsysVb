VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmNPIProductPrice 
   Caption         =   "市场部NPI产品价格维护"
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
      Caption         =   "导出报表"
      Height          =   360
      Left            =   11760
      TabIndex        =   19
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "退出"
      Height          =   360
      Left            =   9180
      TabIndex        =   18
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "清空"
      Height          =   360
      Left            =   6600
      TabIndex        =   17
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "修改"
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
      Caption         =   "记录号:"
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "调价历史1："
      Height          =   195
      Left            =   14040
      TabIndex        =   14
      Top             =   960
      Width           =   990
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "调价历史2："
      Height          =   195
      Left            =   9840
      TabIndex        =   12
      Top             =   960
      Width           =   990
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NRE返还方式："
      Height          =   195
      Left            =   5160
      TabIndex        =   10
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NRE费用(Y/N)&开票日期："
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "测试费(RMB)："
      Height          =   195
      Left            =   13920
      TabIndex        =   6
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "封装费(RMB)："
      Height          =   195
      Left            =   9720
      TabIndex        =   4
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "测试费(USD)："
      Height          =   195
      Left            =   5280
      TabIndex        =   2
      Top             =   360
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "封装费(USD)："
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
Private Enum E_FPS0          'Detail
    E_SeqId = 1                '序号
    E_CustName               '客户代码
    E_QtechPT                '厂内项目名称
    E_QtechPT2                '成品料号
    E_CustPT1                '客户机种名1
    E_CustPT2                '客户机种名2
    E_CustDie                '客户设计die数
    E_XS                   '像素
    
    E_FZU                   '封装费USD
    E_TestU                '测试费USD
    E_FZR                  '封装费RMB
    E_TestR                '测试费RMB
    
    E_NreF                 'NRE费用
    E_NreW                 'NRE返还方式
    E_UP2                  '调价历史2
    E_UP1                  '调价历史1
    
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

'校验是否重复

If UCase(Trim(CmbCustomer.Text)) = "" Or UCase(Trim(TxtQtechPT.Text)) = "" Then
     MsgBox "客户代码或厂内项目名称不可以为空！"
     Exit Sub
End If

If UCase(Trim(TxtCustPT1.Text)) = "" And UCase(Trim(TxtCustPT2.Text)) = "" Then
     MsgBox "客户机种不可以为空！"
     Exit Sub
End If


 Set bomRS2 = GetNpiProductCheck(UCase(Trim(CmbCustomer.Text)), UCase(Trim(TxtQtechPT.Text)), UCase(Trim(TxtCustPT1.Text)), UCase(Trim(TxtCustPT2.Text)), UCase(Trim(TxtQtechPT2.Text)))
If bomRS2.RecordCount > 0 Then
    MsgBox "系统中已存在这笔数据，请重新确认输入是否正确 ！"
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

 MsgBox "新增成功!", vbInformation, "友情提示"
 
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
'修改

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

 MsgBox "修改成功!", vbInformation, "友情提示"

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
        .SetText E_FPS0.E_CustName, 0, "客户代码"
        .SetText E_FPS0.E_QtechPT, 0, "厂内项目名称"
        .SetText E_FPS0.E_QtechPT2, 0, "成品料号"
        .SetText E_FPS0.E_CustPT1, 0, "客户机种名1"
        .SetText E_FPS0.E_CustPT2, 0, "客户机种名2"
        .SetText E_FPS0.E_CustDie, 0, "客户设计die数"
        .SetText E_FPS0.E_XS, 0, "像素"
        
        .SetText E_FPS0.E_FZU, 0, "封装费(USD)"
        .SetText E_FPS0.E_TestU, 0, "测试费(USD)"
        .SetText E_FPS0.E_FZR, 0, "封装费(RMB)"
        .SetText E_FPS0.E_TestR, 0, "测试费(RMB)"
        
        .SetText E_FPS0.E_NreF, 0, "NRE费用(Y/N)&开票日期"
        .SetText E_FPS0.E_NreW, 0, "NRE返还方式"
        .SetText E_FPS0.E_UP2, 0, "调价历史2"
        .SetText E_FPS0.E_UP1, 0, "调价历史1"
        
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



