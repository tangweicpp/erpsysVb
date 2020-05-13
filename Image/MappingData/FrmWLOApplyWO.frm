VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmWLOApplyWO 
   Caption         =   "WLO 下工单"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   18765
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CommandDelBills 
      BackColor       =   &H000000FF&
      Caption         =   "删除工单"
      Height          =   495
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9480
      Width           =   975
   End
   Begin VB.CommandButton CmdBom 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Bom料设定"
      Height          =   480
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "清空数据"
      Height          =   480
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "导出Detail"
      Height          =   480
      Left            =   10770
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "导出Header"
      Height          =   480
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton ComSave 
      BackColor       =   &H00FFC0FF&
      Caption         =   "保存工单"
      Height          =   480
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "工单Detail"
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   18615
      Begin VB.TextBox txtFileName 
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   840
         Width           =   4935
      End
      Begin VB.CommandButton CmdOpenFile 
         Caption         =   ".."
         Height          =   495
         Left            =   5520
         TabIndex        =   22
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton CmdSaveFile 
         Caption         =   "上传DB"
         Height          =   480
         Left            =   6240
         TabIndex        =   21
         Top             =   840
         Width           =   1335
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5655
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   15855
         _Version        =   196613
         _ExtentX        =   27966
         _ExtentY        =   9975
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
         SpreadDesigner  =   "FrmWLOApplyWO.frx":0000
         TextTip         =   2
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2880
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择待上传的xls："
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   24
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "工单Header"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18615
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmWLOApplyWO.frx":4474
         Left            =   6960
         List            =   "FrmWLOApplyWO.frx":4476
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox TxtPiece 
         Height          =   285
         Left            =   1080
         TabIndex        =   25
         Top             =   960
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   9960
         TabIndex        =   13
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   25362433
         CurrentDate     =   40882
      End
      Begin VB.TextBox TxtDate 
         Height          =   285
         Left            =   6960
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtNum 
         Height          =   285
         Left            =   3720
         TabIndex        =   8
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3720
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox CmbCustomer 
         Height          =   315
         ItemData        =   "FrmWLOApplyWO.frx":4478
         Left            =   1080
         List            =   "FrmWLOApplyWO.frx":447A
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo Text3 
         Height          =   315
         Left            =   9840
         TabIndex        =   16
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label LblPiece 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生产片数"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预计完工日"
         Height          =   195
         Left            =   9000
         TabIndex        =   12
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预计开工日"
         Height          =   195
         Left            =   6000
         TabIndex        =   11
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生产数量"
         Height          =   195
         Left            =   3000
         TabIndex        =   9
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "产品料号"
         Height          =   195
         Left            =   9000
         TabIndex        =   7
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单类型"
         Height          =   195
         Left            =   6120
         TabIndex        =   6
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号"
         Height          =   195
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户"
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   360
      End
   End
End
Attribute VB_Name = "FrmWLOApplyWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail
    E_ID = 1                 'id
    E_WaferId                'Waferid
    E_GoodDie                'good数量
    E_End
    
End Enum

Private Enum E_FPS1          'Bom
    E_ID = 0                 'id
    E_BomID                  '材料规范编号
    E_PT                     '料号
    E_Mt                     '物料编号
    E_Name                   '名称
    E_Qty                    '每只用量
    E_Unit                   '单位
    
    E_Pt2                     '料号2
    E_Mt2                     '物料编号2
    E_Name2                   '名称2
    E_Qty2                    '每只用量2
    E_Unit2                   '单位2
    
    E_End
    
End Enum


Dim oiRS        As New ADODB.Recordset
Dim listRS        As New ADODB.Recordset
Dim bomRS2        As New ADODB.Recordset
Dim bomRS3        As New ADODB.Recordset

Dim bomRS        As New ADODB.Recordset

Dim mainItemRS As New ADODB.Recordset
Dim mainItemRS2 As New ADODB.Recordset




Private Sub CmdBom_Click()
FrmWLO_Bom.Show

End Sub

Private Sub CmdM_Click()
FormM.Show

Unload Me

End Sub

Private Sub CmdDelWO_Click()
'7468859023
'TxtSourceBatchId.Text = "7468859023"
If CmbCustomer.Text = "" Or TxtSourceBatchId.Text = "" Then
    MsgBox "请先选择客户，再输入Lot号。请确认!", vbInformation, "友情提示"
    Exit Sub
Else

    Set oiRS = GetOIData(UCase(Trim(CmbCustomer.Text)), UCase(Trim(TxtSourceBatchId.Text)))
    If (oiRS.RecordCount > 0) Then

        TxtPo.Text = getStr(oiRS.fields("po_num").Value)
        TxtCustomerPT.Text = getStr(oiRS.fields("mpn_desc").Value)
        TxtFab.Text = getStr(oiRS.fields("fabrication_facility").Value)
        TxtCusRev.Text = getStr(oiRS.fields("imager_customer_rev").Value)
        TxtDesignId.Text = getStr(oiRS.fields("design_id").Value)
        Txt260.Text = getStr(oiRS.fields("shipping_mst_260").Value)
        Text11.Text = getStr(oiRS.fields("shipping_mst_level").Value)
        TxtMarkingcode.Text = getStr(oiRS.fields("encoded_mark_id").Value)
        TxtCounFab.Text = getStr(oiRS.fields("country_of_fab").Value)
        TxtMMaterial.Text = getStr(oiRS.fields("micron_material").Value)
        TxtPoItem.Text = getStr(oiRS.fields("po_item").Value)
        TxtLotStatus.Text = getStr(oiRS.fields("lot_status").Value)
        TxtMpn.Text = getStr(oiRS.fields("mpn").Value)
        
        If getStr(oiRS.fields("protective_film_apld").Value) = "YES" Then
            TxtFilmApld.Text = "PF"
        Else
            TxtFilmApld.Text = getStr(oiRS.fields("protective_film_apld").Value)
        End If
        
        TxtRequestDate.Text = getStr(oiRS.fields("lot_priority").Value)
        TxtShipSite.Text = getStr(oiRS.fields("ship_site").Value)
        
        If TxtShipSite.Text = "Qtech" And UCase(Trim(CmbCustomer.Text)) = "AA" Then
            CmbCheckCustomer.Text = "WLC"
            
        ElseIf TxtShipSite.Text = "SG" And UCase(Trim(CmbCustomer.Text)) = "AA" Then
            CmbCheckCustomer.Text = "AA"
            
        ElseIf UCase(Trim(CmbCustomer.Text)) = "GC" Then
             CmbCheckCustomer.Text = "GC"
        End If
        
        Call IniProductTwo(UCase(Trim(CmbCustomer.Text)))
        
        '初始化左边的Lot明细表
        
'        Call InitListBox(UCase(Trim(CmbCustomer.Text)))

        
        Call InitListBoxForSo(UCase(Trim(CmbCustomer.Text)), TxtPo.Text)
        
        
        
        
        
        '2012-12-18 jiayun add  根据source_batch_id ，自动带出料号
        '2013-05-13 jiayun add GC
'        Text3.Text = "aa"
         If CmbCustomer.Text = "AA" Then
        
            Call getAutoWo(UCase(Trim(TxtSourceBatchId.Text)))
        
        ElseIf CmbCustomer.Text = "GC" Or CmbCustomer.Text = "SX" Or CmbCustomer.Text = "PT" Or CmbCustomer.Text = "SY" Then
        
             Call getOtherCustomerPt(UCase(Trim(TxtSourceBatchId.Text)))
        
        Else
        
        End If
        
    Else
        MsgBox "查询不到数据，请确认 SourceBatchId "
        Exit Sub
 
    
    End If
    

    
    
End If
End Sub

Private Sub getAutoWo(lotidTemp2 As String)

Dim lotidtemp As String
lotidtemp = lotidTemp2
Dim pfType As String
Dim trayType As String
Dim testno As String

Dim ptFirst As String

pfType = GetString(lotidtemp)
'LblPF.Caption = pfType

trayType = GetTrayString(lotidtemp)
'LblTrayType.Caption = trayType

testno = GetTestNoString(lotidtemp)
'LblTestNo.Caption = testno

'成品料号
'根据OI，查出成品料号的前9位

ptFirst = GetFirstPtString(lotidtemp)

Dim test1 As String
test1 = GetAllPtString(ptFirst, pfType, trayType, testno)

Text3.Text = GetAllPtString(ptFirst, pfType, trayType, testno)

End Sub
Private Function getStr(strTemp As Variant)
getStr = Trim("" & strTemp)
End Function

'2013-05-13 jiayun add
Private Sub getOtherCustomerPt(lotidTemp2 As String)

Text3.Text = GetCustomerPtNum(lotidTemp2)

End Sub


Private Sub Command2_Click()
Dim strTmp As String
Dim strTemp As String
strTemp = ""
With Lst
        '开始查找赋值
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                strTmp = GetLot003(.List(i)) & "','"
                strTemp = strTemp & strTmp

            End If
        Next
 End With
 
 If strTemp = "" Then
 
 MsgBox "请先选择LotId !"
 Exit Sub
 
 Else
 
 strTemp = Mid(strTemp, 1, Len(strTemp) - 3)
 
Call GetFpsData(strTemp, UCase(Trim(CmbCustomer.Text)))

'ChkAll.Value = 1
'ChkAll_Click

End If


End Sub
'2013-04-23 jiayun add
Private Function GetLot003(lotidtemp As String)

GetLot003 = Replace(lotidtemp, "00A", "003")

End Function

Private Sub GetFpsData(strwhereTemp As String, customerTemp As String)
'明细数据

Set listRS = GetFps(strwhereTemp, customerTemp)
If listRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
End If

With fps(0)
        .MaxRows = 0
        If listRS.RecordCount > 0 Then
            Set .DataSource = listRS
        End If
End With

End Sub

Private Sub GetBomData(ptTemp As String)
'明细数据

Set bomRS = GetFpsBom(ptTemp)
If bomRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
End If

With fps(1)
        .MaxRows = 0
        If bomRS.RecordCount > 0 Then
            Set .DataSource = bomRS
        End If
End With

End Sub



Private Sub InitListBox(customerTemp As String)
Dim i As Integer
      Set listRS = GetLotDetailData(customerTemp)
       With Lst
            .Clear
            listRS.MoveFirst
            
            For i = 0 To listRS.RecordCount - 1
            
         
                .AddItem "" & listRS!source_batch_id
                
                If "" & listRS!source_batch_id = TxtSourceBatchId.Text Then
                    Lst.Selected(i) = True
                End If
                
                listRS.MoveNext
         
            
            Next
        End With
        
      
        

listRS.Close
Set listRS = Nothing

End Sub

Private Sub InitListBoxForSo(customerTemp As String, soTemp As String)
Dim i As Integer
      Set listRS = GetLotDetailDataForSo(customerTemp, soTemp)
       With Lst
            .Clear
            listRS.MoveFirst
            
            For i = 0 To listRS.RecordCount - 1
            
         
                .AddItem "" & listRS!source_batch_id
                
                If "" & listRS!source_batch_id = TxtSourceBatchId.Text Then
                    Lst.Selected(i) = True
                End If
                
                listRS.MoveNext
         
            
            Next
        End With
        
      
        

listRS.Close
Set listRS = Nothing

End Sub


Private Sub CmdOpenFile_Click()

On Error Resume Next
Dim FName
'帅选文件
CommonDialog1.Filter = "EXCEL文件(*.xlsx)|*.xlsx"
CommonDialog1.ShowOpen
'得到文件名
FName = CommonDialog1.FileName
If FName <> "" Then
   txtFileName.Text = FName
End If


End Sub

Private Sub CmdSaveFile_Click()

'上传导入的Excel
If txtFileName.Text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String



'清除上一次数据
Dim cmdStr As String
cmdStr = "delete from  WLO_WO_DetailTemp  "
AddSql (cmdStr)




'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(txtFileName.Text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 3 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If


Dim i As Integer
Dim j As Integer
Dim tempVal As String

Dim idTemp As Long
Dim waferidTemp As String
Dim goodQtyTemp As Integer

   
 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    idTemp = 0
    waferidTemp = ""
    goodQtyTemp = 0
    
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
           
        If j = 1 Then
            idTemp = CInt(Trim(tempVal))
        End If
        
        If j = 2 Then
            waferidTemp = Trim(tempVal)
        End If
        
        If j = 3 Then
            goodQtyTemp = CInt(Trim(tempVal))
        End If
        
    Next j
  
    Call AddWLOWaferTemp(idTemp, waferidTemp, goodQtyTemp)
Next i

     xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

  ' VBExcel.Quit


GetFpsWaferData






End Sub


Private Sub GetFpsWaferData()
'明细数据

Set listRS = GetFpsWLOWaferDetail()
If listRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
End If

With fps(0)
        .MaxRows = 0
        If listRS.RecordCount > 0 Then
            Set .DataSource = listRS
        End If
End With


End Sub



Public Sub AddWLOWaferTemp(idTemp As Long, waferidTemp As String, goodQtyTemp As Integer)
Dim cmdStr As String

cmdStr = "insert into WLO_WO_DetailTemp (ID , SUBSTRATEID ,  PASSBINCOUNT   ) values (" & idTemp & ",'" & waferidTemp & "'," & goodQtyTemp & ") "
                     
AddSql (cmdStr)

End Sub


Private Sub Command3_Click()
  
 If UCase(Trim(Text2.Text)) = "" Then
      MsgBox "工单号不可以为空！"
     Exit Sub
 
 End If
 
  
 Dim sqlTemp As String
' sqlTemp = "select SEQ_IBWO,ORDERNAME,ORDERTYPE,DESCRIPTION,EVENTTYPE,ERPUSER,PRODUCT,PRODUCTREVISION,QTY,PRODUCTBOM,ERPCREATEDATE,PLANSTARTDATE,PLANENDDATE," & _
'         " Customer , SalesOrder, PRODUCTFAMILY, ModifyFlag, CUSTOMERPN, FabFacility, ImagerRev, Designid, MLevel235, Mlevel260, NGFlag, Para1, Para2, Para3, Para4, Para5, Para6, PARA7, PARA8, PARA9, PARA10, Protective_Film_Apld, LOT_STATUS, MPN " & _
'         " From IB_WOHISTORY where ORDERNAME='" + Text2.Text + "'order by SEQ_IBWO desc "
  
  
 sqlTemp = " select customer, ordername,ordertype,product,para1 as 片数 , qty,planstartdate,planenddate,erpuser,erpcreatedate from  erpintegration2.wlo_ib_workorder  where ORDERNAME='" + Text2.Text + "' "
  
  
  
  ExporToExcel (sqlTemp)
End Sub

Private Sub Command4_Click()

 If UCase(Trim(Text2.Text)) = "" Then
      MsgBox "工单号不可以为空！"
     Exit Sub
 
 End If
 
 Dim sqlTemp As String

 sqlTemp = " select b.ordername,a.product,b.waferid,dieqty from erpintegration2.wlo_ib_workorder a,  erpintegration2.WLO_IB_WAFERLIST b Where a.OrderName = b.OrderName and a.ORDERNAME='" + Text2.Text + "' order by b.waferid "

  ExporToExcel (sqlTemp)


End Sub

Private Sub Command5_Click()

ClearData

End Sub

Private Sub ClearData()
'清空上一笔的数据
Text2.Text = ""
TxtPiece.Text = ""
TxtNum.Text = ""
txtFileName.Text = ""
fps(0).MaxRows = 0


End Sub

Private Sub Command6_Click()

On Error Resume Next

Dim lotidtemp As String
Dim qtyTemp As Long
Dim productTemp As String
Dim erpdate As String



Set billLotTemp = GetBillLot2()
If (billLotTemp.RecordCount > 0) Then

    Do While Not billLotTemp.EOF
        lotidtemp = billLotTemp.fields("waferlot").Value
        productTemp = billLotTemp.fields("productname").Value
        qtyTemp = CLng(billLotTemp.fields("qty").Value)
        
        erpdate = Format(CDate(billLotTemp.fields("erpcreationdate").Value), "YYYY-MM-DD")
        
   
                
          '-----begin------
        
         Set adoCmd = New ADODB.Command
         Set adoCmd.ActiveConnection = INIadoCon2
             adoCmd.CommandText = "uspPMC_XDInterface"
             adoCmd.Parameters.Refresh
             adoCmd.CommandType = adCmdStoredProc
        
          Set adoprm1 = New ADODB.Parameter   '工单号
          adoprm1.Type = adChar
          adoprm1.Size = 20
          adoprm1.Direction = adParamInput
          adoprm1.Value = lotidtemp
          adoCmd.Parameters.Append adoprm1
        
          Set adoprm2 = New ADODB.Parameter   '料号
          adoprm2.Type = adChar
          adoprm2.Size = 20
          adoprm2.Direction = adParamInput
          adoprm2.Value = productTemp
          adoCmd.Parameters.Append adoprm2
        
          Set adoprm3 = New ADODB.Parameter   '数量
          adoprm3.Type = adInteger
          adoprm3.Direction = adParamInput
          adoprm3.Value = qtyTemp
          adoCmd.Parameters.Append adoprm3
          
            Set adoprm4 = New ADODB.Parameter   '数量
             
          adoprm4.Type = adChar
         adoprm4.Size = 20
          adoprm4.Direction = adParamInput
          adoprm4.Value = erpdate
          adoCmd.Parameters.Append adoprm4
        
          adoCmd.Execute

        
        billLotTemp.MoveNext
   
    Loop
    
End If

MsgBox ("OK")





End Sub

Private Sub CommandDelBills_Click()
'删除工单
Dim woTemp As String

 If UCase(Trim(Text2.Text)) = "" Then
      MsgBox "工单号不可以为空！"
     Exit Sub
 
 End If
 
 '判断是否已经领料，如果领料，则不可以删除
 woTemp = UCase(Trim(Text2.Text))
 
  '2012-11-30 jiayun add 判断料号的bom是否存在
Set bomRS2 = GetWLOWo(woTemp)
If bomRS2.RecordCount > 0 Then
    MsgBox "该笔工单在新系统中已经领过料，不可以删除工单！"
    Exit Sub
End If

'GetWLOWoBomLing

'删除

Call DelWLOBillHeaderWo(woTemp)
 


End Sub

Private Sub ComSave_Click()
'保存工单
Dim headerTemp As WLOBillHeader
Dim detailTemp As BillDetail
Dim typeId As Integer
Dim SumQty As Long
Dim i As Integer
SumQty = 0


'附值
 headerTemp.id = GetSeqID()
 headerTemp.OrderName = UCase(Trim(Text2.Text))
 
 If UCase(Trim(Text2.Text)) = "" Then
      MsgBox "工单号不可以为空！"
     Exit Sub
 
 End If
 
 
 '2012-11-30 jiayun add 判断料号的bom是否存在
Set bomRS2 = GetProductBom(Text3.Text)
If bomRS2.RecordCount <= 0 Then
    MsgBox "新系统中这料号的Bom不存在！请联系相关的人，先维护Bom ！"
    Exit Sub
End If


 
 '2013-06-17 jiayun add 判断工单号是否存在
Set bomRS3 = CheckWLOWo(UCase(Trim(Text2.Text)))
If bomRS3.RecordCount > 0 Then
    MsgBox "工单号: " & UCase(Trim(Text2.Text)) & " 已存在，请重新确认工单号！"
    Exit Sub
End If


' '2012-12-19 jiayun add 校验料号是否存在
'Set bomRS2 = GetProduct_Check(Text3.Text)
'If bomRS2.RecordCount <= 0 Then
'    MsgBox "料号不存在！请联系相关的人，先维护料号 ！"
'    Exit Sub
'End If



 
'Select Case Combo2.Text
'Case "一般工单"
'    typeId = 1
'Case "再加工工单"
'    typeId = 5
'Case "委外工单"
'    typeId = 7
'
'Case "重工委外工单"
'    typeId = 8
'
'Case "拆件式工单"
'    typeId = 11
'
'Case "预测工单"
'    typeId = 13
'Case "试产工单"
'    typeId = 15
'
'Case Else
'   typeId = 0
'End Select

 headerTemp.OrderType = Trim(Combo2.Text)
 
 headerTemp.EventType = "CREATED"
 headerTemp.ERPUser = gUserName
 headerTemp.product = Text3.Text
 headerTemp.Customer = UCase(Trim(CmbCustomer.Text))
                            
 headerTemp.RequestDate = Now
 headerTemp.ERPCreateDate = DateTime.Date
 headerTemp.PlanStartDate = CDate(TxtDate.Text)
 headerTemp.PlanEndDate = DTPicker1.Value


 
With fps(0)

For i = 1 To .MaxRows
        .Row = i
        .Col = 3
        SumQty = SumQty + CInt(.Text)

Next i

End With

headerTemp.qty = SumQty
headerTemp.PieceQty = fps(0).MaxRows





 '--附值End
  Call AddWLOBillHeaderWo(headerTemp)
  
'--保存Heand End

'--- Begin Detail

'判断这笔工单，对应客户的OI,是否已用完



'MsgBox "工单：" & Text2.Text & "建立成功 !"




End Sub

Private Sub Form_Activate()

'IniBillType

IniWLOProduct


'Text3.Text = "S5000100910F"


End Sub

Private Sub Form_Load()
CmbCustomer.AddItem ("XT")

'
'CmbCheckCustomer.AddItem ("AA")
'CmbCheckCustomer.AddItem ("WLC")
'CmbCheckCustomer.AddItem ("GC")
'CmbCheckCustomer.AddItem ("SX")
'CmbCheckCustomer.AddItem ("SY")


'IniProduct

TxtDate.Text = Format(Now, "yyyy-mm-dd")
DTPicker1.Value = TxtDate.Text

'Combo2.AddItem ("一般工单")
'Combo2.AddItem ("再加工工单")
'Combo2.AddItem ("委外工单")
'Combo2.AddItem ("重工委外工单")
'Combo2.AddItem ("拆件式工单")
'Combo2.AddItem ("预测工单")
'Combo2.AddItem ("试产工单")
'Combo2.AddItem ("小批量试产工单")
Combo2.AddItem ("一般工单")
Combo2.AddItem ("再加工工单")
Combo2.AddItem ("委外工单")
Combo2.AddItem ("重工委外工单")
Combo2.AddItem ("拆件式工单")
Combo2.AddItem ("预测工单")
Combo2.AddItem ("试产工单")
Combo2.AddItem ("小批量试产工单")




IniFpsHeader
'IniFpsBom



End Sub

Private Sub IniProduct()
Set mainItemRS = GetProduct()
Set Text3.RowSource = mainItemRS
Text3.ListField = mainItemRS("productname").Name
Text3.BoundColumn = mainItemRS("PID").Name

End Sub

'Private Sub IniBillType()
'Set mainItemRS = GetBillType()
'Set Combo2.RowSource = mainItemRS
'Combo2.ListField = mainItemRS("名称").Name
'Combo2.BoundColumn = mainItemRS("说明2").Name
'
'End Sub

Private Sub IniWLOProduct()
Set mainItemRS2 = GetWLOBomProduct()
Set Text3.RowSource = mainItemRS2
Text3.ListField = mainItemRS2("物料编号1").Name
Text3.BoundColumn = mainItemRS2("物料编号2").Name

End Sub




Private Sub IniProductTwo(customerTemp As String)
If customerTemp = "AA" Then
    Set Text3.RowSource = Nothing
    Set mainItemRS = GetProductAA()
    Set Text3.RowSource = mainItemRS
    Text3.ListField = mainItemRS("productname").Name
    Text3.BoundColumn = mainItemRS("PID").Name
    
 ElseIf customerTemp = "GC" Then
    
    Set Text3.RowSource = Nothing
    Set mainItemRS = GetProductBB()
    Set Text3.RowSource = mainItemRS
    Text3.ListField = mainItemRS("productname").Name
    Text3.BoundColumn = mainItemRS("PID").Name
    
End If

'Set mainItemRS = GetProduct()
'Set Text3.RowSource = mainItemRS
'Text3.ListField = mainItemRS("productname").Name
'Text3.BoundColumn = mainItemRS("PID").Name

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
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
    
          
        .SetText E_FPS0.E_ID, 0, "序号"
        .SetText E_FPS0.E_WaferId, 0, "WaferId"
        .SetText E_FPS0.E_GoodDie, 0, "GoodDie数量"

        
        
        .ColWidth(E_FPS0.E_ID) = 10
        .ColWidth(E_FPS0.E_WaferId) = 15
        .ColWidth(E_FPS0.E_GoodDie) = 12

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub IniFpsBom()
    With fps(1)
        .ReDraw = False
        .MaxCols = E_FPS1.E_End - 1
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
        
      
        
        .SetText E_FPS1.E_ID, 0, "序号"
        .SetText E_FPS1.E_BomID, 0, "材料规范编号"
        .SetText E_FPS1.E_PT, 0, "料号"
        .SetText E_FPS1.E_Mt, 0, "物料编号"
        .SetText E_FPS1.E_Name, 0, "名称"
        .SetText E_FPS1.E_Qty, 0, "每只用量"
        .SetText E_FPS1.E_Unit, 0, "单位"
        
        .SetText E_FPS1.E_Pt2, 0, "备料料号"
        .SetText E_FPS1.E_Mt2, 0, "备料物料编号"
        .SetText E_FPS1.E_Name2, 0, "备料名称"
        .SetText E_FPS1.E_Qty2, 0, "备料每只用量"
        .SetText E_FPS1.E_Unit2, 0, "备料单位"
    
        
        
        .ColWidth(E_FPS1.E_ID) = 6
        .ColWidth(E_FPS1.E_BomID) = 12
        .ColWidth(E_FPS1.E_PT) = 14
        .ColWidth(E_FPS1.E_Mt) = 14
        .ColWidth(E_FPS1.E_Name) = 14
        .ColWidth(E_FPS1.E_Qty) = 10
        .ColWidth(E_FPS1.E_Unit) = 8
        
        .ColWidth(E_FPS1.E_Pt2) = 14
        .ColWidth(E_FPS1.E_Mt2) = 14
        .ColWidth(E_FPS1.E_Name2) = 14
        .ColWidth(E_FPS1.E_Qty2) = 10
        .ColWidth(E_FPS1.E_Unit2) = 8
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
End Sub


'Private Sub Text2_KeyPress(KeyAscii As Integer)
''生成工单号
''年年月月+四位编码
'Dim FirstChar As String
'Dim SeqChar As String
''2012-11-06 因新旧系统　临时取消自动生成
'
''FirstChar = UCase(Trim(Text2.Text))
'' If KeyAscii = 13 Then
''    If FirstChar = "" Then
''        MsgBox "请输入工单前三位!"
''        Exit Sub
''    End If
''
''    FirstChar = FirstChar & "-" & Right(Year(DateTime.Date), 2) & Right("0" & Month(DateTime.Date), 2)
''
''    SeqChar = Right("000" & CStr(CInt(GetSeqChar()) + 1), 4)
''
''    Text2.Text = FirstChar & SeqChar
''
''    If Mid$(Trim(Text2.Text), 2, 1) = "P" Then
''        Combo2.Text = "一般工单"
''    End If
''
''    If Mid$(Trim(Text2.Text), 2, 1) = "T" Then
''        Combo2.Text = "小批量试产工单"
''    End If
''
''
'' End If
'
'
'End Sub

Private Sub Text3_Change()
'选择产品料号，来显示Bom
'Dim ptTemp As String
''ptTemp = Text3.Text
'
'ptTemp = "18V117FD00CF"
' Call GetBomData(ptTemp)



End Sub

