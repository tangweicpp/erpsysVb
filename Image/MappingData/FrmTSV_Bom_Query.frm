VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmTSV_Bom_Query 
   Caption         =   "TSV Bom 查询与修改 与新Bom模板数据上传"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   18585
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "上传"
      Height          =   1215
      Left            =   600
      TabIndex        =   15
      Top             =   120
      Width           =   12255
      Begin VB.CommandButton Command7 
         Caption         =   "上传DB"
         Height          =   480
         Left            =   8280
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   ".."
         Height          =   495
         Left            =   7080
         TabIndex        =   17
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   480
         Width           =   5895
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   3720
         Top             =   -120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径："
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择待上传的xlsx："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   600
         TabIndex        =   19
         Top             =   120
         Width           =   1620
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "修改"
      Height          =   855
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   12255
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H000000FF&
         Caption         =   "删除"
         Height          =   360
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdAddSave 
         BackColor       =   &H000080FF&
         Caption         =   "添加后提交"
         Height          =   360
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "添加一行"
         Height          =   360
         Left            =   6000
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton CmdModify 
         BackColor       =   &H00C0C000&
         Caption         =   "修改用量提交"
         Height          =   360
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtModify 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用量："
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查询"
      Height          =   360
      Left            =   11880
      TabIndex        =   7
      Top             =   1680
      Width           =   990
   End
   Begin VB.TextBox TxtPT2 
      Height          =   375
      Left            =   8880
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox TxtPT 
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox TxtID 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   5775
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   17895
      _Version        =   196613
      _ExtentX        =   31565
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
      SpreadDesigner  =   "FrmTSV_Bom_Query.frx":0000
      TextTip         =   2
   End
   Begin VB.Label LblPT2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "物料料号："
      Height          =   195
      Left            =   7920
      TabIndex        =   5
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label LblPT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "成品料号："
      Height          =   195
      Left            =   4320
      TabIndex        =   3
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label LblId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bom编号："
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   840
   End
End
Attribute VB_Name = "FrmTSV_Bom_Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS1          'Bom
    E_ID = 1                 'id
    E_BomID                  '材料规范编号
    E_Createdby              '建立人
    E_CreatedDate              '建立日期
    E_PT                     '料号
    E_Mt                     '物料编号
    E_Name                   '名称
    E_GG                     '规格
    E_XH                      '型号
    
    E_Qty                    '每只用量
    E_Rate                   '损耗量
    E_Unit                   '单位
    
    E_Typeid                 '序号
    E_TypePT                 '材料类型
    E_End
    
End Enum

Dim bomRS        As New ADODB.Recordset
Public ptTemp As String
Public bomProduct As String




Private Sub CmdAdd_Click()
Dim i As Integer
Dim qtyTemp As String
Dim bomIDtTemp As String
Dim bomProductTemp As String


 With fps(1)
        .MaxRows = .MaxRows + 1
        i = .MaxRows
        
        .Row = i - 1
        .Col = 1
        bomIDtTemp = .Text
        
        
        .Row = i - 1
        .Col = 2
        bomProductTemp = .Text
        
        .Row = i
        .Col = 1
        .Text = bomIDtTemp
        
        .Row = i
        .Col = 2
        .Text = bomProductTemp
        
 End With
 
 
 
 
 
 
End Sub

Private Sub CmdAddSave_Click()
Dim i As Integer
Dim tempProduct As String
Dim bomId As String
Dim product As String
Dim ptid As String
Dim ptname As String
Dim pttype As String
Dim pttypename As String
Dim qty As Double
Dim qtysh As Double
Dim unit As String
Dim notemp As String
Dim noName As String




With fps(1)

     For i = 1 To .MaxRows
     
      .Row = i
      .Col = 5
      tempProduct = Trim(.Text)
      
      If tempProduct = bomProduct Then
      
      '插入到SqlServer里
      
'      sqlTemp = "Update [erpdata].[dbo].[TSVtblMRuleData]  Set 每只用量 = " & qtyTemp & " where 材料规范编号='" & bomIDtTemp & "' and 料号='" & bomProductTemp & "'"
'        AddSql2 (sqlTemp)
        
      .Row = i
      .Col = 1
       bomId = Trim(.Text)
       
      .Row = i
      .Col = 2
       product = Trim(.Text)
       
'     .Row = i
'      .Col = 5
'       product = Trim(.Text)
       
     .Row = i
      .Col = 6
       ptid = Trim(.Text)
       
     .Row = i
      .Col = 7
       ptname = Trim(.Text)
       
      .Row = i
      .Col = 8
       pttype = Trim(.Text)
       
       
      .Row = i
      .Col = 9
       pttypename = Trim(.Text)
       
      .Row = i
      .Col = 10
       qty = CDbl(Trim(.Text))
       
     .Row = i
      .Col = 11
      qtysh = CDbl(Trim(.Text))
      
       
      .Row = i
      .Col = 12
       unit = Trim(.Text)
       
      .Row = i
      .Col = 13
       notemp = Trim(.Text)
       
      .Row = i
      .Col = 14
       noName = Trim(.Text)
       
       
       sqlTemp = " insert into [erpdata].[dbo].[TSVtblMRuleData](材料规范编号,工序号,料号,物料编号,名称,规格," & _
      " 型号 , 每只用量,损耗,单位,序号,材料类型)" & _
      "  values ('" & bomId & "','" & product & "','" & tempProduct & "','" & ptid & "','" & ptname & "','" & pttype & "'," & _
      " '" & pttypename & "'," & qty & "," & qtysh & ",'" & unit & "','" & notemp & "','" & noName & "') "
      
        AddSql2 (sqlTemp)
         
        MsgBox "添加成功!", vbInformation, "友情提示"
        
        CmdQuery_Click

      
      
      
      End If
      
      
     
     Next i
     
  
    
    

End With


End Sub

Private Sub CmdDel_Click()
Dim qtyTemp As String
Dim bomIDtTemp As String
Dim bomProductTemp As String



With fps(1)
    .Row = .ActiveRow
    .Col = 1
    bomIDtTemp = Trim(.Text)
    
    .Row = .ActiveRow
    .Col = 5
    bomProductTemp = Trim(.Text)
    
    

End With

     
'更新到SqlServer

sqlTemp = "delete from  [erpdata].[dbo].[TSVtblMRuleData]  where 材料规范编号='" & bomIDtTemp & "' and 料号='" & bomProductTemp & "'"
AddSql2 (sqlTemp)

 MsgBox "删除成功!", vbInformation, "友情提示"

CmdQuery_Click



End Sub

Private Sub CmdModify_Click()
Dim qtyTemp As String
Dim bomIDtTemp As String
Dim bomProductTemp As String


If Trim(TxtModify.Text) = "" Then

    MsgBox "用量不可以为空！", vbInformation, "友情提示"
    Exit Sub
 
 Else
    qtyTemp = CLng(Trim(TxtModify.Text))
    
End If

With fps(1)
    .Row = .ActiveRow
    .Col = 1
    bomIDtTemp = Trim(.Text)
    
    .Row = .ActiveRow
    .Col = 5
    bomProductTemp = Trim(.Text)
    
    

End With

     
'更新到SqlServer

sqlTemp = "Update [erpdata].[dbo].[TSVtblMRuleData]  Set 每只用量 = " & qtyTemp & " where 材料规范编号='" & bomIDtTemp & "' and 料号='" & bomProductTemp & "'"
AddSql2 (sqlTemp)

 MsgBox "修改用量成功!", vbInformation, "友情提示"

CmdQuery_Click


End Sub

Private Sub CmdQuery_Click()
'查询
Dim sqlTemp As String

Dim sqltemp1 As String

Dim sqlTemp2 As String

Dim sqltemp3 As String


  sqltemp1 = "select a.[材料规范编号],a.[物料编号],a.工艺,a.建立日期,b.料号 , b.物料编号, b.名称, b.规格, b.型号, b.每只用量, b.损耗, b.单位, b.序号, b.材料类型" & _
             " from [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b Where a.[材料规范编号] = b.[材料规范编号]"
             
  sqlTemp2 = ""
             
  sqltemp3 = " order by a.[材料规范编号],a.[物料编号], b.料号"
  
  
 If Trim(TxtID.Text) <> "" Then
 
 sqlTemp2 = sqlTemp2 & " and a.材料规范编号 like '%" & UCase(Trim(TxtID.Text)) & "%'"
 
 End If
 
  If Trim(TxtPT.Text) <> "" Then
 
 'sqltemp2 = sqltemp2 & " and a.物料编号='" & UCase(Trim(TxtPT.Text)) & "'"
 
 sqlTemp2 = sqlTemp2 & " and a.物料编号 like '%" & UCase(Trim(TxtPT.Text)) & "%'"
  
 
 End If
 
 If Trim(TxtPT2.Text) <> "" Then
 
 sqlTemp2 = sqlTemp2 & " and b.料号 like '%" & UCase(Trim(TxtPT2.Text)) & "%'"
 
 End If
 
 sqlTemp = sqltemp1 & sqlTemp2 & sqltemp3



Set bomRS = GetFpsBomQuery(sqlTemp)
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

Private Sub Command6_Click()
On Error Resume Next
Dim FName
    '帅选文件
    CommonDialog2.Filter = "EXCEL文件(*.xlsx)|*.xlsx"
    
    CommonDialog2.ShowOpen
    '得到文件名
    FName = CommonDialog2.FileName
    If FName <> "" Then
       Text3.Text = FName
    End If



End Sub

Private Sub Command7_Click()
'2016-01-06 jiayun add 新Bom上传

Dim recordNoTemp As String
Dim recordNo As String

Dim recordNoCheckTemp As String

Dim productTemp As String
Dim ptTemp As String
Dim qtyTemp As Double
Dim qtyHaoTemp As Integer
Dim specTemp As String
Dim ptTypeTemp As String
Dim addHeaderFlag As Boolean


recordNo = ""
recordNoTemp = ""
recordNoCheckTemp = ""

productTemp = ""
ptTemp = ""
qtyTemp = 0
qtyHaoTemp = 0
specTemp = ""
ptTypeTemp = ""

addHeaderFlag = False


Dim source_batch_id_Temp As String

If Text3.Text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String


    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同
  If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 22 Then
        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If


Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
   


SumCount = 0
BCResultFlag = False

' vtDataTemp.Created_ByTemp = gUserName

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
 
 '一行付值
 
     If Trim(xlSheet.Range(Chr(96 + 8) & i).Value) = "" Or Trim(xlSheet.Range(Chr(96 + 9) & i).Value) = "" Then

        MsgBox "用量与损耗不能为空值，则不允许上传!", vbInformation, "友情提示"
        Exit Sub
     End If

    
 
     recordNoTemp = UCase(CStr(Trim(xlSheet.Range(Chr(96 + 1) & i).Value)))
     productTemp = UCase(CStr(Trim(xlSheet.Range(Chr(96 + 2) & i).Value)))
     ptTemp = UCase(CStr(Trim(xlSheet.Range(Chr(96 + 3) & i).Value)))
     qtyTemp = CDbl(Trim(xlSheet.Range(Chr(96 + 8) & i).Value))
     qtyHaoTemp = CInt(Trim(xlSheet.Range(Chr(96 + 9) & i).Value))
     specTemp = CStr(Trim(xlSheet.Range(Chr(96 + 20) & i).Value))
     ptTypeTemp = CStr(Trim(xlSheet.Range(Chr(96 + 21) & i).Value))
     
     If recordNoCheckTemp = recordNoTemp Then
        addHeaderFlag = True
        
      Else
         addHeaderFlag = False
         
        
     End If
     
     
      '查询此成品料号，是否已建过，如果建过，则不让上传
   
    If addHeaderFlag = False And JudgeBomHeaderStaus(productTemp) Then
        MsgBox "此成品料号的Bom已建过：" & productTemp & "，请确认!", vbInformation, "友情提示"
        Exit Sub
     End If
   
 
'校验是否为空
     If recordNoTemp = "" Or productTemp = "" Or ptTemp = "" Or specTemp = "" Or ptTypeTemp = "" Then

        MsgBox "字段中有空值，则不允许上传!", vbInformation, "友情提示"
        Exit Sub
     End If


'校验成品料号，金碟中有没有  productTemp
     If Not JudgeBomProduct(productTemp) Then
     
        MsgBox "成品料号不对：" & productTemp & "，请确认!", vbInformation, "友情提示"
        Exit Sub
     End If
     
  '校验半成品料号，金碟中有没有  ptTemp
     If Not JudgeBomProduct(ptTemp) Then
     
        MsgBox "半成品料号不对：" & ptTemp & "，请确认!", vbInformation, "友情提示"
        Exit Sub
     End If
 
    '第一次循环，加Bom主表，后面不用加主表
    If addHeaderFlag = False Then
 
       Dim adoprm1 As ADODB.Parameter
       Dim adoPrmReturn As ADODB.Parameter

         Set adoCmd = New ADODB.Command
         Set adoCmd.ActiveConnection = INIadoCon2
         adoCmd.CommandText = "TSVuspgy_setmIndex "
         adoCmd.Parameters.Refresh
         adoCmd.CommandType = adCmdStoredProc

         Set adoPrmReturn = New ADODB.Parameter
         adoPrmReturn.Type = adChar
         adoPrmReturn.Size = 12
         adoPrmReturn.Direction = adParamOutput
         adoPrmReturn.Value = adParamReturnValue
         adoCmd.Parameters.Append adoPrmReturn
         adoCmd.Execute
         recordNo = adoPrmReturn.Value
         
        Call AddBomHeader(recordNo, productTemp)
        
        Call AddBomChild(recordNo, productTemp, ptTemp, qtyTemp, qtyHaoTemp, specTemp, ptTypeTemp)
        
         
       Else
       '只加子表
       
       '(notemp As String, productTemp As String, ptTemp As String, qtyTemp As Double, qtyTemp2 As Integer, specTemp As String, typeTemp As String)
       
        Call AddBomChild(recordNo, productTemp, ptTemp, qtyTemp, qtyHaoTemp, specTemp, ptTypeTemp)
       
       
       End If


    
    SumCount = SumCount + 1
    
    'addHeaderFlag = True
    
    recordNoCheckTemp = recordNoTemp

    '上传到DB
NextRecord2:

Next i


     
     xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit




If SumCount > 0 Then
    MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"
    
    Else
        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"
            Exit Sub
        End If
    
End If






End Sub

Private Sub Form_Load()
IniFpsBom

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
        .Lock = False
        

        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
        .SetText E_FPS1.E_ID, 0, "材料规范编号"
        .SetText E_FPS1.E_BomID, 0, "成品料号"
        .SetText E_FPS1.E_Createdby, 0, "建立人"
        .SetText E_FPS1.E_CreatedDate, 0, "建立日期"
        .SetText E_FPS1.E_PT, 0, "料号"
        .SetText E_FPS1.E_Mt, 0, "物料编号"
        .SetText E_FPS1.E_Name, 0, "名称"
        .SetText E_FPS1.E_GG, 0, "规格"
        .SetText E_FPS1.E_XH, 0, "型号"
        .SetText E_FPS1.E_Qty, 0, "每只用量"
        .SetText E_FPS1.E_Rate, 0, "损耗"
        .SetText E_FPS1.E_Unit, 0, "单位"
        .SetText E_FPS1.E_Typeid, 0, "Bom序号"
        .SetText E_FPS1.E_TypePT, 0, "材料类型"
        

        .ColWidth(E_FPS1.E_ID) = 10
        .ColWidth(E_FPS1.E_BomID) = 12
        .ColWidth(E_FPS1.E_Createdby) = 12
        .ColWidth(E_FPS1.E_CreatedDate) = 12
        
        .ColWidth(E_FPS1.E_PT) = 14
        .ColWidth(E_FPS1.E_Mt) = 14
        .ColWidth(E_FPS1.E_Name) = 14
         .ColWidth(E_FPS1.E_GG) = 14
        .ColWidth(E_FPS1.E_XH) = 14
        
        .ColWidth(E_FPS1.E_Qty) = 10
        .ColWidth(E_FPS1.E_Rate) = 6
        .ColWidth(E_FPS1.E_Unit) = 8
        
        .ColWidth(E_FPS1.E_Typeid) = 6
        .ColWidth(E_FPS1.E_TypePT) = 8

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub fps_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

Dim bomProductTemp As String

   If (Col = E_FPS1.E_PT And NewCol = E_FPS1.E_Mt And NewRow = Row) Then
        With fps(1)
            .Row = Row
            .Col = Col

        bomProductTemp = .Text
        bomProduct = bomProductTemp
        
        '根据料号，查询相关信息
        
          Set oiRS = GetProductChildBomAdd(bomProductTemp)
  
            If (oiRS.RecordCount > 0) Then
              
            .Row = Row
            .Col = Col + 1
            .Text = Trim(oiRS.fields("物料编号").Value)
            
            .Row = Row
            .Col = Col + 2
            .Text = Trim(oiRS.fields("物料名称").Value)
            
            .Row = Row
            .Col = Col + 3
            .Text = Trim(oiRS.fields("规格型号").Value)
            
            .Row = Row
            .Col = Col + 4
            .Text = Trim(oiRS.fields("型号").Value)
            .Row = Row
            .Col = Col + 6
            .Text = "0"
            
            .Row = Row
            .Col = Col + 7
            .Text = Trim(oiRS.fields("计量单位名称").Value)
                
        
            End If
    End With
    End If



End Sub
