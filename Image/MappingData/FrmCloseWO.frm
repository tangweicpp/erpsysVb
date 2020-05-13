VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmCloseWO 
   Caption         =   "工单关闭"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19725
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
   ScaleHeight     =   9675
   ScaleWidth      =   19725
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "不可关闭"
      Height          =   6375
      Left            =   9960
      TabIndex        =   7
      Top             =   240
      Width           =   9615
      Begin VB.CommandButton Command4 
         Caption         =   "导出明细"
         Height          =   465
         Left            =   4320
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "查询工单"
         Height          =   465
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5415
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   960
         Width           =   9615
         _Version        =   196613
         _ExtentX        =   16960
         _ExtentY        =   9551
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
         SpreadDesigner  =   "FrmCloseWO.frx":0000
         TextTip         =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "可关闭"
      Height          =   6375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   9615
      Begin VB.CommandButton Command1 
         Caption         =   "查询工单"
         Height          =   465
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "导出明细"
         Height          =   465
         Left            =   4320
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5415
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   960
         Width           =   9615
         _Version        =   196613
         _ExtentX        =   16960
         _ExtentY        =   9551
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
         SpreadDesigner  =   "FrmCloseWO.frx":4474
         TextTip         =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "单笔关闭"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   19455
      Begin VB.CommandButton CmdQuit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "退出"
         Height          =   480
         Left            =   13320
         TabIndex        =   14
         Top             =   360
         Width           =   990
      End
      Begin VB.CommandButton CmdClose 
         BackColor       =   &H000000FF&
         Caption         =   "关闭工单"
         Height          =   480
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   990
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
      Begin MSDataListLib.DataCombo DtComb 
         Height          =   315
         Left            =   6240
         TabIndex        =   2
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label LblWo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单："
         Height          =   315
         Left            =   5640
         TabIndex        =   4
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "产线类别："
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmCloseWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mainItemRS As New ADODB.Recordset
Dim bomRS2        As New ADODB.Recordset
Dim bomRS3        As New ADODB.Recordset
Private Enum E_FPS0          'Detail
    E_ID = 0                 'id
    E_WOId                   'wo
    E_Product                   '料号
    E_CreatedQty             '开单数
    E_InvQty                 '库存数
    E_WipQty                 '在制数
    E_FinishRate             '完工比率
    E_BomFlag                'Bom领料flag
    E_CloseFlag              '是否关闭
    E_End
    
End Enum

Private Enum E_FPS1          'Detail
    E_ID = 0                 'id
    E_WOId                   'wo
    E_Product                   '料号
    E_CreatedQty             '开单数
    E_InvQty                 '库存数
    E_WipQty                 '在制数
    E_FinishRate             '完工比率
    E_BomFlag                'Bom领料flag
    E_CloseFlag              '是否关闭
    E_End
    
End Enum



Private Sub CmdClose_Click()

Dim userid As String
Dim queryWoTemp As String

userid = UCase(gUserName)
queryWoTemp = ""


'2015-11-24 jiayun add check 线别
If Combo1.Text = "" Then
    MsgBox "请选择产线类别！", vbInformation, "友情提示"
    Exit Sub
End If


queryWoTemp = UCase(Trim(DtComb.Text))

If queryWoTemp = "" Then
    MsgBox "请选择工单号!", vbInformation, "友情提示"
    Exit Sub
     
Else
   '校验Wo是否正确
   If (Not JudgeOracleCloseWo(queryWoTemp)) Then
   
      MsgBox "请确认工单号是否正确!", vbInformation, "友情提示"
      Exit Sub
   
   End If
 
End If




'Dim woTemp As String
'
'If DtComb.Text = "" Then
' MsgBox "请选择要关闭的工单！"
'     Exit Sub
'End If
'woTemp = DtComb.Text
'
'
''判断 Oracle中 Wip上是否有数据，如果有，不允许关闭。
'
'If Combo1.Text = "TSV" Then
'
'    If (JudgeOracleWipWo(Trim(woTemp))) Then
'       MsgBox "该笔工单：" & woTemp & " 存在于Mes Wip上，不可以关闭！"
'       Exit Sub
'
'    End If
'
'End If
'
''2013-05-51 jiayun add
'
''判断ERP中是否存在没有领的料，如果存在，则不允许关
'
'If Combo1.Text = "WLO" Then
'    Set bomRS2 = GetWLOWoBomLing(woTemp)
'    If bomRS2.RecordCount > 0 Then
'        MsgBox "该笔工单在新系统中还有料没有领，不可以关闭工单！"
'        Exit Sub
'    End If
'End If
'
'
'Call DoCloseWo(woTemp)
'
'If Combo1.Text = "TSV" Then
'
'Call IniWO(1)
'
'ElseIf Combo1.Text = "WLO" Then
'Call IniWO(2)
'
'End If



Dim woTemp As String

If DtComb.Text <> "" Then

'单笔关闭工单

woTemp = UCase(Trim(DtComb.Text))


'判断 Oracle中 Wip上是否有数据，如果有，不允许关闭。

If Combo1.Text = "TSV" Then

    If (JudgeOracleWipWo(Trim(woTemp))) Then
       MsgBox "该笔工单：" & woTemp & " 存在于Mes Wip上，不可以关闭！"
       Exit Sub
    
    End If

End If

'2013-05-51 jiayun add

'判断ERP中是否存在没有领的料，如果存在，则不允许关

If Combo1.Text = "WLO" Then
    Set bomRS2 = GetWLOWoBomLing(woTemp)
    If bomRS2.RecordCount > 0 Then
        MsgBox "该笔工单在新系统中还有料没有领，不可以关闭工单！"
        Exit Sub
    End If
End If


Call DoCloseWoNew(woTemp, userid)

If Combo1.Text = "TSV" Then

Call IniWO(1)

ElseIf Combo1.Text = "WLO" Then
Call IniWO(2)

End If

Else

'从表格从选择工单关闭

'DoFtpData
'GetFpsData

Dim aa As Integer
aa = 0


End If


End Sub

Private Sub DoFtpData()
Dim woTemp As String

With fps(0)

For i = 1 To .MaxRows

    .Row = i
    .Col = 8
    If .Text = "1" Then

    .Row = i
    .Col = 1
    woTemp = Trim(.Text)
    
    Call DoCloseWo(woTemp)
 
    End If

Next i


End With


End Sub


Private Sub CmdDelMesWo_Click()
Dim woTemp As String
Dim createDateTemp As Date

Dim i As Integer


If Trim(TxtWO2.Text) = "" Then
 MsgBox "请输入工单号！"
     Exit Sub
End If
woTemp = UCase(Trim(TxtWO2.Text))

'查询一下，这笔工单建立的日期

Set bomRS2 = GetWOCreateDate(woTemp)
If bomRS2.RecordCount <= 0 Then
    MsgBox "这笔工单不存在，请确认工单号 ！"
    Exit Sub
    
Else
    createDateTemp = CDate(bomRS2.fields("createDate").Value)
    
    i = DateDiff("n", createDateTemp, Now)

    If i > 10 Then
        MsgBox "时间隔了太久，不允许删除 ！"
        Exit Sub
        
   Else

        Call DelMesWO(woTemp)
    
    End If


End If



End Sub

Private Sub CmdQuit_Click()
Unload Me
End Sub

Private Sub Combo1_Change()
If Combo1.Text = "TSV" Then

Call IniWO(1)

ElseIf Combo1.Text = "WLO" Then
Call IniWO(2)

End If


End Sub

Private Sub Combo1_Click()
If Combo1.Text = "TSV" Then

Call IniWO(1)

ElseIf Combo1.Text = "WLO" Then
Call IniWO(2)

End If
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
        
        .Col = E_FPS0.E_CloseFlag
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
  
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
          
        .SetText E_FPS0.E_ID, 0, "序号"
        .SetText E_FPS0.E_WOId, 0, "工单号"
        .SetText E_FPS0.E_Product, 0, "成品料号"
        .SetText E_FPS0.E_CreatedQty, 0, "开单数量"
        .SetText E_FPS0.E_InvQty, 0, "入库数量"
        .SetText E_FPS0.E_WipQty, 0, "在制数量"
        .SetText E_FPS0.E_FinishRate, 0, "完工比率"
        .SetText E_FPS0.E_BomFlag, 0, "Bom料是否已领满"
        .SetText E_FPS0.E_CloseFlag, 0, "是否关闭"

        .ColWidth(E_FPS0.E_ID) = 5
        .ColWidth(E_FPS0.E_WOId) = 15
        .ColWidth(E_FPS0.E_Product) = 15
        .ColWidth(E_FPS0.E_CreatedQty) = 15
        .ColWidth(E_FPS0.E_InvQty) = 15
        .ColWidth(E_FPS0.E_WipQty) = 15
        .ColWidth(E_FPS0.E_FinishRate) = 15
        .ColWidth(E_FPS0.E_BomFlag) = 15
        .ColWidth(E_FPS0.E_CloseFlag) = 15
     

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .Col = E_FPS0.E_CloseFlag
        .Lock = False
        
        .ReDraw = True
    End With
    
    
    
    
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
        
        .Col = E_FPS1.E_CloseFlag
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
  
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
          
        .SetText E_FPS1.E_ID, 0, "序号"
        .SetText E_FPS1.E_WOId, 0, "工单号"
        .SetText E_FPS1.E_Product, 0, "成品料号"
        .SetText E_FPS1.E_CreatedQty, 0, "开单数量"
        .SetText E_FPS1.E_InvQty, 0, "入库数量"
        .SetText E_FPS1.E_WipQty, 0, "在制数量"
        .SetText E_FPS1.E_FinishRate, 0, "完工比率"
        .SetText E_FPS1.E_BomFlag, 0, "Bom料是否已领满"
        .SetText E_FPS1.E_CloseFlag, 0, "是否关闭"

        .ColWidth(E_FPS1.E_ID) = 5
        .ColWidth(E_FPS1.E_WOId) = 15
        .ColWidth(E_FPS1.E_Product) = 15
        .ColWidth(E_FPS1.E_CreatedQty) = 15
        .ColWidth(E_FPS1.E_InvQty) = 15
        .ColWidth(E_FPS1.E_WipQty) = 15
        .ColWidth(E_FPS1.E_FinishRate) = 15
        .ColWidth(E_FPS1.E_BomFlag) = 15
        .ColWidth(E_FPS1.E_CloseFlag) = 15
     

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .Col = E_FPS1.E_CloseFlag
        .Lock = False
        
        .ReDraw = True
    End With
    
    

End Sub


Private Sub Command1_Click()
'查询报表

GetFpsData



End Sub

Private Sub Command2_Click()

Dim sqltemp As String

' sqlTemp = " select 工单号,PRODUCT as 成品料号,Qty as 开单数量,FGQty as 入库数量,Qty-FGQty as 在制数量,rtrim(round(cast((FGQty/Qty)* 100 as decimal(10,2)),2))+'%' as 完工比率,BomStatus as Bom料是否已领满 ,'' from ( " & _
'" select x.工单号,x.PRODUCT,x.QTY,  erpdata.dbo.Get_TSV_WO_DieQty(x.工单号) as FGQty,erpdata.dbo.Get_TSV_WO_BomFinish(x.工单号) as BomStatus  from ( " & _
'" select distinct e.工单号,f.PRODUCT ,f.QTY from  [erpbase].[dbo].[tblllplan] e , [erpdata].[dbo].[tblTSVworkorder] f " & _
'" where    f.ORDERNAME=e.工单号 and  e.产线标记=1  and e.完工标记=0 ) X)Y "
  
  sqltemp = "SELECT  orderName,PRODUCT,woQty,invQty,wipQty,finishRate,BomStatus,flag FROM [erpdata].[dbo].[Vw_TSV_CloseWo] where wipQty<1 and BomStatus='是'"
  
  
  SqlServerExporToExcel (sqltemp)





End Sub

Private Sub Command3_Click()
GetFpsData2
End Sub

Private Sub Command4_Click()


Dim sqltemp As String

' sqlTemp = " select 工单号,PRODUCT as 成品料号,Qty as 开单数量,FGQty as 入库数量,Qty-FGQty as 在制数量,rtrim(round(cast((FGQty/Qty)* 100 as decimal(10,2)),2))+'%' as 完工比率,BomStatus as Bom料是否已领满 ,'' from ( " & _
'" select x.工单号,x.PRODUCT,x.QTY,  erpdata.dbo.Get_TSV_WO_DieQty(x.工单号) as FGQty,erpdata.dbo.Get_TSV_WO_BomFinish(x.工单号) as BomStatus  from ( " & _
'" select distinct e.工单号,f.PRODUCT ,f.QTY from  [erpbase].[dbo].[tblllplan] e , [erpdata].[dbo].[tblTSVworkorder] f " & _
'" where    f.ORDERNAME=e.工单号 and  e.产线标记=1  and e.完工标记=0 ) X)Y "
  
  
  sqltemp = "SELECT  orderName,PRODUCT,woQty,invQty,wipQty,finishRate,BomStatus,flag FROM [erpdata].[dbo].[Vw_TSV_CloseWo] where wipQty>0 or BomStatus='否'"
  
  
  SqlServerExporToExcel (sqltemp)


End Sub

Private Sub DtComb_Click(Area As Integer)
'选择工单后，查询数量及Bom量




End Sub


Private Sub GetFpsData()
'明细数据

Set bomRS2 = GetSqlServerFpsCloseWo1()
If bomRS2.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
End If


With fps(0)
        .MaxRows = 0
        If bomRS2.RecordCount > 0 Then
            Set .DataSource = bomRS2
        End If
End With



End Sub

Private Sub GetFpsData2()
'明细数据

Set bomRS3 = GetSqlServerFpsCloseWo2()
If bomRS3.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
End If


With fps(1)
        .MaxRows = 0
        If bomRS3.RecordCount > 0 Then
            Set .DataSource = bomRS3
        End If
End With



End Sub



Private Sub Form_Activate()
Combo1.SetFocus
IniFpsHeader
End Sub

Private Sub Form_Load()
Combo1.AddItem ("TSV")
Combo1.AddItem ("WLO")




'IniWO
End Sub

Private Sub IniWO(lineTypeTemp As Integer)
Set mainItemRS = GetCloseWo(lineTypeTemp)
Set DtComb.RowSource = mainItemRS
DtComb.ListField = mainItemRS("productname").Name
DtComb.BoundColumn = mainItemRS("PID").Name

End Sub
