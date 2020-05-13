VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmTSV_Bom2 
   Caption         =   "Bom料(分批,样品,重工)"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17085
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
   ScaleHeight     =   7005
   ScaleWidth      =   17085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdAddRow 
      Caption         =   "add"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton CmdDelRow 
      Caption         =   "del"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "工单Bom"
      Height          =   6615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   16575
      Begin VB.CommandButton CmdReturn 
         Caption         =   "返回前窗体"
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Add后保存"
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   6015
         Index           =   1
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   15975
         _Version        =   196613
         _ExtentX        =   28178
         _ExtentY        =   10610
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
         SpreadDesigner  =   "FrmTSV_Bom2.frx":0000
         TextTip         =   2
      End
   End
End
Attribute VB_Name = "FrmTSV_Bom2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS1          'Bom
    E_ID = 1                 'id
    E_OK                     '选择
    E_BomID                  '材料规范编号
    E_PT                     '料号
    E_Mt                     '物料编号
    E_Name                   '名称
    E_Qty                    '每只用量
    E_Rate                   '损耗量
    E_Unit                   '单位
    E_Pt2                     '料号2
    E_Mt2                     '物料编号2
    E_Name2                   '名称2
    E_Qty2                    '每只用量2
    E_Rate2                   '损耗量
    E_Unit2                   '单位2
    E_End
    
End Enum

Dim bomRS        As New ADODB.Recordset
Public ptTemp As String



Private Sub CmdAddRow_Click()
 With fps(1)
        .MaxRows = .MaxRows + 1
    End With
End Sub

Private Sub CmdDelRow_Click()
'删除行
'取当前行的Id
Dim currentRowTemp As String

With fps(1)
    Set .DataSource = Nothing
    .Row = .ActiveRow
    .Col = 1
    currentRowTemp = .Text
    .DeleteRows .ActiveRow, 1
    .MaxRows = .MaxRows - 1
End With

'删除当前行
Call DelRowBomDataNew(currentRowTemp, gUserName)

'Call GetBomData(ptTemp)



End Sub

Private Sub CmdReturn_Click()
'FrmApplyWO2.Show
Unload Me

End Sub

Private Sub CmdSave_Click()
'添加新行后，保存到DB
Dim i As Integer
Dim tempID As String
Dim tempID2 As Integer
Dim tempPt As String
Dim tempEMt As String
Dim tempName As String
Dim tempQty As String
Dim tempUnit As String
Dim tempRate As String

i = 1
With fps(1)
    For i = 1 To .MaxRows
        empInfAll = False
        
        .Row = i
        .Col = E_FPS1.E_ID
        tempID = .Text
        
        If tempID = "" Then
            .Row = i
            .Col = E_FPS1.E_PT
            tempPt = .Text
            
            .Row = i
            .Col = E_FPS1.E_Mt
            tempEMt = .Text
            
            .Row = i
            .Col = E_FPS1.E_Name
            tempName = .Text
            
            .Row = i
            .Col = E_FPS1.E_Qty
            tempQty = .Text
            
            .Row = i
            .Col = E_FPS1.E_Rate
            tempRate = .Text
            
            .Row = i
            .Col = E_FPS1.E_Unit
            tempUnit = .Text
            
            If tempPt <> "" And CDbl(tempQty) > 0 Then
            
                AddRecordNew tempPt, tempEMt, tempName, tempQty, tempUnit, tempRate, gUserName
            End If
        
        End If
         
    Next i

End With

'2013-04-25 jiayun add 修改用量

Dim qty2Temp As Double

 With fps(1)

For i = 1 To .MaxRows

    .Row = i
    .Col = E_FPS1.E_OK
    
    If .Text <> "" Then
    
        If CInt(.Text) = 1 Then
        .Row = i
        .Col = E_FPS1.E_ID
        
         tempID2 = CInt(.Text)
         
        .Row = i
        .Col = E_FPS1.E_Qty
        
        qty2Temp = CDbl(.Text)
        
         cmdStrSql = " update  [erpdata].[dbo].[TSVtblBillBom2InitData] set Qty=" & qty2Temp & "  where ID=" & tempID2 & "  and employ='" & gUserName & "' "
    
        AddSql2 (cmdStrSql)
        End If
    
    End If

Next i


End With




'刷新一下fps

Call GetBomDataNew(ptTemp, gUserName)



End Sub

Public Sub AddRecord(tempPt As String, tempEMt As String, tempName As String, tempQty As String, tempUnit As String, tempRate As String)
'把添加的行，保存到DB中
Dim cmdStr As String
         
cmdStr = "INSERT INTO [erpdata].[dbo].[TSVtblBillBom2InitData]([PT],[WLID],[Name],[Qty],[Unit],[SHRateQty]) " & _
         " VALUES('" + tempPt + "','" + tempEMt + "','" + tempName + "','" + tempQty + "','" + tempUnit + "','" + tempRate + "')"


                             
 Call AddSql2(cmdStr)
 
End Sub

Public Sub AddRecordNew(tempPt As String, tempEMt As String, tempName As String, tempQty As String, tempUnit As String, tempRate As String, userID As String)
'把添加的行，保存到DB中
Dim cmdStr As String
         
cmdStr = "INSERT INTO [erpdata].[dbo].[TSVtblBillBom2InitData]([PT],[WLID],[Name],[Qty],[Unit],[SHRateQty],employ) " & _
         " VALUES('" + tempPt + "','" + tempEMt + "','" + tempName + "','" + tempQty + "','" + tempUnit + "','" + tempRate + "','" + userID + "')"


                             
 Call AddSql2(cmdStr)
 
End Sub




Private Sub Form_Load()
IniFpsBom

InitData

End Sub

'Private Sub fps_ComboSelChange(Index As Integer, ByVal Col As Long, ByVal Row As Long)
'Dim tempPT As String
'Dim tempE_Mt As String
'
'
'If Col = E_FPS1.E_Pt Then
'    With fps(1)
'            .Row = .ActiveRow
'            .Col = E_FPS1.E_BomId
'
'            If .Text <> "" Then
'            '材料规范编号为空，说明是新加的行
'            .Row = .ActiveRow
'            .Col = E_FPS1.E_Pt
'            tempPT = UCase$(.Text)
'            If tempPT <> "" Then
'                tempE_Mt = GetE_Mt(tempPT)
'                .Row = .ActiveRow
'                .Col = E_FPS1.E_Mt
'                .Text = tempE_Mt
'
'
'            End If
'
'
'            End If
'
'    End With
'
'End If
'
'End Sub

Public Function GetE_Mt(tempPt As String) As String
'查询物料编号
Dim cmdStr As String
Dim RSResult As String
cmdStr = "SELECT 物料编号 FROM dbo.tblSmainM2 WHERE 料号='" + tempPt + "' "
     
RSResult = GetSqlServerStr(cmdStr)
GetE_Mt = RSResult
End Function

Public Function GetE_Name(tempPt As String) As String
'查询名称
Dim cmdStr As String
Dim RSResult As String
cmdStr = "SELECT 物料名称 FROM dbo.tblSmainM2 WHERE 料号='" + tempPt + "' "
     
RSResult = GetSqlServerStr(cmdStr)
GetE_Name = RSResult
End Function

Public Function GetE_Unit(tempPt As String) As String
'查询单位
Dim cmdStr As String
Dim RSResult As String
cmdStr = "SELECT 计量单位名称 FROM dbo.tblSmainM2 WHERE 料号='" + tempPt + "' "
     
RSResult = GetSqlServerStr(cmdStr)
GetE_Unit = RSResult
End Function





Private Sub InitData()
'选择产品料号，来显示Bom

'ptTemp = FrmApplyWO2.Text3.Text

ptTemp = bomProductTemp

'ptTemp = "18V117FD00CF"

'2012-04-18
'建立一个临时表，用来存放每次查询出来的Bom;
'删除数据


DelBomDataNew (gUserName)


If Mid(UCase(Trim(woSendTemp)), 2, 1) = "R" Or Mid(UCase(Trim(woSendTemp)), 2, 1) = "R" Then
Call AddBomDataForReworkNew(ptTemp, gUserName)

Else

Call AddBomDataNew(ptTemp, gUserName)


End If



Call GetBomDataNew(ptTemp, gUserName)



End Sub

Private Sub GetBomData(ptTemp As String)
'明细数据

Set bomRS = GetFpsReworkBom(ptTemp)
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

Private Sub GetBomDataNew(ptTemp As String, userID As String)
'明细数据

Set bomRS = GetFpsReworkBomNew(ptTemp, userID)
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


Private Sub DelBomData()
'删掉初始表数据

Dim sqlTemp As String

sqlTemp = "DELETE FROM [erpdata].[dbo].[TSVtblBillBom2InitData]"
Call AddSql2(sqlTemp)

End Sub

Private Sub DelBomDataNew(userID As String)
'删掉初始表数据

Dim sqlTemp As String

sqlTemp = "DELETE FROM [erpdata].[dbo].[TSVtblBillBom2InitData] where employ = '" & userID & "'"
Call AddSql2(sqlTemp)

End Sub


Private Sub DelRowBomData(rowId As String)
'删掉当前行数据

Dim sqlTemp As String

sqlTemp = "DELETE FROM [erpdata].[dbo].[TSVtblBillBom2InitData] where id=" & rowId & ""
Call AddSql2(sqlTemp)

End Sub

Private Sub DelRowBomDataNew(rowId As String, userID As String)
'删掉当前行数据

Dim sqlTemp As String

sqlTemp = "DELETE FROM [erpdata].[dbo].[TSVtblBillBom2InitData] where id=" & rowId & " and employ ='" & userID & "' "
Call AddSql2(sqlTemp)

End Sub



Private Sub AddBomData(ptTemp As String)
'添加初始表数据

Dim sqlTemp As String

'sqlTemp = "INSERT INTO [erpdata].[dbo].[TSVtblBillBom2InitData] ( [MID],[PT],[WLID],[Name],[Qty],[SHRateQty],[Unit],[PT1],[WLID1],[Name1],[Qty1],[SHRateQty1],[Unit1] ) " & _
'" SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.损耗,b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.损耗1,b.单位1 " & _
'" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
'" Where a.材料规范编号 = b.材料规范编号 AND (a.物料编号='" & ptTemp & "'   OR a.是否共用标记=1) AND b.料号 NOT LIKE '%TCF' AND b.料号 NOT LIKE '%CCF' AND b.料号1 NOT LIKE '%TCF' AND b.料号1 NOT LIKE '%CCF' ORDER BY b.材料规范编号 DESC"
'
 
 sqlTemp = "INSERT INTO [erpdata].[dbo].[TSVtblBillBom2InitData] ( [MID],[PT],[WLID],[Name],[Qty],[SHRateQty],[Unit],[PT1],[WLID1],[Name1],[Qty1],[SHRateQty1],[Unit1] ) " & _
" SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.损耗,b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.损耗1,b.单位1 " & _
" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
" Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & ptTemp & "' AND b.料号 NOT LIKE '18%' AND b.料号 NOT LIKE '19%' AND b.料号1 NOT LIKE '18%' AND b.料号1 NOT LIKE '19%' Union  " & _
" SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.损耗,b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.损耗1,b.单位1 " & _
" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
" Where a.材料规范编号 = b.材料规范编号 AND a.材料规范编号=( select childBomNameID  from [erpdata].[dbo].[TSVtblBomSetup] c where c.ProductName='" & ptTemp & "' and c.Flag='Y') ORDER BY b.材料规范编号 DESC "

 
Call AddSql2(sqlTemp)

End Sub


Private Sub AddBomDataNew(ptTemp As String, userID As String)
'添加初始表数据

Dim sqlTemp As String

'sqlTemp = "INSERT INTO [erpdata].[dbo].[TSVtblBillBom2InitData] ( [MID],[PT],[WLID],[Name],[Qty],[SHRateQty],[Unit],[PT1],[WLID1],[Name1],[Qty1],[SHRateQty1],[Unit1] ) " & _
'" SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.损耗,b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.损耗1,b.单位1 " & _
'" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
'" Where a.材料规范编号 = b.材料规范编号 AND (a.物料编号='" & ptTemp & "'   OR a.是否共用标记=1) AND b.料号 NOT LIKE '%TCF' AND b.料号 NOT LIKE '%CCF' AND b.料号1 NOT LIKE '%TCF' AND b.料号1 NOT LIKE '%CCF' ORDER BY b.材料规范编号 DESC"
'
 
 sqlTemp = "INSERT INTO [erpdata].[dbo].[TSVtblBillBom2InitData] ( [MID],[PT],[WLID],[Name],[Qty],[SHRateQty],[Unit],[PT1],[WLID1],[Name1],[Qty1],[SHRateQty1],[Unit1],employ ) " & _
" SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.损耗,b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.损耗1,b.单位1, '" & userID & "' " & _
" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
" Where a.材料规范编号 = b.材料规范编号 AND a.物料编号='" & ptTemp & "' AND b.料号 NOT LIKE '18%' AND b.料号 NOT LIKE '19%' AND b.料号1 NOT LIKE '18%' AND b.料号1 NOT LIKE '19%' Union  " & _
" SELECT  b.材料规范编号,b.料号,b.物料编号,b.名称,cast(b.每只用量 as varchar),b.损耗,b.单位,b.料号1,b.物料编号1,b.名称1,cast(b.备料用量 as varchar),b.损耗1,b.单位1 ,'" & userID & "' " & _
" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
" Where a.材料规范编号 = b.材料规范编号 AND a.材料规范编号=( select childBomNameID  from [erpdata].[dbo].[TSVtblBomSetup] c where c.ProductName='" & ptTemp & "' and c.Flag='Y') ORDER BY b.材料规范编号 DESC "

 
Call AddSql2(sqlTemp)

End Sub


Private Sub AddBomDataForRework(ptTemp As String)
'添加初始表数据

Dim sqlTemp As String

sqlTemp = " INSERT INTO [erpdata].[dbo].[TSVtblBillBom2InitData] " & _
          " ( [MID],[PT],[WLID],[Name],[Qty],[SHRateQty],[Unit],[PT1],[WLID1],[Name1],[Qty1],[SHRateQty1],[Unit1] )" & _
          "  select '' as 材料规范编号, b.料号,b.物料编号,b.物料名称 as 名称,'1' as 每只用量, 0 as 损耗, " & _
          "  'PCS' as 单位, '' as 料号1,'' as 物料编号1,'' as 名称1,'' as 备料用量,'' as 损耗1,'' as 单位1 " & _
          "  from  [erpdata].[dbo].[tblSmainM2] b where b.料号='" & ptTemp & "' "
  
Call AddSql2(sqlTemp)

End Sub

Private Sub AddBomDataForReworkNew(ptTemp As String, userID As String)
'添加初始表数据

Dim sqlTemp As String

sqlTemp = " INSERT INTO [erpdata].[dbo].[TSVtblBillBom2InitData] " & _
          " ( [MID],[PT],[WLID],[Name],[Qty],[SHRateQty],[Unit],[PT1],[WLID1],[Name1],[Qty1],[SHRateQty1],[Unit1],employ )" & _
          "  select '' as 材料规范编号, b.料号,b.物料编号,b.物料名称 as 名称,'1' as 每只用量, 0 as 损耗, " & _
          "  'PCS' as 单位, '' as 料号1,'' as 物料编号1,'' as 名称1,'' as 备料用量,'' as 损耗1,'' as 单位1 , '" & userID & "' " & _
          "  from  [erpdata].[dbo].[tblSmainM2] b where b.料号='" & ptTemp & "' "
  
Call AddSql2(sqlTemp)

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
        
         .Col = E_FPS1.E_OK
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        

        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
      
        
        .SetText E_FPS1.E_ID, 0, "序号"
        .SetText E_FPS1.E_OK, 0, "选择"
        
        .SetText E_FPS1.E_BomID, 0, "材料规范编号"
        .SetText E_FPS1.E_PT, 0, "料号"
        .SetText E_FPS1.E_Mt, 0, "物料编号"
        .SetText E_FPS1.E_Name, 0, "名称"
        .SetText E_FPS1.E_Qty, 0, "每只用量"
        .SetText E_FPS1.E_Rate, 0, "损耗"
        .SetText E_FPS1.E_Unit, 0, "单位"
        
        .SetText E_FPS1.E_Pt2, 0, "备料料号"
        .SetText E_FPS1.E_Mt2, 0, "备料物料编号"
        .SetText E_FPS1.E_Name2, 0, "备料名称"
        .SetText E_FPS1.E_Qty2, 0, "备料每只用量"
        .SetText E_FPS1.E_Rate2, 0, "备料损耗"
        .SetText E_FPS1.E_Unit2, 0, "备料单位"
    
        
        .ColWidth(E_FPS1.E_ID) = 6
        .ColWidth(E_FPS1.E_OK) = 10
        .ColWidth(E_FPS1.E_BomID) = 12
        .ColWidth(E_FPS1.E_PT) = 14
        .ColWidth(E_FPS1.E_Mt) = 14
        .ColWidth(E_FPS1.E_Name) = 14
        .ColWidth(E_FPS1.E_Qty) = 10
        .ColWidth(E_FPS1.E_Rate) = 6
        .ColWidth(E_FPS1.E_Unit) = 8
        
        .ColWidth(E_FPS1.E_Pt2) = 14
        .ColWidth(E_FPS1.E_Mt2) = 14
        .ColWidth(E_FPS1.E_Name2) = 14
        .ColWidth(E_FPS1.E_Qty2) = 10
         .ColWidth(E_FPS1.E_Rate2) = 6
        .ColWidth(E_FPS1.E_Unit2) = 8
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .Col = E_FPS1.E_ID
        .Lock = True
        
        .Col = E_FPS1.E_BomID
        .Lock = True
        
        .Col = E_FPS1.E_OK
        .Lock = False
        
        
        
        .ReDraw = True
    End With
    
    
    

End Sub

Private Sub fps_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim tempPt As String
Dim tempE_Mt As String
Dim tempName As String
Dim tempUnit As String


If Col = E_FPS1.E_PT Then
    With fps(1)
            .Row = .ActiveRow
            .Col = E_FPS1.E_BomID
            
            If .Text = "" Then
            '材料规范编号为空，说明是新加的行
            .Row = .ActiveRow
            .Col = E_FPS1.E_PT
            tempPt = Trim(UCase$(.Text))
            If tempPt <> "" Then
                tempE_Mt = GetE_Mt(tempPt)
                .Row = .ActiveRow
                .Col = E_FPS1.E_Mt
                .Text = tempE_Mt
                
                
                '名称
                tempName = GetE_Name(tempPt)
                .Row = .ActiveRow
                .Col = E_FPS1.E_Name
                .Text = tempName
                
                '单位
                tempUnit = GetE_Unit(tempPt)
                .Row = .ActiveRow
                .Col = E_FPS1.E_Unit
                .Text = tempUnit
                
                
            
            End If
            
            
            End If

    End With
    
End If

End Sub
