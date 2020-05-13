VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Form_ShelfNo 
   Caption         =   "货架号"
   ClientHeight    =   9945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15825
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
   ScaleHeight     =   9945
   ScaleWidth      =   15825
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "录入"
      Height          =   2415
      Left            =   240
      TabIndex        =   6
      Top             =   8400
      Width           =   14415
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1560
         Width           =   1485
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "录    入"
         Height          =   720
         Left            =   3840
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         Height          =   315
         Left            =   720
         TabIndex        =   7
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label lblLotId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "货架类型"
         Height          =   195
         Left            =   960
         TabIndex        =   10
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label lblOrderId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "货架编号"
         Height          =   195
         Left            =   960
         TabIndex        =   9
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "查询"
      Height          =   6975
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   14415
      Begin VB.CommandButton cmdExport 
         Caption         =   "导出"
         Height          =   720
         Left            =   3600
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "查     询"
         Height          =   720
         Left            =   720
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   1845
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   5760
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   1485
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   3735
         Index           =   0
         Left            =   720
         TabIndex        =   12
         Top             =   2400
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   6588
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
         SpreadDesigner  =   "Form_ShelfNo.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类型"
         Height          =   195
         Left            =   3960
         TabIndex        =   15
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LotId"
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OrderId"
         Height          =   195
         Left            =   5760
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   570
      End
   End
End
Attribute VB_Name = "Form_ShelfNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0
    E_id = 0
    E_OrderId
    E_Batch
    E_LOTId
    E_WaferId
    E_ShelfId   ' 货架号
    
    E_End
    
End Enum
Dim mainItemRS As New adodb.Recordset

Private Sub cmdExport_Click()
Dim LotID As String
Dim Customer As String

LotID = UCase(Trim(Text2.Text))
Customer = Combo2.Text

If LotID = "" And Customer = "" Then
   ' MsgBox ("请输入OrderId, LotId")
    MsgBox ("请输入LotId或者选择客户类型")
Else
    
    If LotID <> "" Then
        ExportShelfInfo (LotID)
    End If
    
    If Customer <> "" Then
        ExportShelfInfo2 (Customer)
    End If

End If


End Sub

Private Sub cmdInsert_Click()

Dim ShelfNo As String
Dim ShelfDep As String

ShelfNo = UCase(Trim(Text10.Text))
ShelfDep = Trim(Combo1.Text)

If ShelfNo = "" Or ShelfDep = "" Then
    MsgBox ("请输入货架号, 选择类别")
Else
    Call InsertShelfInfo(ShelfNo, ShelfDep)
    MsgBox ("录入成功!")
End If

End Sub

Private Sub cmdSearch_Click()

'Dim OrderId As String

' Dim Rs As adodb.Recordset

Dim LotID As String
Dim Customer As String

'OrderId = UCase(Trim(Text1.Text))


LotID = UCase(Trim(Text2.Text))
Customer = Combo2.Text

If LotID = "" And Customer = "" Then
   ' MsgBox ("请输入OrderId, LotId")
    MsgBox ("请输入LotId或者选择客户类型")
Else
    Set mainItemRS = GetShelfInfo(LotID)
    'Set mainItemRs = GetShelfInfo(OrderId, LotId)
    
    With fps(0)
        .MaxRows = 0
        
        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
       
        End If
    End With

   If Customer <> "" Then
        Set mainItemRS = GetShelfInfo2(Customer)
        
        With fps(0)
            .MaxRows = 0
        
            If mainItemRS.RecordCount > 0 Then
                Set .DataSource = mainItemRS
       
            End If
        End With
   End If
    
'    If Rs.RecordCount <= 0 Then
'        MsgBox "查询不到信息, 请确认查询条件是否有误！"
'        cmdSearch.Enabled = True
'        Exit Sub
'    Else
'        Text3.Text = IIf(IsNull(Rs.fields("Batch").Value), "", Rs.fields("Batch").Value)
'        Text4.Text = IIf(IsNull(Rs.fields("waferid").Value), "", Rs.fields("waferid").Value)
'        Text5.Text = IIf(IsNull(Rs.fields("stockid").Value), "", Rs.fields("stockid").Value)
'    End If

End If

End Sub

Private Sub Form_Load()

Combo1.AddItem ("CIS")
Combo1.AddItem ("bumping")
Combo1.AddItem ("SSP")

Combo2.AddItem ("CIS")
Combo2.AddItem ("bumping")
Combo2.AddItem ("SSP")

InitDetail


End Sub


Public Function GetShelfInfo(lot As String) As adodb.Recordset

Dim cmdStr As String
Dim RSResult As New adodb.Recordset

cmdStr = " SELECT  a.orderid, a.Batch, a.lotid, a.waferid, a.stockid FROM [erpdata].[dbo].[ZDHW] a Where a.lotid ='" & lot & "'"
'cmdStr = " SELECT  a.orderid, a.Batch, a.lotid, a.waferid, a.stockid FROM [erpdata].[dbo].[ZDHW] a Where a.orderid = '" & Order & "' AND a.lotid ='" & Lot & "'"
 
        
Set RSResult = GetShelfInfoSQL(cmdStr)

Set GetShelfInfo = RSResult
End Function
Public Function GetShelfInfo2(Customer As String) As adodb.Recordset

Dim cmdStr As String
Dim cusStr As String
Dim RSResult As New adodb.Recordset


'cmdStr = " SELECT  a.orderid, a.Batch, a.lotid, a.waferid, a.stockid FROM [erpdata].[dbo].[ZDHW] a Where a.lotid ='" & lot & "'"
'cmdStr = " SELECT  a.orderid, a.Batch, a.lotid, a.waferid, a.stockid FROM [erpdata].[dbo].[ZDHW] a Where a.orderid = '" & Order & "' AND a.lotid ='" & Lot & "'"
Select Case Customer

    Case "bumping"
        cmdStr = "SELECT  a.orderid, a.Batch, a.lotid, a.waferid, a.stockid FROM [erpdata].[dbo].[ZDHW] a where SUBSTRING (a.orderid, 3, 1) = 'B' or SUBSTRING (a.orderid, 3, 1) = 'W' "
    Case "CIS"
        cmdStr = "SELECT  a.orderid, a.Batch, a.lotid, a.waferid, a.stockid FROM [erpdata].[dbo].[ZDHW] a where SUBSTRING (a.orderid, 3, 1) = 'T' "
    Case "SSP"
        cmdStr = "SELECT  a.orderid, a.Batch, a.lotid, a.waferid, a.stockid FROM [erpdata].[dbo].[ZDHW] a where SUBSTRING (a.orderid, 3, 1) = 'Z' "
    Case Else
    
End Select
       
Set RSResult = GetShelfInfoSQL(cmdStr)

Set GetShelfInfo2 = RSResult
End Function
'插入数据
Public Sub InsertShelfInfo(No As String, Dep As String)

Dim cmdStr2 As String
                                                                                                                                               
cmdStr2 = "insert into [ERPDATA].[dbo].[HJ] (HJBH,HJLX) values('" & No & "','" & Dep & "')"
                                                               
AddSql2 (cmdStr2)

End Sub


'初始化表格
Public Sub InitDetail()

Dim cmdStr2 As String
                                                                                                                                               
 With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
        .SetText E_FPS0.E_id, 0, "序号"
        .SetText E_FPS0.E_OrderId, 0, "工单号"
        .SetText E_FPS0.E_Batch, 0, "批号"
        .SetText E_FPS0.E_LOTId, 0, "LOT号"
        .SetText E_FPS0.E_WaferId, 0, "WaferId"
        .SetText E_FPS0.E_ShelfId, 0, "货架号"
                
        .ColWidth(E_FPS0.E_id) = 10
        .ColWidth(E_FPS0.E_OrderId) = 10
        .ColWidth(E_FPS0.E_Batch) = 10
        .ColWidth(E_FPS0.E_LOTId) = 10
        .ColWidth(E_FPS0.E_WaferId) = 10
        .ColWidth(E_FPS0.E_ShelfId) = 10
        
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
             
        .ReDraw = True
    End With

End Sub
'导出表格
Public Sub ExportShelfInfo2(Customer As String)

Dim cmdStr As String

Select Case Customer

    Case "bumping"
        cmdStr = "SELECT  a.orderid as 工单, a.Batch as 批号, a.lotid as lot号, a.waferid as wafer号, a.stockid as 货架号 FROM [erpdata].[dbo].[ZDHW] a where SUBSTRING (a.orderid, 3, 1) = 'B' or SUBSTRING (a.orderid, 3, 1) = 'W' "
    Case "CIS"
        cmdStr = "SELECT  a.orderid as 工单, a.Batch as 批号, a.lotid as lot号, a.waferid as wafer号, a.stockid as 货架号 FROM [erpdata].[dbo].[ZDHW] a where SUBSTRING (a.orderid, 3, 1) = 'T' "
    Case "SSP"
        cmdStr = "SELECT  a.orderid as 工单, a.Batch as 批号, a.lotid as lot号, a.waferid as wafer号, a.stockid as 货架号 FROM [erpdata].[dbo].[ZDHW] a where SUBSTRING (a.orderid, 3, 1) = 'Z' "
    Case Else
    
End Select
       
SqlServer2ExporToExcel (cmdStr)

End Sub

'导出表格
Public Sub ExportShelfInfo(lot As String)

Dim cmdStr As String

cmdStr = " SELECT  a.orderid as 工单, a.Batch as 批号, a.lotid as lot号, a.waferid as wafer号, a.stockid as 货架号 FROM [erpdata].[dbo].[ZDHW] a Where a.lotid ='" & lot & "'"

SqlServer2ExporToExcel (cmdStr)

End Sub



