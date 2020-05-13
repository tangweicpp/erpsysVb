VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmTSV_Bom_Setup 
   Caption         =   "TSV Bom 设定"
   ClientHeight    =   9765
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   17085
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox DcbProductName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "取消"
      Height          =   480
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   990
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "绑定"
      Height          =   480
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   990
   End
   Begin VB.TextBox TxtBom2 
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox TxtBom1 
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo DcbTYBomName 
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   7335
      Index           =   0
      Left            =   720
      TabIndex        =   8
      Top             =   2280
      Width           =   12855
      _Version        =   524288
      _ExtentX        =   22675
      _ExtentY        =   12938
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
      SpreadDesigner  =   "FrmTSV_Bom_Setup.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "通用Bom："
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "成品Bom："
      Height          =   195
      Left            =   6360
      TabIndex        =   1
      Top             =   360
      Width           =   840
   End
   Begin VB.Label LblProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "成品料号："
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "FrmTSV_Bom_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Bom
    E_id = 1                 'id
    E_PRODUCT                '材料规范编号
    E_BomID                  'Bom号
    E_BomName                 'Bom料号
    E_BomDate                 '建立日期
    
    E_End
    
End Enum

Dim bomRS        As New ADODB.Recordset
Dim bomRS1        As New ADODB.Recordset
Dim oiRS        As New ADODB.Recordset
Public ptTemp As String




Private Sub cmdQuery_Click()
'查询
Dim sqlTemp As String

Dim sqlTemp1 As String

Dim sqlTemp2 As String

Dim sqltemp3 As String


  sqlTemp1 = "select a.[材料规范编号],a.[物料编号],a.工艺,a.建立日期,b.料号 , b.物料编号, b.名称, b.规格, b.型号, b.每只用量, b.损耗, b.单位, b.序号, b.材料类型" & _
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

 sqlTemp = sqlTemp1 & sqlTemp2 & sqltemp3



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

Private Sub cmdExit_Click()
TxtBom1.Text = ""
TxtBom2.Text = ""
End Sub



Private Sub CmdOK_Click()
Dim productNameTemp As String
Dim childBomId As String
Dim tyBomId As String


If Trim(DcbProductName.Text) = "" Or Trim(TxtBom1.Text) = "" Or Trim(DcbTYBomName.Text) = "" Then

    MsgBox "资料没有选全，请确认！", vbInformation, "友情提示"

    Exit Sub
Else

productNameTemp = Trim(DcbProductName.Text)
childBomId = Trim(TxtBom1.Text)
tyBomId = Left(Trim(DcbTYBomName.Text), InStr(Trim(DcbTYBomName.Text), "/") - 1)


End If

'判断是否已存在，如果存在，则不添加
If JudgeTSVBomAdd(productNameTemp) = True Then

    MsgBox "已存在，不用添加!", vbInformation, "友情提示"

Else



'sqlTemp = "insert into [erpdata].[dbo].[TSVtblBomSetup](ProductName,childBomNameID,Flag,CreateDate) values('" & productNameTemp & "','" & childBomId & "','Y',GETDATE())"
'AddSql2 (sqlTemp)

sqlTemp2 = "insert into [erpdata].[dbo].[TSVtblBomSetup](ProductName,childBomNameID,Flag,CreateDate) values('" & productNameTemp & "','" & tyBomId & "','Y',GETDATE())"
AddSql2 (sqlTemp2)


 MsgBox "添加成功!", vbInformation, "友情提示"

GetBomData
End If



End Sub

Private Sub DcbProductName_Change()
Dim productTemp As String
productTemp = Trim(DcbProductName.Text)

'根据成品料号，查询出成品Bom


  Set oiRS = GetProductChildBom(productTemp)

    If (oiRS.RecordCount > 0) Then

        TxtBom1.Text = Trim(oiRS.Fields("材料规范编号").Value)
        TxtBom2.Text = Trim(oiRS.Fields("物料编号").Value)

    End If


End Sub


Private Sub DcbProductName_Click()

Dim productTemp As String
productTemp = Trim(DcbProductName.Text)

'根据成品料号，查询出成品Bom


  Set oiRS = GetProductChildBom(productTemp)

    If (oiRS.RecordCount > 0) Then

        TxtBom1.Text = Trim(oiRS.Fields("材料规范编号").Value)
        TxtBom2.Text = Trim(oiRS.Fields("物料编号").Value)

    End If

IniProduct2
End Sub

Private Sub Form_Load()
IniFpsBom
GetBomData

IniProduct
IniProduct2

End Sub




Private Sub IniProduct()
Set mainItemRS1 = GetBomProductName()
'Set DcbProductName.RowSource = mainItemRS1
'DcbProductName.ListField = mainItemRS1("productname").Name
'DcbProductName.BoundColumn = mainItemRS1("PID").Name


 If Not mainItemRS1.EOF Then
      mainItemRS1.MoveFirst
    While Not mainItemRS1.EOF
        With DcbProductName
            .AddItem Trim(mainItemRS1!productName)
        End With
        mainItemRS1.MoveNext
       Wend

 End If





End Sub

Private Sub IniProduct2()
Set mainItemRS = GetBomTYName()
Set DcbTYBomName.RowSource = mainItemRS
DcbTYBomName.ListField = mainItemRS("productname").Name
DcbTYBomName.BoundColumn = mainItemRS("PID").Name


End Sub




Private Sub GetBomData()
'明细数据

Set bomRS = GetFpsBomSetUp()
If bomRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
End If

With fps(0)
        .MaxRows = 0
        If bomRS.RecordCount > 0 Then
            Set .DataSource = bomRS
        End If
End With

End Sub









Private Sub IniFpsBom()
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
        

        .SetText E_FPS0.E_id, 0, "序号"
        .SetText E_FPS0.E_PRODUCT, 0, "成品料号"
        .SetText E_FPS0.E_BomID, 0, "Bom编号"
        .SetText E_FPS0.E_BomName, 0, "Bom料号"
        .SetText E_FPS0.E_BomDate, 0, "建立日期"
        

        .ColWidth(E_FPS0.E_id) = 10
        .ColWidth(E_FPS0.E_PRODUCT) = 12
        .ColWidth(E_FPS0.E_BomID) = 12
        .ColWidth(E_FPS0.E_BomName) = 12
        .ColWidth(E_FPS0.E_BomDate) = 18
        
       

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        
        .ReDraw = True
    End With
    
    
    

End Sub

