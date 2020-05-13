VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmGCDie 
   Caption         =   "客户机种Die数设定"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13410
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
   ScaleHeight     =   7260
   ScaleWidth      =   13410
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "信息录入"
      Height          =   2535
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.TextBox TxtProductNew 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   1320
         Width           =   5175
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "新增"
         Height          =   360
         Left            =   2520
         TabIndex        =   5
         Top             =   1920
         Width           =   990
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改"
         Height          =   360
         Left            =   4080
         TabIndex        =   4
         Top             =   1920
         Width           =   990
      End
      Begin VB.CommandButton Command3 
         Caption         =   "取消"
         Height          =   360
         Left            =   5640
         TabIndex        =   3
         Top             =   1920
         Width           =   990
      End
      Begin VB.TextBox TxtProduct 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   840
         Width           =   5175
      End
      Begin MSDataListLib.DataCombo CmbCustomer 
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码："
         Height          =   195
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Good Die Qty："
         Height          =   195
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label LblProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Device："
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Top             =   960
         Width           =   1395
      End
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   4335
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   2760
      Width           =   11895
      _Version        =   196613
      _ExtentX        =   20981
      _ExtentY        =   7646
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
      SpreadDesigner  =   "FrmGCDie.frx":0000
      TextTip         =   2
   End
End
Attribute VB_Name = "FrmGCDie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
'    E_ID = 1                 'id
    E_Product = 1             '产品型号
    E_ProductNew              '成品料号
    E_TestNo                  '测试版本号
   
    E_End
    
End Enum

Private Sub CmdAdd_Click()
'新增
Dim customerTemp As String
Dim tempProduct As String
Dim dieTemp As Long
Dim tempTestNo As String


Dim sqlTemp As String

customerTemp = UCase(Trim(CmbCustomer.Text))
tempProduct = UCase(Trim(TxtProduct.Text))
dieTemp = CLng(UCase(Trim(TxtProductNew.Text)))



'判断是否已输入
 If customerTemp = "" Or tempProduct = "" Or dieTemp < 0 Then
    MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
    Exit Sub
 
 End If


 
sqlTemp = " insert into tblCustomerDieQty (CustomerName,CustomerPT,DieQty,createdby,createddate,flag ) values  ('" & customerTemp & "','" & tempProduct & "'," & dieTemp & ",'Auto',sysdate,'Y')"
AddSql (sqlTemp)


 MsgBox "添加成功!", vbInformation, "友情提示"
 
ShowData_Where



End Sub

Private Sub Command2_Click()
'修改




Dim customerTemp As String
Dim tempProduct As String
Dim dieTemp As Long
Dim tempTestNo As String


Dim sqlTemp As String

customerTemp = UCase(Trim(CmbCustomer.Text))
tempProduct = UCase(Trim(TxtProduct.Text))
dieTemp = CLng(UCase(Trim(TxtProductNew.Text)))



'判断是否已输入
 If customerTemp = "" Or tempProduct = "" Or dieTemp < 0 Then
    MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
    Exit Sub
 
 End If
 

'判断输入的Lot号，是否存在于BC表中
If (Not JudGCDieNoExist(customerTemp, tempProduct)) Then
   MsgBox "这笔：" & tempProduct & " 不存在，无需修改！"
Exit Sub

End If


Call DelGCDieQty(customerTemp, tempProduct, dieTemp)
ShowData_Where


End Sub

Private Sub Form_Load()


IniCustomerName


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
        

        .SetText E_FPS0.E_Product, 0, "客户代码"
        .SetText E_FPS0.E_ProductNew, 0, "客户机种"
        .SetText E_FPS0.E_TestNo, 0, "DieQty"

        
        .ColWidth(E_FPS0.E_Product) = 20
        .ColWidth(E_FPS0.E_ProductNew) = 30
        .ColWidth(E_FPS0.E_TestNo) = 40

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        

        
        
        .ReDraw = True
    End With
    
    ShowData_Where
    
    
End Sub

Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CmbCustomer.RowSource = mainItemRS
CmbCustomer.ListField = mainItemRS("productname").Name
CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub

Private Sub ShowData_Where()
Set reportRS = GetGCDieQty()

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub



