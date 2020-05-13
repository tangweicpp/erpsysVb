VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmTestNo 
   Caption         =   "测试版本设定"
   ClientHeight    =   9615
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
   ScaleHeight     =   9615
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
         TabIndex        =   9
         Top             =   960
         Width           =   5175
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "新增"
         Height          =   360
         Left            =   2520
         TabIndex        =   7
         Top             =   1920
         Width           =   990
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改"
         Height          =   360
         Left            =   4080
         TabIndex        =   6
         Top             =   1920
         Width           =   990
      End
      Begin VB.CommandButton Command3 
         Caption         =   "取消"
         Height          =   360
         Left            =   5640
         TabIndex        =   5
         Top             =   1920
         Width           =   990
      End
      Begin VB.TextBox TxtTestNo 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   1440
         Width           =   5175
      End
      Begin VB.TextBox TxtProduct 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "备注：当工程文档提出，要修改测试版本号时，只需在此界面修改，其它界面不用再次修改！"
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   7920
         TabIndex        =   11
         Top             =   1920
         Width           =   3780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "成品料号："
         Height          =   195
         Left            =   1200
         TabIndex        =   10
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label LblTestNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "测试版本号："
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label LblProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "产品型号："
         Height          =   195
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   900
      End
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   6735
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   2760
      Width           =   11895
      _Version        =   196613
      _ExtentX        =   20981
      _ExtentY        =   11880
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
      SpreadDesigner  =   "FrmTestNo.frx":0000
      TextTip         =   2
   End
End
Attribute VB_Name = "FrmTestNo"
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
Dim tempProduct As String
Dim tempProductNew As String
Dim tempTestNo As String


Dim sqlTemp As String

tempProduct = UCase(Trim(TxtProduct.Text))
tempProductNew = UCase(Trim(TxtProductNew.Text))
tempTestNo = UCase(Trim(TxtTestNo.Text))


'判断是否已输入
 If tempProduct = "" Or tempTestNo = "" Or tempProductNew = "" Then
    MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
    Exit Sub
 
 End If


 
sqlTemp = " insert into tblTestNo(productname,testedition,createdby,createddate,flag,productnamenew ) values  ('" & tempProduct & "','" & tempTestNo & "','Auto',sysdate,'Y','" & tempProductNew & "')"
AddSql (sqlTemp)


'2013-03-26 jiayun add 流程卡用
sqlTemp = "insert into TSVCard_EDT( id,productname,testedition,createdby,createddate,flag,productnamenew) values (RCardTestVersionId.Nextval,'" & tempProduct & "','" & tempTestNo & "','Auto',sysdate,'Y','" & tempProductNew & "')"
AddSql (sqlTemp)




 MsgBox "添加成功!", vbInformation, "友情提示"
 
ShowData_Where



End Sub

Private Sub Command2_Click()
'修改

Dim tempProduct As String
Dim tempTestNo As String
Dim tempProductNew As String


tempProduct = UCase(Trim(TxtProduct.Text))
tempProductNew = UCase(Trim(TxtProductNew.Text))
tempTestNo = UCase(Trim(TxtTestNo.Text))

'判断是否已输入
 If tempProduct = "" Or tempTestNo = "" Or tempProductNew = "" Then
    MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
    Exit Sub
 
 End If
 

'判断输入的Lot号，是否存在于BC表中
If (Not JudgetestNoExist(tempProduct, tempProductNew)) Then
   MsgBox "这笔：" & tempProduct & " 不存在，无需修改！"
Exit Sub

End If


Call DeltestNo(tempProduct, tempTestNo, tempProductNew)
ShowData_Where


End Sub

Private Sub Form_Load()
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
        

        .SetText E_FPS0.E_Product, 0, "产品型号"
        .SetText E_FPS0.E_ProductNew, 0, "成品料号"
        .SetText E_FPS0.E_TestNo, 0, "测试版本号"

        
        .ColWidth(E_FPS0.E_Product) = 20
        .ColWidth(E_FPS0.E_ProductNew) = 30
        .ColWidth(E_FPS0.E_TestNo) = 40

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        

        
        
        .ReDraw = True
    End With
    
    ShowData_Where
    
    
End Sub


Private Sub ShowData_Where()
Set reportRS = GettestNo()

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub



