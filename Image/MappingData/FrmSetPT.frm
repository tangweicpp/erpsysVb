VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmSetPT 
   Caption         =   "AA客户料号的设定"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13725
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
   ScaleHeight     =   7650
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox DCbMainItem 
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "新增"
      Height          =   360
      Left            =   4920
      TabIndex        =   8
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton Command3 
      Caption         =   "修改"
      Height          =   360
      Left            =   6960
      TabIndex        =   7
      Top             =   1560
      Width           =   990
   End
   Begin VB.ComboBox CombTray 
      Height          =   315
      ItemData        =   "FrmSetPT.frx":0000
      Left            =   7200
      List            =   "FrmSetPT.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   3735
   End
   Begin VB.ComboBox Combpf 
      Height          =   315
      ItemData        =   "FrmSetPT.frx":0025
      Left            =   2040
      List            =   "FrmSetPT.frx":002F
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox TxtPT 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   5175
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   2280
      Width           =   12495
      _Version        =   196613
      _ExtentX        =   22040
      _ExtentY        =   9128
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
      SpreadDesigner  =   "FrmSetPT.frx":0041
      TextTip         =   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "测试版本号："
      Height          =   195
      Left            =   6120
      TabIndex        =   6
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tray 标记："
      Height          =   195
      Left            =   6240
      TabIndex        =   4
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "贴膜标记："
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label LblPT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "最终成本料号："
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1260
   End
End
Attribute VB_Name = "FrmSetPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail
'    E_ID = 1                 'id
    E_Product = 1             '成本料号
    E_Pf                      'PF标记
    E_Tray                    'Tray标记
    E_TestNo                  '测试版本号
    E_End
    
End Enum

Dim mainItemRS As New ADODB.Recordset
Dim reportRS As New ADODB.Recordset


Private Sub IniProductTestNo(productNameTemp As String)
Dim testval As String

Set mainItemRS = GetMainItemProduct(productNameTemp)
'Set DCbMainItem.RowSource = mainItemRS
testval = mainItemRS("typename").Name

'DCbMainItem.ListField = mainItemRS("typename").Name
'DCbMainItem.BoundColumn = mainItemRS("id").Name

DCbMainItem.Text = mainItemRS.fields(1).Value

End Sub


'Private Sub IniTestNo()
'Set mainItemRS = GetMainItem()
'Set DCbMainItem.RowSource = mainItemRS
'DCbMainItem.ListField = mainItemRS("typename").Name
'DCbMainItem.BoundColumn = mainItemRS("id").Name
'
'End Sub

Private Sub CmdAdd_Click()
'新增
Dim tempProductName As String
Dim tempPf As String
Dim tempTray As String
Dim tempTestNo As String

Dim sqlTemp As String

tempProductName = UCase(Trim(TxtPT.Text))
tempPf = Combpf.Text
tempTray = CombTray.Text
tempTestNo = DCbMainItem.Text



'判断是否已输入
 If tempProductName = "" Or tempPf = "" Or tempTray = "" Or tempTestNo = "" Then
    MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
    Exit Sub
 
 End If


 
sqlTemp = "insert into TBLSETPT(productName,pfStaus,trayStaus,testNo,flag,createby,createdate) values ('" & tempProductName & "','" & tempPf & "','" & tempTray & "','" & tempTestNo & "','Y','Auto',sysdate)"
AddSql (sqlTemp)

 MsgBox "添加成功!", vbInformation, "友情提示"
 
ShowData_Where

End Sub

Private Sub Command3_Click()
'修改

Dim tempProductName As String
Dim tempPf As String
Dim tempTray As String
Dim tempTestNo As String

Dim sqlTemp As String

tempProductName = UCase(Trim(TxtPT.Text))
tempPf = Combpf.Text
tempTray = CombTray.Text
tempTestNo = DCbMainItem.Text



'判断是否已输入
 If tempProductName = "" Or tempPf = "" Or tempTray = "" Or tempTestNo = "" Then
    MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
    Exit Sub
 
 End If
 
 '判断料号是否存在
If (Not JudgePtExist(tempProductName)) Then
   MsgBox "这笔：" & tempProductName & " 不存在，无需修改！"
Exit Sub

End If


sqlTemp = " update TBLSETPT set pfStaus='" & tempPf & "',trayStaus='" & tempTray & "',testNo='" & tempTestNo & "',lastupdateby='Auto',lastupdatedate=sysdate where productName='" & tempProductName & "' and flag='Y' "

AddSql (sqlTemp)

 MsgBox "修改成功!", vbInformation, "友情提示"
 
ShowData_Where



End Sub

Private Sub Form_Load()
'IniTestNo

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
        
        
        .SetText E_FPS0.E_Product, 0, "成本料号"
        .SetText E_FPS0.E_Pf, 0, "PF标记"
        .SetText E_FPS0.E_Tray, 0, "Tray标记"
        .SetText E_FPS0.E_TestNo, 0, "测试版本号"
    
        
        .ColWidth(E_FPS0.E_Product) = 20
        .ColWidth(E_FPS0.E_Pf) = 20
        .ColWidth(E_FPS0.E_Tray) = 20
        .ColWidth(E_FPS0.E_TestNo) = 25
       
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        

        
        
        .ReDraw = True
    End With
    
    
ShowData_Where

End Sub


Private Sub ShowData_Where()
Set reportRS = GetptData()

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub


Private Sub TxtPT_KeyPress(KeyAscii As Integer)
Dim tempProductName As String

If KeyAscii = 13 Then
'查询测试版本号
tempProductName = UCase(Trim(TxtPT.Text))
    If tempProductName = "" Then
    
     MsgBox "请输入成品料号！", vbInformation, "友情提示"
     
    Else
    IniProductTestNo tempProductName
    
    
    
    End If

End If

End Sub
