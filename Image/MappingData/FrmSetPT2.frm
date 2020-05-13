VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmSetPT2 
   Caption         =   "客户料号与厂内料号对应关系的设定(除AA)"
   ClientHeight    =   9630
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   11640
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox TxtQtechPT 
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "新增"
      Height          =   360
      Left            =   3960
      TabIndex        =   5
      Top             =   840
      Width           =   990
   End
   Begin VB.CommandButton Command3 
      Caption         =   "修改"
      Height          =   360
      Left            =   6840
      TabIndex        =   4
      Top             =   840
      Width           =   990
   End
   Begin VB.TextBox TxtCustomerPT 
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   8055
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   12495
      _Version        =   196613
      _ExtentX        =   22040
      _ExtentY        =   14208
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
      SpreadDesigner  =   "FrmSetPT2.frx":0000
      TextTip         =   2
   End
   Begin MSDataListLib.DataCombo CmbCustomer 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "二级代码："
      Height          =   195
      Left            =   10800
      TabIndex        =   9
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "厂内料号："
      Height          =   195
      Left            =   6960
      TabIndex        =   3
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户代码："
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   900
   End
   Begin VB.Label LblPT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户料号："
      Height          =   195
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "FrmSetPT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail
'    E_ID = 1                 'id
    E_Product = 1             '客户代码
    E_Pf                      '客户料号
    E_Tray                    '成品料号
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

TxtQtechPT.Text = mainItemRS.fields(1).Value

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

tempProductName = UCase(Trim(CmbCustomer.Text))
tempPf = TxtCustomerPT.Text
tempTray = TxtQtechPT.Text
secondCode = Trim(Text1.Text)



'判断是否已输入
 If tempProductName = "" Or tempPf = "" Or tempTray = "" Then
    MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
    Exit Sub
 
 End If


 
sqlTemp = "insert into TBLSETPT(customercode,customerpt,qtechpt,flag,createby,createdate) values ('" & tempProductName & "','" & tempPf & "','" & tempTray & "','Y','Auto',sysdate)"
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

tempProductName = UCase(Trim(TxtCustomerPT.Text))
tempPf = ComCustomer.Text
tempTray = CombTray.Text
tempTestNo = TxtQtechPT.Text



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

Private Sub Form_Activate()
CmbCustomer.SetFocus



End Sub

Private Sub Form_Load()
'IniTestNo

IniCustomerName

 With Fps(0)
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
        .SetText E_FPS0.E_Pf, 0, "客户料号"
        .SetText E_FPS0.E_Tray, 0, "成品料号"

    
        
        .ColWidth(E_FPS0.E_Product) = 20
        .ColWidth(E_FPS0.E_Pf) = 20
        .ColWidth(E_FPS0.E_Tray) = 20

       
        

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
Set reportRS = GetptOtherCustPT()


With Fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub


Private Sub Label2_Click()

End Sub

Private Sub TxtCustomerPT_KeyPress(KeyAscii As Integer)
Dim tempProductName As String

If KeyAscii = 13 Then
'查询测试版本号
tempProductName = UCase(Trim(TxtCustomerPT.Text))
    If tempProductName = "" Then
    
     MsgBox "请输入成品料号！", vbInformation, "友情提示"
     
    Else
    IniProductTestNo tempProductName
    
    
    
    End If

End If

End Sub
