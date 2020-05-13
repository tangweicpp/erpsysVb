VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmTestNo2 
   Caption         =   "测试版本与OI设定"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16485
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
   ScaleHeight     =   7860
   ScaleWidth      =   16485
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtProductNew 
      Height          =   375
      Left            =   8160
      TabIndex        =   29
      Top             =   240
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmTestNo2.frx":0000
      Left            =   360
      List            =   "FrmTestNo2.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox TxtOldTestNo 
      Height          =   375
      Left            =   8160
      TabIndex        =   26
      Top             =   2520
      Width           =   3615
   End
   Begin VB.TextBox TxtAttri3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   12960
      TabIndex        =   23
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox TxtAttri2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   12960
      TabIndex        =   21
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox TxtAttri 
      Enabled         =   0   'False
      Height          =   375
      Left            =   12960
      TabIndex        =   19
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   360
      Left            =   8760
      TabIndex        =   18
      Top             =   3120
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "修改"
      Height          =   360
      Left            =   6480
      TabIndex        =   17
      Top             =   3120
      Width           =   990
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "新增"
      Height          =   360
      Left            =   4200
      TabIndex        =   16
      Top             =   3120
      Width           =   990
   End
   Begin VB.TextBox txtKey3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox txtValue3 
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox txtKey2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox txtValue2 
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox TxtTestNo 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox TxtProduct 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   3495
   End
   Begin VB.TextBox txtKey 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   4335
      Index           =   0
      Left            =   600
      TabIndex        =   25
      Top             =   3480
      Width           =   15615
      _Version        =   524288
      _ExtentX        =   27543
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
      SpreadDesigner  =   "FrmTestNo2.frx":001C
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "成品料号："
      Height          =   195
      Left            =   7200
      TabIndex        =   30
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "旧版本号："
      Height          =   195
      Left            =   7200
      TabIndex        =   27
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注："
      Height          =   195
      Left            =   12360
      TabIndex        =   24
      Top             =   2160
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注："
      Height          =   195
      Left            =   12360
      TabIndex        =   22
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注："
      Height          =   195
      Left            =   12360
      TabIndex        =   20
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户OI Excel字段名："
      Height          =   315
      Left            =   840
      TabIndex        =   15
      Top             =   2160
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户OI Excel字段值："
      Height          =   195
      Left            =   6480
      TabIndex        =   14
      Top             =   2160
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户OI Excel字段名："
      Height          =   315
      Left            =   840
      TabIndex        =   11
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户OI Excel字段值："
      Height          =   195
      Left            =   6480
      TabIndex        =   10
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Label LblTestNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "测试版本号："
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label LblProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "产品型号："
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   900
   End
   Begin VB.Label LblKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户OI Excel字段名："
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   1680
   End
   Begin VB.Label LblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户OI Excel字段值："
      Height          =   195
      Left            =   6480
      TabIndex        =   2
      Top             =   960
      Width           =   1680
   End
End
Attribute VB_Name = "FrmTestNo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail

    '    E_ID = 1                 'id
    E_PRODUCT = 1             '产品号
    E_ProductNew              '成品料号
     
    E_TestNo                  '测试版本号
    
    E_Key1                    'Key1
    E_Value1                  'Value1
    E_Attr1                   '备注1
    
    E_Key2                    'Key2
    E_Value2                  'Value2
    E_Attr2                   '备注2
    
    E_Key3                    'Key3
    E_Value3                  'Value3
    E_Attr3                   '备注3
    
    E_End
    
End Enum

Private Sub CmdAdd_Click()

    '新增
    Dim tempProduct    As String

    Dim tempProductNew As String

    Dim tempTestNo     As String

    Dim tempKey1       As String

    Dim tempValue1     As String

    Dim otherValue1    As String

    Dim tempKey2       As String

    Dim tempValue2     As String

    Dim otherValue2    As String

    Dim tempKey3       As String

    Dim tempValue3     As String

    Dim otherValue3    As String

    Dim sqlTemp        As String

    tempProduct = UCase(Trim(TxtProduct.Text))
    tempProductNew = UCase(Trim(TxtProductNew.Text))
    tempTestNo = UCase(Trim(TxtTestNo.Text))

    tempKey1 = UCase(Trim(txtKey.Text))
    tempValue1 = Trim(txtValue.Text)
    otherValue1 = Trim(TxtAttri.Text)

    tempKey2 = UCase(Trim(txtKey2.Text))
    tempValue2 = Trim(txtValue2.Text)
    otherValue2 = Trim(TxtAttri2.Text)

    tempKey3 = UCase(Trim(txtKey3.Text))
    tempValue3 = Trim(txtValue3.Text)
    otherValue3 = Trim(TxtAttri3.Text)

    '判断是否已输入
    If tempProduct = "" Or tempProductNew = "" Or tempTestNo = "" Or tempKey1 = "" Or tempValue1 = "" Or otherValue1 = "" Then
        MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
        Exit Sub
 
    End If
 
    sqlTemp = " insert into tblTestNo2(productname,testedition,FIELDNAME1,FIELDVALUE1,Remark1,FIELDNAME2,FIELDVALUE2,Remark2,FIELDNAME3,FIELDVALUE3,Remark3,createdby,createddate,flag,productnamenew ) " & " values  ('" & tempProduct & "','" & tempTestNo & "','" & tempKey1 & "','" & tempValue1 & "','" & otherValue1 & "','" & tempKey2 & "','" & tempValue2 & "','" & otherValue2 & "','" & tempKey3 & "','" & tempValue3 & "','" & otherValue3 & "','Auto',sysdate,'Y','" & tempProductNew & "')"
    AddSql (sqlTemp)

    MsgBox "添加成功!", vbInformation, "友情提示"
 
    'ShowData_Where

End Sub

Private Sub Combo1_Change()

    If Combo1.Text = "规则1" Then
        txtKey.Text = "DESIGN_ID"
        TxtAttri.Text = "R栏"
    
        txtKey2.Text = "MPN_DESC"
        TxtAttri2.Text = "J栏"
    
        txtKey3.Text = "ASSY_SERIAL_TYPE"
        TxtAttri3.Text = "AJ栏"
    ElseIf Combo1.Text = "规则2" Then

        txtKey.Text = "DESIGN_ID"
        TxtAttri.Text = "R栏"
    
        txtKey2.Text = "PROBE_SHIP_PART_TYPE"
        TxtAttri2.Text = "M栏"
    
        txtKey3.Text = "ASSY_SERIAL_TYPE"
        TxtAttri3.Text = "AJ栏"

    End If

End Sub

Private Sub Combo1_Click()

    If Combo1.Text = "规则1" Then
        txtKey.Text = "DESIGN_ID"
        TxtAttri.Text = "R栏"
    
        txtKey2.Text = "MPN_DESC"
        TxtAttri2.Text = "J栏"
    
        txtKey3.Text = "ASSY_SERIAL_TYPE"
        TxtAttri3.Text = "AJ栏"
    ElseIf Combo1.Text = "规则2" Then

        txtKey.Text = "DESIGN_ID"
        TxtAttri.Text = "R栏"
    
        txtKey2.Text = "probe_ship_part_type"
        TxtAttri2.Text = "M栏"
    
        txtKey3.Text = "ASSY_SERIAL_TYPE"
        TxtAttri3.Text = "AJ栏"

    End If

End Sub

Private Sub Command2_Click()
    '修改

    Dim tempProduct    As String

    Dim tempProductNew As String

    Dim tempTestNo     As String

    Dim tempOldtestNo  As String

    tempProduct = UCase(Trim(TxtProduct.Text))
    tempProductNew = UCase(Trim(TxtProductNew.Text))
    tempTestNo = UCase(Trim(TxtTestNo.Text))
    tempOldtestNo = UCase(Trim(TxtOldTestNo.Text))

    '判断是否已输入
    If tempProduct = "" Or tempProductNew = "" Or tempTestNo = "" Or tempOldtestNo = "" Then
        MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
        Exit Sub
 
    End If

    '判断输入的Lot号，是否存在于BC表中
    If (Not JudgetestNoExist2(tempProduct, tempOldtestNo, tempProductNew)) Then
        MsgBox "这笔：" & tempProduct & " 不存在，无需修改！"
        Exit Sub

    End If

    Call DeltestNo2(tempProduct, tempTestNo, tempOldtestNo, tempProductNew)
    ShowData_Where

End Sub

Private Sub Form_Activate()
    TxtProduct.SetFocus

End Sub

Private Sub Form_Load()
    txtKey.Text = "DESIGN_ID"
    TxtAttri.Text = "R栏"

    txtKey2.Text = "MPN_DESC"
    TxtAttri2.Text = "J栏"

    txtKey3.Text = "ASSY_SERIAL_TYPE"
    TxtAttri3.Text = "AJ栏"

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

        .SetText E_FPS0.E_PRODUCT, 0, "产品型号"
        .SetText E_FPS0.E_ProductNew, 0, "成品料号"
              
        .SetText E_FPS0.E_TestNo, 0, "测试版本号"
        
        .SetText E_FPS0.E_Key1, 0, "字段名1"
        .SetText E_FPS0.E_Value1, 0, "字段值1"
        .SetText E_FPS0.E_Attr1, 0, "备注1"
        
        .SetText E_FPS0.E_Key2, 0, "字段名2"
        .SetText E_FPS0.E_Value2, 0, "字段值2"
        .SetText E_FPS0.E_Attr2, 0, "备注2"
        
        .SetText E_FPS0.E_Key3, 0, "字段名3"
        .SetText E_FPS0.E_Value3, 0, "字段值3"
        .SetText E_FPS0.E_Attr3, 0, "备注3"
        
        .ColWidth(E_FPS0.E_PRODUCT) = 8
        .ColWidth(E_FPS0.E_ProductNew) = 12
        .ColWidth(E_FPS0.E_TestNo) = 18
        
        .ColWidth(E_FPS0.E_Key1) = 13
        .ColWidth(E_FPS0.E_Value1) = 8
        .ColWidth(E_FPS0.E_Attr1) = 6
                        
        .ColWidth(E_FPS0.E_Key2) = 13
        .ColWidth(E_FPS0.E_Value2) = 8
        .ColWidth(E_FPS0.E_Attr2) = 6
        
        .ColWidth(E_FPS0.E_Key3) = 13
        .ColWidth(E_FPS0.E_Value3) = 8
        .ColWidth(E_FPS0.E_Attr3) = 6

        .RowHeight(0) = 20
        .RowHeight(-1) = 15

        .ReDraw = True

    End With

    ShowData_Where
    
End Sub

Private Sub ShowData_Where()
    Set reportRS = GettestNo2()

    With fps(0)
        .MaxRows = 0

        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If

    End With

End Sub

Private Sub TxtProductNew_KeyPress(KeyAscii As Integer)

    Dim tempProductName As String

    If KeyAscii = 13 Then
        '查询测试版本号
        tempProductName = UCase(Trim(TxtProductNew.Text))

        If tempProductName = "" Then
    
            MsgBox "请输入成品料号！", vbInformation, "友情提示"
     
        Else
            IniProductTestNo tempProductName
    
        End If

    End If

End Sub

Private Sub IniProductTestNo(productNameTemp As String)

    Dim testval As String

    Set mainItemRS = GetMainItemProduct(productNameTemp)
    'Set DCbMainItem.RowSource = mainItemRS
    testval = mainItemRS("typename").Name

    'DCbMainItem.ListField = mainItemRS("typename").Name
    'DCbMainItem.BoundColumn = mainItemRS("id").Name

    TxtTestNo.Text = mainItemRS.Fields(1).Value

End Sub

