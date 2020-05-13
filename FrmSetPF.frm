VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmSetPF 
   Caption         =   "贴膜参数设定"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13530
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
   ScaleHeight     =   7815
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   360
      Left            =   4920
      TabIndex        =   10
      Top             =   2880
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "修改"
      Height          =   360
      Left            =   3360
      TabIndex        =   9
      Top             =   2880
      Width           =   990
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "新增"
      Height          =   360
      Left            =   1800
      TabIndex        =   8
      Top             =   2880
      Width           =   990
   End
   Begin VB.TextBox TxtAttri 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1920
      Width           =   3495
   End
   Begin VB.ComboBox CombMo 
      Height          =   315
      ItemData        =   "FrmSetPF.frx":0000
      Left            =   2400
      List            =   "FrmSetPF.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txtKey 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   2415
      Index           =   0
      Left            =   1320
      TabIndex        =   11
      Top             =   3720
      Width           =   9015
      _Version        =   524288
      _ExtentX        =   15901
      _ExtentY        =   4260
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
      SpreadDesigner  =   "FrmSetPF.frx":001C
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注："
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "是否贴膜："
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label LblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户OI Excel字段值："
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   1680
   End
   Begin VB.Label LblKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户OI Excel字段名："
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1680
   End
End
Attribute VB_Name = "FrmSetPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail

    '    E_ID = 1                 'id
    E_Key = 1                'Key
    E_Value                  'Value
    E_getValue               'getValue
    E_otherValue             '备注
    E_End
    
End Enum

Dim reportRS As New ADODB.Recordset

Private Sub CmdAdd_Click()

    '新增
    Dim tempKey    As String

    Dim tempValue  As String

    Dim getValue   As String

    Dim otherValue As String

    Dim sqlTemp    As String

    tempKey = UCase(Trim(txtKey.Text))
    tempValue = Trim(txtValue.Text)
    getValue = CombMo.Text
    otherValue = Trim(TxtAttri.Text)

    '判断是否已输入
    If tempKey = "" Or getValue = "" Then
        MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
        Exit Sub
 
    End If
 
    sqlTemp = " insert into  tblsetpf(fieldName,fieldValue,resultValue,other,flag,createby,createdate) values ('" & tempKey & "','" & tempValue & "','" & getValue & "','" & otherValue & "','Y','Auto',sysdate)"
    AddSql (sqlTemp)

    MsgBox "添加成功!", vbInformation, "友情提示"
 
    ShowData_Where

End Sub

Private Sub Command2_Click()
    '修改

    Dim tempKey    As String

    Dim tempValue  As String

    Dim getValue   As String

    Dim otherValue As String

    Dim sqlTemp    As String

    tempKey = UCase(Trim(txtKey.Text))
    tempValue = Trim(txtValue.Text)
    getValue = CombMo.Text
    otherValue = Trim(TxtAttri.Text)

    '判断是否已输入
    If tempKey = "" Or getValue = "" Then
        MsgBox "输入完整后，再提交！", vbInformation, "友情提示"
        Exit Sub
 
    End If

    Call UpdatePfData(tempValue, getValue)
    ShowData_Where

End Sub

Private Sub Form_Load()

    txtKey.Text = "PROTECTIVE_FILM_APLD"
    TxtAttri.Text = "BB栏"

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
        
        .SetText E_FPS0.E_Key, 0, "字段名"
        .SetText E_FPS0.E_Value, 0, "字段值"
        .SetText E_FPS0.E_getValue, 0, "是否贴膜"
        .SetText E_FPS0.E_otherValue, 0, "备注"
        
        .ColWidth(E_FPS0.E_Key) = 20
        .ColWidth(E_FPS0.E_Value) = 15
        .ColWidth(E_FPS0.E_getValue) = 15
        .ColWidth(E_FPS0.E_otherValue) = 25

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .ReDraw = True

    End With
    
    ShowData_Where

End Sub

Private Sub ShowData_Where()
    Set reportRS = GetpfData()

    With fps(0)
        .MaxRows = 0

        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If

    End With

End Sub

