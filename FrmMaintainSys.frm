VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmMaintainSys 
   Caption         =   "数据综合维护平台"
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12735
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
   ScaleHeight     =   11085
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame FrameToolBar 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   -120
      Width           =   12735
      Begin VB.CommandButton btnExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5760
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMaintainSys.frx":0000
         Picture         =   "FrmMaintainSys.frx":3072
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "新增"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton btnExport 
         BackColor       =   &H00E0E0E0&
         Caption         =   "导出"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4800
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMaintainSys.frx":5814
         Picture         =   "FrmMaintainSys.frx":8886
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton btnImport 
         BackColor       =   &H00E0E0E0&
         Caption         =   "导入"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3840
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMaintainSys.frx":A9B0
         Picture         =   "FrmMaintainSys.frx":DA22
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton btnDelete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "删除"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2880
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMaintainSys.frx":F71C
         Picture         =   "FrmMaintainSys.frx":1278E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton btnModify 
         BackColor       =   &H00E0E0E0&
         Caption         =   "修改"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMaintainSys.frx":14F30
         Picture         =   "FrmMaintainSys.frx":17FA2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton btnNew 
         BackColor       =   &H00E0E0E0&
         Caption         =   "新增"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   960
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMaintainSys.frx":1B014
         Picture         =   "FrmMaintainSys.frx":1E086
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton btnQuery 
         BackColor       =   &H00E0E0E0&
         Caption         =   "查询"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         MaskColor       =   &H008080FF&
         Picture         =   "FrmMaintainSys.frx":201B0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9855
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   12735
      Begin VB.TextBox txtModifyValue1 
         Height          =   375
         Left            =   9000
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "全选/反选"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtCol3Value 
         Height          =   375
         Left            =   5880
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtCol2Value 
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtCol1Value 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox cboMaintainType 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "FrmMaintainSys.frx":23222
         Left            =   1080
         List            =   "FrmMaintainSys.frx":23224
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   4455
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   8895
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   12375
         _Version        =   524288
         _ExtentX        =   21828
         _ExtentY        =   15690
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
         MaxCols         =   20
         MaxRows         =   0
         SpreadDesigner  =   "FrmMaintainSys.frx":23226
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lblModifyName1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ModifyName1"
         Height          =   195
         Left            =   9000
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblCol3Name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "name3"
         Height          =   195
         Left            =   5880
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblCol2Name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "name2"
         Height          =   195
         Left            =   3000
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblCol1Name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "name1"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblMaintainType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "维护类型"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmMaintainSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type T_SQL

    T_BODY As String
    T_COLUMN1 As String
    T_NAME1 As String
    T_VALUE1 As String
    T_COLUMN2 As String
    T_NAME2 As String
    T_VALUE2 As String
    T_COLUMN3 As String
    T_NAME3 As String
    T_VALUE3 As String

End Type

Private Type T_SQL_UPDATE

    T_BODY As String
    T_BODY2 As String
    T_COLUMN1 As String
    T_NAME1 As String
    T_VALUE1 As String
    T_SET_DATA1 As String
    T_SET_DATA2 As String

End Type

Private TSQL        As T_SQL

Private TSQL_UPDATE As T_SQL_UPDATE

Private STR_QUERY   As String

Private Sub btnExit_Click()
Unload Me

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitMaintainType
' Description:       初始化维护类型
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/12/17-9:49:49
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitMaintainType()
Dim rs     As New ADODB.Recordset
Dim strSql As String

strSql = "select maintain_type from TBL_MAINTAIN_SYS where instr(AUTH_USER ,'" & gUserName & "') > 0"
Set rs = Get_OracleRs(strSql)
cboMaintainType.Clear
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cboMaintainType.AddItem Trim("" & rs!maintain_type)
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

Private Sub btnExport_Click()
If STR_QUERY = "" Then
    MsgBox "请先查询出数据,再点击导出按钮", vbInformation, "提示"
    Exit Sub

End If

Call ExporToExcel(STR_QUERY)

End Sub

Private Sub btnModify_Click()
Dim strSql       As String
Dim i            As Integer
Dim iRes         As Integer
Dim strEventDesc As String
Dim j            As Integer

With fps
    If .MaxRows = 0 Then
        MsgBox "请先查询出数据,否则无法修改", vbInformation, "提示"
        Exit Sub

    End If

End With

TSQL_UPDATE.T_VALUE1 = Trim(txtModifyValue1.Text)
If TSQL_UPDATE.T_NAME1 <> "" Then
    If TSQL_UPDATE.T_VALUE1 = "" Then
        MsgBox "请填写更改条件:" & TSQL_UPDATE.T_NAME1 & "的值", vbInformation, "提示"
        Exit Sub

    End If

    '填写理由
    strGReason = ""
    DlgEventReason.Show 1
    If strGReason = "" Then
        MsgBox "必须填写理由,否则不可修改", vbInformation, "提示"
        Exit Sub

    End If

    '记录
    strEventDesc = "修改出货地址记录"
    strEventID = SaveTblEventRec(E_TBL_EVENT.E_UPDATE, TSQL.T_VALUE1, strEventDesc, strGReason, "")
    If strEventID = "" Then
        MsgBox "修改事件未记录,无法修改", vbCritical, "提示"
        Exit Sub

    End If

    '    '备份
    '    strSql = "insert into ERPBASE..tbltoinrec_wafer_bak select * from ERPBASE..tbltoinrec_wafer where 入库单编号 = '" & strInRecID & "'  "
    '    AddSql2 (strSql)
    '    strNewID = strInRecID & "|" & strEventID
    '    strSql = "update ERPBASE..tbltoinrec_wafer_bak set 入库单编号 = '" & strNewID & "' where 入库单编号 = '" & strInRecID & "'  "
    '
    '    If AddSql2(strSql) > 0 Then
    '        MsgBox "数据已备份", vbInformation, "提示"
    '
    '    End If
    '
    '    '删除
    '    strSql = "delete from ERPBASE..tbltoinrec_wafer where 入库单编号 = '" & strInRecID & "'"
    '    If AddSql2(strSql) > 0 Then
    '        MsgBox "数据已删除", vbInformation, "提示"
    '
    '    End If
    '
    With fps

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .Value = 1 Then
                .Col = 2
                strSql = TSQL_UPDATE.T_BODY & .Text
                strSql = Replace(strSql, "%1", TSQL_UPDATE.T_VALUE1)
                AddSql (strSql)
                If TSQL_UPDATE.T_BODY2 <> "" Then
                    strSql = TSQL_UPDATE.T_BODY2 & .Text
                    strSql = Replace(strSql, "%1", TSQL_UPDATE.T_VALUE1)
                    AddSql2 (strSql)

                End If

                iRes = iRes + 1

            End If

        Next

    End With

End If

If iRes = 0 Then
    MsgBox "未更新", vbInformation, "提示"
Else
    MsgBox "更新了" & iRes & "笔数据", vbInformation, "提示"

End If

Call ShowData(STR_QUERY)

End Sub

Private Sub btnQuery_Click()
Dim strSql As String

TSQL.T_VALUE1 = Trim(txtCol1Value.Text)
TSQL.T_VALUE2 = Trim(txtCol2Value.Text)
TSQL.T_VALUE3 = Trim(txtCol3Value.Text)
If TSQL.T_COLUMN1 <> "" Then
    If TSQL.T_VALUE1 = "" Then
        MsgBox "请填写查询条件:" & TSQL.T_NAME1 & "的值", vbInformation, "提示"
        Exit Sub

    End If

    strSql = TSQL.T_BODY & " where 1= 1 and " & TSQL.T_COLUMN1 & " = " & " '" & TSQL.T_VALUE1 & "' "
ElseIf TSQL.T_COLUMN2 <> "" Then
    If TSQL.T_VALUE2 = "" Then
        MsgBox "请填写查询条件:" & TSQL.T_NAME2 & "的值", vbInformation, "提示"
        Exit Sub

    End If

    strSql = TSQL.T_BODY & " where 1= 1 and " & TSQL.T_COLUMN2 & " = " & " '" & TSQL.T_VALUE2 & "' "
ElseIf TSQL.T_COLUMN3 <> "" Then
    If TSQL.T_VALUE3 = "" Then
        MsgBox "请填写查询条件:" & TSQL.T_NAME3 & "的值", vbInformation, "提示"
        Exit Sub

    End If

    strSql = TSQL.T_BODY & " where 1= 1 and " & TSQL.T_COLUMN3 & " = " & " '" & TSQL.T_VALUE3 & "' "
Else
    Exit Sub

End If

STR_QUERY = strSql
Call ShowData(strSql)

End Sub

Private Sub ShowData(strSql As String)
Dim rs As New ADODB.Recordset
Dim i  As Integer

Set rs = Get_OracleRs(strSql)

With fps
    .MaxRows = 0
    Set .DataSource = rs

End With

Set rs = Nothing

For i = 1 To fps.MaxRows

    With fps
        .Row = i
        .Col = 1
        .Text = 1

    End With

Next i

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cboMaintainType_Click
' Description:       根据类型带出支持选项
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/12/17-10:08:47
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cboMaintainType_Click()
Dim strMaintainType As String
Dim rs              As New ADODB.Recordset
Dim strSql          As String

Call ClearData
strMaintainType = cboMaintainType.Text
strSql = "select t1.query_flag,t1.new_flag,t1.modify_flag,t1.delete_flag,t1.import_flag,t1.export_flag from TBL_MAINTAIN_SYS t1 where t1.maintain_type = '" & strMaintainType & "' "
Set rs = Get_OracleRs(strSql)
btnQuery.Enabled = IIf(rs!query_flag = "Y", True, False)
btnNew.Enabled = IIf(rs!new_flag = "Y", True, False)
btnModify.Enabled = IIf(rs!modify_flag = "Y", True, False)
btnDelete.Enabled = IIf(rs!delete_flag = "Y", True, False)
btnImport.Enabled = IIf(rs!import_flag = "Y", True, False)
btnExport.Enabled = IIf(rs!export_flag = "Y", True, False)
Call GetSqlQuery(strMaintainType, TSQL, TSQL_UPDATE)

End Sub

Private Sub Check1_Click()
Dim i As Integer

If Check1.Value = 1 Then

    For i = 1 To fps.MaxRows

        With fps
            .Row = i
            .Col = 1
            .Text = 1

        End With

    Next i

ElseIf Check1.Value = 0 Then

    For i = 1 To fps.MaxRows

        With fps
            .Row = i
            .Col = 1
            .Text = 0

        End With

    Next i

End If

End Sub

Private Sub Form_Load()
Call InitMaintainType
Call InitFps

End Sub

Private Sub InitFps()

With fps
    .TypeMaxEditLen = 500
    .MaxRows = 0
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsBest
    .Col = -1
    .Row = -1
    .Lock = True
    .Col = 1
    .CellType = CellTypeCheckBox
    .TypeHAlign = TypeHAlignCenter
    .TypeVAlign = TypeVAlignCenter
    .Lock = False
    .ColWidth(1) = 2

End With

End Sub

Private Function GetSqlQuery(strMaintainType As String, _
                             TSQL As T_SQL, _
                             TSQL_UPDATE As T_SQL_UPDATE) As String
Dim strCOLUMN1_NAME1 As String
Dim strCOLUMN2_NAME2 As String
Dim strCOLUMN3_NAME3 As String
Dim strMODIFY_NAME1  As String

TSQL.T_BODY = Get_OracleStr("select t1.query_sql from TBL_MAINTAIN_SYS t1 where t1.maintain_type = '" & cboMaintainType.Text & "'")
strCOLUMN1_NAME1 = Get_OracleStr("select t1.T_COLUMN1_NAME1 from TBL_MAINTAIN_SYS t1 where t1.maintain_type = '" & cboMaintainType.Text & "'")
If strCOLUMN1_NAME1 <> "" Then
    TSQL.T_COLUMN1 = Split(strCOLUMN1_NAME1, "|")(0)
    TSQL.T_NAME1 = Split(strCOLUMN1_NAME1, "|")(1)
    lblCol1Name.Visible = True
    txtCol1Value.Visible = True
    lblCol1Name.Caption = TSQL.T_NAME1
Else
    lblCol1Name.Visible = False
    txtCol1Value.Visible = False
    txtCol1Value.Text = ""

End If

strCOLUMN2_NAME2 = Get_OracleStr("select t1.T_COLUMN2_NAME2 from TBL_MAINTAIN_SYS t1 where t1.maintain_type = '" & cboMaintainType.Text & "'")
If strCOLUMN2_NAME2 <> "" Then
    TSQL.T_COLUMN2 = Split(strCOLUMN2_NAME2, "|")(0)
    TSQL.T_NAME2 = Split(strCOLUMN2_NAME2, "|")(1)
    lblCol2Name.Visible = True
    txtCol2Value.Visible = True
    lblCol2Name.Caption = TSQL.T_NAME2
Else
    lblCol2Name.Visible = False
    txtCol2Value.Visible = False
    txtCol2Value.Text = ""

End If

strCOLUMN3_NAME3 = Get_OracleStr("select t1.T_COLUMN3_NAME3 from TBL_MAINTAIN_SYS t1 where t1.maintain_type = '" & cboMaintainType.Text & "'")
If strCOLUMN3_NAME3 <> "" Then
    TSQL.T_COLUMN3 = Split(strCOLUMN3_NAME3, "|")(0)
    TSQL.T_NAME3 = Split(strCOLUMN3_NAME3, "|")(1)
    lblCol3Name.Visible = True
    txtCol3Value.Visible = True
    lblCol3Name.Caption = TSQL.T_NAME3
Else
    lblCol3Name.Visible = False
    txtCol3Value.Visible = False
    txtCol3Value.Text = ""

End If

TSQL_UPDATE.T_BODY = Get_OracleStr("select t1.modify_sql from TBL_MAINTAIN_SYS t1 where t1.maintain_type = '" & cboMaintainType.Text & "'")
TSQL_UPDATE.T_BODY2 = Get_OracleStr("select t1.modify_sql2 from TBL_MAINTAIN_SYS t1 where t1.maintain_type = '" & cboMaintainType.Text & "'")
TSQL_UPDATE.T_NAME1 = Get_OracleStr("select t1.T_MODIFY_NAME1 from TBL_MAINTAIN_SYS t1 where t1.maintain_type = '" & cboMaintainType.Text & "'")
If TSQL_UPDATE.T_NAME1 <> "" Then
    lblModifyName1.Visible = True
    txtModifyValue1.Visible = True
    lblModifyName1.Caption = TSQL_UPDATE.T_NAME1
Else
    lblModifyName1.Visible = False
    txtModifyValue1.Visible = False
    txtModifyValue1.Text = ""

End If

End Function

Private Sub ClearData()
fps.MaxRows = 0
TSQL.T_BODY = ""
TSQL.T_COLUMN1 = ""
TSQL.T_NAME1 = ""
TSQL.T_VALUE1 = ""
TSQL.T_COLUMN2 = ""
TSQL.T_NAME2 = ""
TSQL.T_VALUE2 = ""
TSQL.T_COLUMN3 = ""
TSQL.T_NAME3 = ""
TSQL.T_VALUE3 = ""
txtCol1Value.Text = ""
txtCol2Value.Text = ""
txtCol3Value.Text = ""
txtModifyValue1.Text = ""
TSQL_UPDATE.T_BODY = ""
TSQL_UPDATE.T_NAME1 = ""
TSQL_UPDATE.T_SET_DATA1 = ""
TSQL_UPDATE.T_SET_DATA2 = ""
TSQL_UPDATE.T_VALUE1 = ""
TSQL_UPDATE.T_COLUMN1 = ""

End Sub

Private Function SaveTblEventRec(enumEventType As E_TBL_EVENT, _
                                 strEventKey As String, _
                                 strEventDesc As String, _
                                 strEventReason As String, _
                                 strEventRemark As String) As String
Dim strEventID   As String
Dim strSql       As String
Dim strUserName  As String
Dim strEventType As String

Select Case enumEventType

    Case E_INSERT
        strEventType = "INSERT"

    Case E_DELETE
        strEventType = "DELETE"

    Case E_UPDATE
        strEventType = "UPDATE"

    Case E_QUERY
        strEventType = "QUERY"

End Select

strEventID = Right("00" & Year(Now), 2) & Right("00" & Month(Now), 2) & Right$("00" & Day(Now), 2)
strEventID = strEventID & Right$("00" & Get_OracleStr("select nvl(max(EVENT_ID),0) + 1 from TBL_EVENT_RECORD where  instr(EVENT_ID,'" & strEventID & "') > 0 "), 2)
strSql = "insert into TBL_EVENT_RECORD(EVENT_ID,EVENT_TYPE,EVENT_KEY,EVENT_DESC,EVENT_REASON,USER_ID,USER_NAME,DATETIME,REMARK) values('" & strEventID & "','" & strEventType & "','" & strEventKey & "','" & strEventDesc & "','" & strEventReason & "','" & gUserName & "','" & gUserRealName & "',sysdate,'" & strEventRemark & "') "
If AddSql(strSql) > 0 Then
    MsgBox "事件已记录", vbInformation, "提示"
    SaveTblEventRec = strEventID
Else
    MsgBox "事件未记录", vbCritical, "提示"
    SaveTblEventRec = ""

End If

End Function
