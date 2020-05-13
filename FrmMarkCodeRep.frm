VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmMarkCodeRep 
   BackColor       =   &H8000000B&
   Caption         =   "机种打标码规则维护系统"
   ClientHeight    =   12270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17415
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
   ScaleHeight     =   12270
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   13095
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   17415
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5280
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "待验收项(NPI)"
         Height          =   195
         Left            =   2160
         TabIndex        =   20
         Top             =   2640
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "待建立项(IT)"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   13920
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton btnUpload 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4680
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":0000
         Picture         =   "FrmMarkCodeRep.frx":3072
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "导入PDS原件或截图"
         Top             =   1800
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton btnExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "退出"
         Height          =   615
         Left            =   3600
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":5814
         Picture         =   "FrmMarkCodeRep.frx":8886
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "新增"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnExport 
         BackColor       =   &H00E0E0E0&
         Caption         =   "导出"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3000
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":B028
         Picture         =   "FrmMarkCodeRep.frx":E09A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "新增"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnImport 
         BackColor       =   &H00E0E0E0&
         Caption         =   "导入"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2400
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":101C4
         Picture         =   "FrmMarkCodeRep.frx":13236
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "新增"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnDelete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "删除"
         Height          =   615
         Left            =   1800
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":14F30
         Picture         =   "FrmMarkCodeRep.frx":17FA2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "新增"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnModify 
         BackColor       =   &H00E0E0E0&
         Caption         =   "修改"
         Height          =   615
         Left            =   1200
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":1A744
         Picture         =   "FrmMarkCodeRep.frx":1D7B6
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "新增"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnNew 
         BackColor       =   &H00E0E0E0&
         Caption         =   "新增"
         Height          =   615
         Left            =   600
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":20828
         Picture         =   "FrmMarkCodeRep.frx":2389A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "新增"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtPDS 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtHTPN 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton btnQuery 
         BackColor       =   &H00E0E0E0&
         Caption         =   "查询"
         Height          =   615
         Left            =   0
         MaskColor       =   &H008080FF&
         Picture         =   "FrmMarkCodeRep.frx":259C4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "查询：不输入厂内机种，则查询所有机种"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "动态码用****表示,换行符用\\表示,固定码不变"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00FFC0FF&
         Height          =   1005
         Left            =   6600
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         ToolTipText     =   "请输入打标码相关描述,如第一行的AB取值来源,DATECODE的定义等信息"
         Top             =   1080
         Width           =   9495
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   8535
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   16095
         _Version        =   524288
         _ExtentX        =   28390
         _ExtentY        =   15055
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
         MaxCols         =   10
         MaxRows         =   0
         SpreadDesigner  =   "FrmMarkCodeRep.frx":28A36
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "厂内机种"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   9
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "识别码"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   810
         TabIndex        =   8
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "描述"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6120
         TabIndex        =   7
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PDS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   6
         Top             =   1800
         Width           =   435
      End
   End
End
Attribute VB_Name = "FrmMarkCodeRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_MK

    E_ID = 1
    E_HTPN
    E_CREATE_DATE
    E_CREATE_BY
    E_ESTABLISH_FLAG
    E_BUY_OFF_FLAG
    E_REMARK
    E_DESCRIBE
    E_END

End Enum

Dim strTitle   As String
Dim strContent As String
Dim strPDSPath As String
'
'Private Sub btnExit_Click()
'Unload Me
'End Sub
'
'Private Sub btnExport_Click()
'Call ExporToExcel("select a.cust_code 客户机种,a.HT_PN  厂内机种,a.DEFINED_FLAG 是否打标,a.CARD_CONTROL_FLAG 是否申请, " & _
'"a.CARD_CONTROL_FLAG 是否建立,a.BUY_OFF_FLAG 是否验收,a.CARD_CONTROL_FLAG 是否卡控, " & _
'"a.DESCRIBE 打标规则描述,a.LINES_CNT 打标码行数,a.CHAR_CNT 打标码位数,a.REMARK 备注 from tbl_markingcode_rep a  " & _
'" order by a.cust_code,a.HT_PN ")
'End Sub

Private Sub btnDelete_Click()
Dim strHTPN As String
Dim strID   As String
Dim strSql  As String

'厂内机种
If Len(Trim(txtHTPN.Text)) = 0 Then
    MsgBox "请输入需要删除打标的厂内机种", vbInformation, "提示"
    Exit Sub

End If

strHTPN = Trim$(txtHTPN.Text)

If txtID.Text = "" Then
    MsgBox "请先搜索再双击需要删除机种的行", vbInformation, "提示"
    Exit Sub

End If

If MsgBox("请再次确认需要删除的打标规则" & vbCrLf & "删除请点击 是(Y),否则请点击 否(N) ", vbYesNo, "删除确认???") = vbNo Then
    Exit Sub

End If

strID = Trim$(txtID.Text)
strSql = "delete from tbl_markingcode_rep where HT_PN = '" & strHTPN & "' and id = " & strID & "  and  ESTABLISH_FLAG = 'N' and BUY_OFF_FLAG = 'N' "

If AddSql(strSql) > 0 Then
    strTitle = strHTPN & "打标码规则已删除"
    strContent = "NPI:" & gUserRealName & " " & strHTPN & "的打标码规则已删除"
    Call SentEml(strTitle, strContent)
    MsgBox "机种:" & strHTPN & "打标码规则已删除,请知悉", vbInformation, "提示"
Else
    MsgBox "删除失败", vbInformation, "提示"

End If

Call ClearTextBox

End Sub

Private Sub btnExit_Click()
Unload Me

End Sub

Private Sub btnExport_Click()
ExporToExcel ("select ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE from tbl_markingcode_rep  order by HT_PN")
End Sub

Private Sub btnModify_Click()
Dim strHTPN   As String
Dim strRemark As String
Dim strDesc   As String
Dim strPDS    As String
Dim strSql    As String
Dim lID       As Long

'厂内机种
If Len(Trim(txtHTPN.Text)) = 0 Then
    MsgBox "请输入需要建立打标的厂内机种", vbInformation, "提示"
    Exit Sub

End If

strHTPN = Trim$(txtHTPN.Text)

If Get_OracleCnt("select * from tbltsvnpiproduct where qtechptno = '" & strHTPN & "'") = 0 Then
    MsgBox "NPI对照表未维护该厂内机种: " & strHTPN & vbCrLf & "请输入正确的厂内机种", vbCritical, "警告"
    Exit Sub

End If

If Get_OracleCnt("select * from tbl_markingcode_rep where HT_PN = '" & strHTPN & "'") = 0 Then
    MsgBox "该厂内机种: " & strHTPN & "未申请过打标规则,无法修改规则" & vbCrLf & "如新增规则,请执行新增功能", vbCritical, "警告"
    Exit Sub

End If

'识别码
If Len(Trim(txtRemark.Text)) = 0 Then
    MsgBox "请输入打标规则的识别码", vbInformation, "提示"
    Exit Sub

End If

strRemark = Trim$(txtRemark.Text)

'PDS
If Len(Trim(txtPDS.Text)) = 0 Then
    MsgBox "请上传打标规则的PDS原件或截图", vbInformation, "提示"
    Exit Sub

End If

strPDS = Trim$(txtPDS.Text)

'描述
If Len(Trim(txtDesc.Text)) = 0 Then
    MsgBox "请输入打标规则的相关描述", vbInformation, "提示"
    Exit Sub

End If

strDesc = Trim$(txtDesc.Text)

If MsgBox("请再次确认修改后的打标相关信息是否准确" & vbCrLf & "修改请点击 是(Y),否则请点击 否(N) ", vbYesNo, "修改确认???") = vbNo Then
    Exit Sub

End If

strSql = "update tbl_markingcode_rep set UPDATE_DATE = sysdate,ESTABLISH_FLAG = 'N',BUY_OFF_FLAG= 'N',UPDATE_BY = '" & gUserName & "' || '" & gUserRealName & "' ,REMARK = '" & strRemark & "',DESCRIBE = '" & strDesc & "' where HT_PN = '" & strHTPN & "' and instr(CREATE_BY,'" & gUserName & "') > 0 "

If AddSql(strSql) > 0 Then
    strTitle = strHTPN & "打标码规则已修改"
    strContent = "NPI:" & gUserRealName & " " & strHTPN & "的打标码规则已修改,请IT尽快重新建立该规则"
    Call SentEml(strTitle, strContent)
    MsgBox "厂内机种:" & strHTPN & "的打标码规则已经修改,相关邮件已发送至IT窗口,请等待重新建立", vbInformation, "申请修改成功"
    strPDSPath = "\\10.160.1.84\public\FileServer\36.机种打标码PDS仓库\" & strHTPN

    If Dir(strPDSPath, vbDirectory) = "" Then
        MkDir strPDSPath

    End If

    Call CopyFileToFtp(txtPDS.Text, strPDSPath & "\")
Else
    MsgBox "修改失败", vbInformation, "提示"

End If

Call QueryData
Call ClearTextBox

End Sub

Private Sub btnNew_Click()
Dim strHTPN   As String
Dim strRemark As String
Dim strDesc   As String
Dim strPDS    As String
Dim strSql    As String
Dim lID       As Long

'厂内机种
If Len(Trim(txtHTPN.Text)) = 0 Then
    MsgBox "请输入需要建立打标的厂内机种", vbInformation, "提示"
    Exit Sub

End If

strHTPN = Trim$(txtHTPN.Text)

If Get_OracleCnt("select * from tbltsvnpiproduct where qtechptno = '" & strHTPN & "'") = 0 Then
    MsgBox "NPI对照表未维护该厂内机种: " & strHTPN & vbCrLf & "请输入正确的厂内机种", vbCritical, "警告"
    Exit Sub

End If

If Get_OracleCnt("select * from tbl_markingcode_rep where HT_PN = '" & strHTPN & "'") > 0 Then
    MsgBox "该厂内机种: " & strHTPN & "已申请过打标规则,无法新增规则" & vbCrLf & "如变更规则,请执行修改功能", vbCritical, "警告"
    Exit Sub

End If

'识别码
If Len(Trim(txtRemark.Text)) = 0 Then
    MsgBox "请输入打标规则的识别码", vbInformation, "提示"
    Exit Sub

End If

strRemark = Trim$(txtRemark.Text)

'PDS
If Len(Trim(txtPDS.Text)) = 0 Then
    MsgBox "请上传打标规则的PDS原件或截图", vbInformation, "提示"
    Exit Sub

End If

strPDS = Trim$(txtPDS.Text)

'描述
If Len(Trim(txtDesc.Text)) = 0 Then
    MsgBox "请输入打标规则的相关描述", vbInformation, "提示"
    Exit Sub

End If

strDesc = Trim$(txtDesc.Text)

If MsgBox("请再次确认打标相关信息是否准确" & vbCrLf & "新增请点击 是(Y),否则请点击 否(N) ", vbYesNo, "新增确认???") = vbNo Then
    Exit Sub

End If

lID = Get_OracleNo("select max(ID) + 1 from tbl_markingcode_rep")
strSql = "insert into tbl_markingcode_rep(ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE) values(" & lID & ",'" & strHTPN & "',sysdate,'" & gUserName & "' || '" & gUserRealName & "','N','N','" & strRemark & "','" & strDesc & "')  "

If AddSql(strSql) > 0 Then
    strTitle = strHTPN & "打标码规则已申请"
    strContent = "NPI:" & gUserRealName & " " & strHTPN & "的打标码规则已申请,请IT尽快建立该规则"
    Call SentEml(strTitle, strContent)
    MsgBox "厂内机种:" & strHTPN & "的打标码规则已经申请,相关邮件已发送至IT窗口,请等待建立", vbInformation, "申请建立成功"
    strPDSPath = "\\10.160.1.84\public\FileServer\36.机种打标码PDS仓库\" & strHTPN

    If Dir(strPDSPath, vbDirectory) = "" Then
        MkDir strPDSPath

    End If

    Call CopyFileToFtp(txtPDS.Text, strPDSPath & "\")

End If

Call QueryData
Call ClearTextBox

End Sub

Private Sub btnQuery_Click()
Call QueryData

End Sub

Private Sub QueryData()
Dim strHTPN As String
Dim strSql  As String
Dim rs      As New ADODB.Recordset
strHTPN = Trim$("" & txtHTPN.Text)
fps.MaxRows = 0

If strHTPN <> "" Then
    strSql = "select ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE from tbl_markingcode_rep  where HT_PN = '" & strHTPN & "' "
Else
    strSql = "select ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE from tbl_markingcode_rep  order by HT_PN"

End If

Set rs = Get_OracleRs(strSql)

If Not rs.EOF Then

    With fps
        Set .DataSource = rs

    End With

Else
    MsgBox "查询不到信息", vbInformation
    Exit Sub

End If

End Sub

Private Sub btnUpload_Click()
CommonDialog1.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
CommonDialog1.ShowOpen

If CommonDialog1.filename = "" Then
    Exit Sub

End If

txtPDS.Text = CommonDialog1.filename
CommonDialog1.filename = ""

If txtPDS.Text = "" Then
    MsgBox "请选择要上传的文件", vbInformation, "提示"
    Exit Sub

End If

End Sub

Private Sub Form_Load()
Call InitCtrls

End Sub

''--------------------------------------------------------------------------------
'' Project    :       正式工程1
'' Procedure  :       InitCtrls
'' Description:       初始化控件
'' Created by :       Project Administrator
'' Machine    :       0-354AD8C194ED4
'' Date-Time  :       2019-11-29-11:40:10
''
'' Parameters :
''--------------------------------------------------------------------------------
Private Sub InitCtrls()
'cbIsDefined.AddItem ("Y")
'cbIsDefined.AddItem ("N")
'cbIsApply.AddItem ("Y")
'cbIsApply.AddItem ("N")
'cbIsEstablish.AddItem ("Y")
'cbIsEstablish.AddItem ("N")
'cbIsBuyOff.AddItem ("Y")
'cbIsBuyOff.AddItem ("N")
'cbIsChecked.AddItem ("Y")
'cbIsChecked.AddItem ("N")
Call InitFps

End Sub

'
''--------------------------------------------------------------------------------
'' Project    :       正式工程1
'' Procedure  :       InitFps
'' Description:       初始化fps
'' Created by :       Project Administrator
'' Machine    :       0-354AD8C194ED4
'' Date-Time  :       2019-11-29-11:55:48
''
'' Parameters :
''--------------------------------------------------------------------------------
Private Sub InitFps()

With fps
    .MaxCols = E_MK.E_END - 1
    .Col = -1
    .Row = -1
    .TypeMaxEditLen = 500
    .Lock = True
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .SetText E_MK.E_ID, 0, "序号"
    .ColWidth(E_MK.E_ID) = 4
    .SetText E_MK.E_HTPN, 0, "厂内机种"
    .ColWidth(E_MK.E_HTPN) = 8
    .SetText E_MK.E_CREATE_DATE, 0, "申请日期"
    .ColWidth(E_MK.E_CREATE_DATE) = 8
    .SetText E_MK.E_CREATE_BY, 0, "申请人"
    .ColWidth(E_MK.E_CREATE_BY) = 8
    .SetText E_MK.E_ESTABLISH_FLAG, 0, "是否建立"
    .ColWidth(E_MK.E_ESTABLISH_FLAG) = 8
    .SetText E_MK.E_BUY_OFF_FLAG, 0, "是否验收"
    .ColWidth(E_MK.E_BUY_OFF_FLAG) = 8
    .SetText E_MK.E_REMARK, 0, "识别码"
    .ColWidth(E_MK.E_REMARK) = 15
    .SetText E_MK.E_DESCRIBE, 0, "描述"
    .ColWidth(E_MK.E_DESCRIBE) = 60

End With

End Sub

Private Sub fps_Click(ByVal Col As Long, ByVal Row As Long)
Dim strID  As String
Dim strSql As String
Dim rs     As New ADODB.Recordset

If Row = 0 Then Exit Sub

With fps
    .Row = Row
    .Col = 1
    strID = .Text

End With

txtID.Text = strID
strSql = "select HT_PN,REMARK,DESCRIBE from tbl_markingcode_rep where ID = " & strID & " "
Set rs = Get_OracleRs(strSql)
txtHTPN.Text = Trim$("" & rs!HT_PN)
txtRemark.Text = Trim$("" & rs!REMARK)
txtDesc.Text = Trim$("" & rs!DESCRIBE)

End Sub

Private Sub Fps_RightClick(ByVal ClickType As Integer, _
                           ByVal Col As Long, _
                           ByVal Row As Long, _
                           ByVal MouseX As Long, _
                           ByVal MouseY As Long)
Dim strHTPN     As String
Dim strCreateBy As String

With fps
    .Row = Row
    .Col = 2
    strHTPN = Trim$("" & .Text)
    .Col = 4
    strCreateBy = Trim$("" & .Text)

End With

If gUserName = "07885" And Option1.Value = True Then
    If MsgBox("是否已经建立好" & strHTPN & "的打标码规则,选择是(Y)将发送通知邮件到对应NPI窗口,否(N)退出", vbYesNo, "已建立???") = vbNo Then
        Exit Sub

    End If

    If AddSql("update tbl_markingcode_rep set ESTABLISH_FLAG = 'Y' where HT_PN = '" & strHTPN & "' and ESTABLISH_FLAG = 'N' ") > 0 Then
        strTitle = strHTPN & "打标码规则已建立"
        strContent = "IT:" & gUserRealName & " " & strHTPN & "的打标码规则已经验收,请NPI尽快测试验收"
        Call SentEml(strTitle, strContent)
        MsgBox "邮件已通知对应NPI窗口,厂内机种" & strHTPN & "打标码规则已建立", vbInformation, "提示"

    End If

    Call Option1_Click
ElseIf gUserName <> "07885" And Option2.Value = True Then

    If InStr(strCreateBy, gUserName) = 0 Then
        MsgBox "你的账号只能BuyOff自己负责的打标码", vbCritical, "警告"
        Exit Sub

    End If

    If MsgBox("是否已经验收" & strHTPN & "的打标码规则,选择是(Y)将发送通知邮件到对应IT窗口,否(N)退出", vbYesNo, "已验收???") = vbNo Then
        Exit Sub

    End If

    If AddSql("update tbl_markingcode_rep set BUY_OFF_FLAG = 'Y' where HT_PN = '" & strHTPN & "' and ESTABLISH_FLAG = 'Y' and  BUY_OFF_FLAG = 'N'") > 0 Then
        strTitle = strHTPN & "打标码规则已验收"
        strContent = "NPI:" & gUserRealName & " " & strHTPN & "的打标码规则已经验收,请IT知悉"
        Call SentEml(strTitle, strContent)
        MsgBox "邮件已通知对应IT窗口,厂内机种" & strHTPN & "打标码规则已验收", vbInformation, "提示"

    End If

    Call Option2_Click

End If

End Sub

Private Sub Option1_Click()
Dim strSql As String
Dim rs     As New ADODB.Recordset
fps.MaxRows = 0

If Option1.Value = True Then
    strSql = "select ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE from tbl_markingcode_rep  where ESTABLISH_FLAG = 'N' and BUY_OFF_FLAG = 'N' "
    Set rs = Get_OracleRs(strSql)

    If Not rs.EOF Then

        With fps
            Set .DataSource = rs

        End With

    Else
        MsgBox "查询不到待建立项", vbInformation, "提示"
        Exit Sub

    End If

End If

End Sub

Private Sub Option2_Click()
Dim strSql As String
Dim rs     As New ADODB.Recordset
fps.MaxRows = 0

If Option2.Value = True Then
    strSql = "select ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE from tbl_markingcode_rep  where ESTABLISH_FLAG = 'Y' and BUY_OFF_FLAG = 'N'  "
    Set rs = Get_OracleRs(strSql)

    If Not rs.EOF Then

        With fps
            Set .DataSource = rs

        End With

    Else
        MsgBox "查询不到待验收项", vbInformation, "提示"
        Exit Sub

    End If

End If

End Sub

Private Sub ClearTextBox()
Dim obj As Object

For Each obj In Me.Controls

    If TYPENAME(obj) = "TextBox" Then
        obj.Text = ""

    End If

Next

End Sub

Private Function SentEml(strSentTitle As String, strSentText As String) As Boolean
Dim i             As Integer
Dim dirtemp       As String
Dim strSentTo(50) As String
Dim strSentCC(10) As String
i = 0
dirtemp = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\邮件接收\SentTo_MarkingCode.cfg"
Open dirtemp For Input As #1

Do While Not EOF(1)
    Line Input #1, strTemp
    strSentTo(i) = Trim$(strTemp)
    i = i + 1
Loop
Close #1

If SentMes(strSentTitle, strSentText, strSentTo, txtPDS.Text, strSentCC) = True Then
    SentEml = True
Else
    SentEml = False

End If

End Function

Private Function SentMes(Subject As String, _
                         SentText As String, _
                         Recipient() As String, _
                         Attachment As String, _
                         RecipientCC() As String) As Boolean
Dim JM As Object
Set JM = CreateObject("JMAIL.Message")
Dim Recipients()   As String
Dim RecipientCCs() As String
Dim strBodyinfo    As String
Dim i              As Integer
Dim strSql         As String
Dim j              As Integer
Dim rs             As New ADODB.Recordset
Dim RsD            As New ADODB.Recordset

On Error GoTo ErrHandler

SentMes = False
JM.Charset = "gb2312"
JM.Silent = False
JM.Priority = 1
JM.MailServerUserName = "sqladmin" '帐号
JM.MailServerPassWord = "ksitadmin" '密码
JM.From = "sqladmin@ht-tech.com"    '名字
JM.FromName = "sqladmin"  '发件人姓名

'收件人
For i = 0 To UBound(Recipient) - 1

    If Recipient(i) <> "" Then
        JM.AddRecipient Recipient(i)

    End If

Next

'抄送人
For i = 0 To UBound(RecipientCC) - 1

    If RecipientCC(i) <> "" Then
        JM.AddRecipientCC RecipientCC(i)

    End If

Next

'附件
If Attachment <> "" Then
    If Dir(Attachment, vbNormal Or vbArchive) = "" Then
        Exit Function
    Else
        JM.AddAttachment Attachment

    End If

End If

JM.Subject = Subject
JM.AppendText SentText
SentMes = JM.Send("mail.ht-tech.com")
ErrHandler:
Set JM = Nothing
Exit Function

End Function
