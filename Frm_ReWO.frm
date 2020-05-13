VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_ReWO 
   Caption         =   "WO维护"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8055
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
   ScaleHeight     =   6105
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtWaferCnt 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6240
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ReWO.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ReWO.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ReWO.frx":18A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   870
      Left            =   0
      TabIndex        =   3
      Top             =   5235
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1535
      ButtonWidth     =   1561
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   查找    "
            Key             =   "SEARCH"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "生成"
            Key             =   "ADD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "EXIT"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00FF0000&
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.TextBox txtWaferNo 
         Height          =   285
         Left            =   4200
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtLotID 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   3855
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   7575
         _Version        =   524288
         _ExtentX        =   13361
         _ExtentY        =   6800
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
         SpreadDesigner  =   "Frm_ReWO.frx":1BF6
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wafer片数"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WaferNo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   570
      End
   End
End
Attribute VB_Name = "Frm_ReWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0

    E_LOT = 1
    e_ID
    E_Wafer
    E_TotalDie
    E_CHECK
    E_END

End Enum

Private Sub Form_Load()
Init

End Sub

Private Sub Init()
InitWidget
InitFps

End Sub

Private Sub InitWidget()

Select Case Frm_ProductionPlan.cbWOType.text

    Case "重工工单"
        Label1(0).Visible = True
        Label1(1).Visible = True
        txtLotID.Visible = True
        txtWaferNo.Visible = True

    Case "Dummy工单"
        Label1(2).Visible = True
        txtWaferCnt.Visible = True

    Case "玻璃工单"
        Label1(2).Visible = True
        txtWaferCnt.Visible = True

    Case "FO_CSP工单"
        Label1(2).Visible = True
        txtWaferCnt.Visible = True

    Case "硅基工单"
        Label1(2).Visible = True
        txtWaferCnt.Visible = True

End Select

End Sub

Private Sub InitFps()

With Fps(0)
    .ReDraw = False
    .MaxCols = E_FPS0.E_END - 1
    .MaxRows = 0
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    .Col = E_FPS0.E_CHECK
    .CellType = CellTypeCheckBox
    .TypeHAlign = TypeHAlignCenter
    .TypeVAlign = TypeVAlignCenter
    .SetText E_FPS0.E_LOT, 0, "LotID"
    .SetText E_FPS0.e_ID, 0, "NO"
    .SetText E_FPS0.E_Wafer, 0, "WaferID"
    .SetText E_FPS0.E_TotalDie, 0, "TotalDies"
    .SetText E_FPS0.E_CHECK, 0, "选择"
    .ColWidth(E_FPS0.E_LOT) = 15
    .ColWidth(E_FPS0.e_ID) = 5
    .ColWidth(E_FPS0.E_Wafer) = 20
    .ColWidth(E_FPS0.E_TotalDie) = 10
    .ColWidth(E_FPS0.E_CHECK) = 4
    .RowHeight(0) = 20
    .RowHeight(-1) = 15
    .Col = E_FPS0.E_CHECK
    .Lock = False
    .ReDraw = True

End With

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "SEARCH"
        ForSearch

    Case "EXIT"
        ForExit

    Case "ADD"
        ForAdd

End Select

End Sub

Private Sub ForExit()
Unload Me

End Sub

Private Sub ForSearch()
If Frm_ProductionPlan.cbWOType.text = "玻璃工单" Then
    If InStr(Frm_ProductionPlan.txtCusPN.text, "-CV") = 0 Then
        Frm_ProductionPlan.txtCusPN.text = Frm_ProductionPlan.txtCusPN.text & "-CV"

    End If

End If

If Frm_ProductionPlan.cbWOType.text = "重工工单" Then
    SearchRWO
Else
    SearchOther

End If

End Sub

Private Sub SearchRWO()
Dim strLotID       As String
Dim strWaferNo     As String
Dim strSubstrateid As String
Dim Rs2            As New ADODB.Recordset
Dim strSql         As String
Dim iLastRows      As Integer
Dim rs             As New ADODB.Recordset

iLastRows = 0
If txtLotID.text = "" Then
    MsgBox "请输入LOTID", vbCritical, "提醒"
    Exit Sub

End If

strLotID = Trim(txtLotID.text)
If txtWaferNo.text = "" Then
    MsgBox "请输入WaferNO", vbCritical, "提醒"
    Exit Sub

End If

strWaferNo = CInt(Trim$(txtWaferNo.text))

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 2
        If strWaferNo = Trim$(.text) Then
            MsgBox "请不要重复搜索:" & strWaferNo & "片", vbCritical, "警告"
            Exit Sub

        End If

    Next

End With

strSql = "select distinct b.substrateid from customeroitbl_test a, mappingdatatest b where b.lotid = '" & strLotID & "' and to_number(b.wafer_id) = '" & strWaferNo & "' and to_char(a.id) = b.filename and a.source_batch_id = b.lotid " & " and a.invflag = 0 and instr(b.substrateid, '+') > 0 and not exists (select 1 from ib_waferlist c where b.substrateid = c.waferid) "
If Rs2.State = adStateOpen Then Rs2.Close
Rs2.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If Rs2.RecordCount > 0 Then
    MsgBox "第" & strWaferNo & "片已经生成了未开立工单的WaferID: " & Rs2("substrateid") & ", 请不要再维护该片", vbCritical, "提醒"
    Exit Sub

End If

strora = "select lotid,substrateid || '+' as substrateid, '','' from mappingdatatest where to_number(wafer_id) = '" & strWaferNo & "' and lotid = '" & strLotID & "' order by to_number(filename) desc"
If rs.State = adStateOpen Then rs.Close
rs.Open strora, Cnn, adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount = 0 Then
    MsgBox "查不到该LOT", vbCritical, "警告"
    Exit Sub

End If

With Fps(0)
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
    .Col = 1
    .text = rs("lotid")
    .Col = 2
    .text = strWaferNo
    .Col = 3
    .text = rs("substrateid")
    .Col = 4
    .text = ""
    .Lock = False
    .CellType = CellTypeEdit
    .BackColorStyle = BackColorStyleUnderGrid
    .BackColor = vbCyan
    .Col = 5
    .text = CStr("1")

End With

txtWaferNo.text = ""
Toolbar1.Buttons("ADD").Enabled = True

End Sub

Private Sub SearchOther()
Dim sWaferQty As String
Dim iWaferQty As Long
Dim iLotQty   As Long
Dim lDies     As Long
Dim strSql    As String

If Frm_ProductionPlan.cbCusCode.text = "" Then
    MsgBox "请输入客户代码", vbCritical, "警告"
    Exit Sub

End If

If Frm_ProductionPlan.txtCusPN.text = "" Then
    MsgBox "请输入客户机种", vbCritical, "警告"
    Exit Sub

End If

If Frm_ProductionPlan.cbPN.text <> "" Then
    strSql = "select distinct customerdieqty from tbltsvnpiproduct where  customershortname = '" & Trim(Frm_ProductionPlan.cbCusCode.text) & "' and  customerptno1 = '" & Trim(Frm_ProductionPlan.txtCusPN.text) & "' and qtechptno2 = '" & Trim(Frm_ProductionPlan.cbPN.text) & "'  "
Else
    strSql = "select distinct customerdieqty from tbltsvnpiproduct where  customershortname = '" & Trim(Frm_ProductionPlan.cbCusCode.text) & "' and  customerptno1 = '" & Trim(Frm_ProductionPlan.txtCusPN.text) & "' "

End If

lDies = Get_OracleNo(strSql)
If lDies = 0 Then
    MsgBox "NPI没有维护该客户代码对应的客户机种:" & Frm_ProductionPlan.txtCusPN.text
    Exit Sub

End If

If Frm_ProductionPlan.cbWOType.text = "硅基工单" Then
    If InStr(Frm_ProductionPlan.txtCusPN.text, "-FO") = 0 Then
        MsgBox "硅基工单的客户机种和厂内机种的格式必须为: '" & Frm_ProductionPlan.txtCusPN.text & "-FO', 请联系NPI维护对应机种和硅基料号", vbInformation, "提醒"
        Exit Sub

    End If

End If

'判断系统是否有未开立工单的waferid, 否则提示不让新建
If Frm_ProductionPlan.cbWOType.text = "玻璃工单" Then
    strSql = "select distinct a.source_batch_id from customeroitbl_test a, mappingdatatest b where a.customershortname = '" & Trim(Frm_ProductionPlan.cbCusCode.text) & "' and a.mpn_desc = '" & Trim(Frm_ProductionPlan.txtCusPN.text) & "' and a.flag = 'T'  " & "and to_char(a.id) = b.filename and a.source_batch_id = b.lotid and a.source_batch_id like 'G%' and instr(b.substrateid ,'+') = 0 and not exists (select 1 from ib_waferlist c where b.substrateid = c.waferid)   " & " order by a.source_batch_id "
Else
    strSql = "select distinct a.source_batch_id from customeroitbl_test a, mappingdatatest b where a.customershortname = '" & Trim(Frm_ProductionPlan.cbCusCode.text) & "' and a.mpn_desc = '" & Trim(Frm_ProductionPlan.txtCusPN.text) & "' and a.flag = 'T'  " & "and to_char(a.id) = b.filename and a.source_batch_id = b.lotid and instr(b.substrateid ,'+') = 0 and not exists (select 1 from ib_waferlist c where b.substrateid = c.waferid)   " & " order by a.source_batch_id "

End If

If Get_OracleCnt(strSql) > 0 Then
    MsgBox "之前已经维护过该客户机种的waferlot,尚未开工单; 请不要再维护多余的", vbInformation, "提示"
    Exit Sub

End If

sWaferQty = Trim(txtWaferCnt.text)
If sWaferQty = "" Then
    MsgBox "请输入wafer片数", vbInformation, "友情提示"
    Exit Sub
Else
    If IsNumeric(sWaferQty) = False Then
        MsgBox "请输入数字", vbInformation, "友情提示"
        Exit Sub
    Else
        If CLng(sWaferQty) < 1 Then
            MsgBox "请输入至少1片wafer数量", vbInformation, "友情提示"
            Exit Sub

        End If

    End If

End If

If Frm_ProductionPlan.cbWOType.text = "玻璃工单" Then
    iWaferQty = CLng(sWaferQty)
    iLotQty = IIf((iWaferQty Mod 12) = 0, iWaferQty \ 12, iWaferQty \ 12 + 1)
    If iLotQty > 1 Then

        For i = 1 To (iLotQty - 1)
            Call ShowOtherData(12, lDies)
        Next
        iWaferQty = IIf((iWaferQty Mod 12) = 0, 12, iWaferQty Mod 12)
        Call ShowOtherData(iWaferQty, lDies)
    Else
        iWaferQty = IIf((iWaferQty Mod 12) = 0, 12, iWaferQty Mod 12)
        Call ShowOtherData(iWaferQty, lDies)

    End If

Else
    iWaferQty = CLng(sWaferQty)
    iLotQty = IIf((iWaferQty Mod 25) = 0, iWaferQty \ 25, iWaferQty \ 25 + 1)
    If iLotQty > 1 Then

        For i = 1 To (iLotQty - 1)
            Call ShowOtherData(25, lDies)
        Next
        iWaferQty = IIf((iWaferQty Mod 25) = 0, 25, iWaferQty Mod 25)
        Call ShowOtherData(iWaferQty, lDies)
    Else
        iWaferQty = IIf((iWaferQty Mod 25) = 0, 25, iWaferQty Mod 25)
        Call ShowOtherData(iWaferQty, lDies)

    End If

End If

Toolbar1.Buttons("SEARCH").Enabled = False
Toolbar1.Buttons("ADD").Enabled = True

End Sub

Private Sub ShowOtherData(iWaferQty As Long, lDies As Long)
Dim sOra     As String
Dim sLotTmp  As String
Dim sLotType As String
Dim strSql   As String
Dim htdevice As String

strSql = "select distinct replace( qtechptno,'-CV','') from tbltsvnpiproduct where  customershortname = '" & Trim(Frm_ProductionPlan.cbCusCode.text) & "' and  customerptno1 = '" & Trim(Frm_ProductionPlan.txtCusPN.text) & "' "
htdevice = Get_OracleStr(strSql)

Select Case Frm_ProductionPlan.cbWOType.text

    Case "Dummy工单"
        sLotType = "D"

    Case "玻璃工单"
        sLotType = "G"

    Case "硅基工单"
        sLotType = "SI"

    Case "FO_CSP工单"
        sLotType = "SI"

    Case Else
        MsgBox "未知工单"
        Exit Sub

End Select

If sLotType <> "G" Then
    sOra = "select SPECIALLOT.nextval ID from dual"
    sLotTmp = sLotType & Right(Year(Now), 2) & Right(("0" & Month(Now)), 2) & Right(("0" & Day(Now)), 2) & Right("000" & Hex(Get_OracleNo(sOra)), 4)
Else
    sOra = "select seq_Glass_id.Nextval ID   from dual"
    sLotTmp = sLotType & htdevice & Right("000" & Hex(Get_OracleNo(sOra)), 4)

End If

With Fps(0)

    For iWafer = 1 To iWaferQty
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1
        .text = sLotTmp
        .Col = 2
        .text = iWafer
        .Col = 3
        .text = sLotTmp & Right$("0" & iWafer, 2)
        .Col = 4
        If Frm_ProductionPlan.cbWOType.text = "硅基工单" Then
            .text = "1"
        Else
            .text = CStr(lDies)
            .text = CStr(lDies)
            .Lock = False
            .CellType = CellTypeEdit
            .BackColorStyle = BackColorStyleUnderGrid
            .BackColor = vbCyan

        End If

        .Col = 5
        .text = CStr("1")
    Next

End With

End Sub

Private Sub ForAdd()
If Frm_ProductionPlan.cbWOType.text = "重工工单" Then
    AddReWO
Else
    AddOther

End If

End Sub

Private Sub AddReWO()
Dim strWaferID    As String
Dim strWaferIDNew As String
Dim lGoodDies     As Long
Dim lNgDies       As Long
Dim strWaferNo    As String
Dim strLotID      As String

lNgDies = 0

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_FPS0.E_CHECK
        If .text = "1" Then
            .Col = E_FPS0.E_TotalDie
            If .text = "" Then
                MsgBox "请输入总DIE数", vbCritical, "提醒"
                Exit Sub

            End If

        End If

    Next

End With

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_FPS0.E_CHECK
        If .text = "1" Then
            .Col = E_FPS0.E_TotalDie
            lGoodDies = Val(Trim$(.text))
            .Col = E_FPS0.E_Wafer
            strWaferIDNew = Trim$(.text)
            strWaferID = Left(Trim$(.text), Len(Trim(.text)) - 1)
            .Col = E_FPS0.e_ID
            strWaferNo = Trim$(.text)
            .Col = E_FPS0.E_LOT
            strLotID = Trim$(.text)
            Call InsertToDB(strWaferID, strWaferIDNew, strWaferNo, lGoodDies, lNgDies, strLotID)

        End If

    Next

End With

End Sub

Private Sub AddOther()
Dim strLotID      As String
Dim lGoodDies     As Long
Dim lNgDies       As Long
Dim strWaferIDNew As String
Dim strWaferNo    As String

lNgDies = 0
If Frm_ProductionPlan.cbCusCode.text = "" Then
    MsgBox "请输入客户代码", vbCritical, "警告"
    Exit Sub

End If

If Frm_ProductionPlan.txtCusPN.text = "" Then
    MsgBox "请输入客户机种", vbCritical, "警告"
    Exit Sub

End If

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_FPS0.E_CHECK
        If .text = "1" Then
            .Col = E_FPS0.E_LOT
            strLotID = Trim$(.text)
            .Col = E_FPS0.E_TotalDie
            lGoodDies = Val(Trim$(.text))
            .Col = E_FPS0.E_Wafer
            strWaferIDNew = Trim$(.text)
            .Col = E_FPS0.e_ID
            strWaferNo = Trim$(.text)
            .Col = E_FPS0.E_LOT
            strLotID = Trim$(.text)
            Call InsertToDB(strLotID, strWaferIDNew, strWaferNo, lGoodDies, lNgDies, strLotID)

        End If

    Next

End With

MsgBox "WO插入成功, 请开立工单", vbInformation, "提示"
Unload Me

End Sub

Private Sub InsertToDB(strWaferID As String, _
                       strWaferIDNew As String, _
                       strWaferNo As String, _
                       lGoodDIe As Long, _
                       lNGDie As Long, _
                       strLotID As String)
Dim cmdStr     As String
Dim cmdStr2    As String
Dim sSeqID     As Long
Dim strCusPN   As String
Dim strCusCode As String
Dim strMark    As String
Dim rs         As New ADODB.Recordset
Dim lID        As Long

On Error GoTo ERRORON

Cnn.BeginTrans
INIadoCon.BeginTrans
If Frm_ProductionPlan.cbWOType.text = "重工工单" Then
    ' 获取ID
    lID = Get_OracleNo("select filename from mappingdatatest where to_number(wafer_id) = '" & strWaferNo & "' and lotid = '" & strLotID & "' order by to_number(filename) desc ")
    ' 检查
    cmdStr = "select * from mappingDataTest where substrateid = '" & strWaferIDNew & "'"
    If rs.State = adStateOpen Then rs.Close
    rs.Open cmdStr, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
        MsgBox strWaferIDNew & "已经存在,请不要添加同样的WAFERID", vbCritical, "警告"
        Exit Sub

    End If

    ' 插入
    sSeqID = GetMaxID()
    cmdStr = " insert into CustomerOItbl_test(id, po_num,po_item,source_batch_id,source_mtrl_num,mtrl_num,mtrl_desc,test_mtrl_num,test_mtrl_desc,mpn,mpn_desc,source_mtrl_sloc, " & _
       " mtrl_num_mtrlgrp,probe_ship_part_type,offshore_asm_company,offshore_test_company,current_wafer_qty,die_qty,design_id,country_of_fab,fab_conv_id,fab_excr_id,reticle_level_71, " & _
       " reticle_level_72,reticle_level_73,wafer_size,imager_customer_rev,chromaticity,micro_lens_shift,temperature_spec,prb_containment_type,fabrication_facility,prb_excr_id,batch_comment_probe, " & _
       " assy_process_id,dark_bond_pad_assy,assy_serial_type,sticky_backs_to_save,optical_quality,encoded_mark_id,planned_laser_scribe,package_lid_type,package_type,pb_free_package,target_waf_thickness, " & _
       " reliability_sampling,lot_priority,wafer_box_type,test_site,assembly_facility,batch_comment_assy,tst_process_id,elec_special_test,box_type,protective_film_apld,shipping_mst_260,shipping_mst_level, " & _
       " t_price,ship_comment,batch_comment_test,created_date,created_time,unit_price,ref_po,ref_po_item,country_of_assembly,micron_material,date_code,ship_site,special_process_lot,lot_status,custom_part_no, " & _
       " flag,qtech_created_by,qtech_created_date,qtech_lastupdate_by,qtech_lastupdate_date,customershortname,downqty,invflag,wafer_visual_inspect,comp_code,eqdatacode,jobno,zx_fromsite,zx_invoice)   " & _
       " select  distinct '" & sSeqID & "',ct.po_num,ct.po_item,ct.source_batch_id,ct.source_mtrl_num,ct.mtrl_num,ct.mtrl_desc,ct.test_mtrl_num,ct.test_mtrl_desc,ct.mpn,ct.mpn_desc,ct.source_mtrl_sloc,ct.mtrl_num_mtrlgrp, " & _
       " ct.probe_ship_part_type,ct.offshore_asm_company,ct.offshore_test_company,ct.current_wafer_qty,ct.die_qty,ct.design_id,ct.country_of_fab,ct.fab_conv_id,ct.fab_excr_id,ct.reticle_level_71,ct.reticle_level_72, " & _
       " ct.reticle_level_73,ct.wafer_size,ct.imager_customer_rev,ct.chromaticity,ct.micro_lens_shift,ct.temperature_spec,ct.prb_containment_type,ct.fabrication_facility,ct.prb_excr_id,ct.batch_comment_probe, " & _
       " ct.assy_process_id,ct.dark_bond_pad_assy,ct.assy_serial_type,ct.sticky_backs_to_save,ct.optical_quality,ct.encoded_mark_id,ct.planned_laser_scribe,ct.package_lid_type,ct.package_type,ct.pb_free_package, " & _
       " ct.target_waf_thickness,ct.reliability_sampling,ct.lot_priority,ct.wafer_box_type,ct.test_site,ct.assembly_facility,ct.batch_comment_assy,ct.tst_process_id,ct.elec_special_test,ct.box_type, " & _
       " ct.protective_film_apld,ct.shipping_mst_260,ct.shipping_mst_level,ct.t_price,ct.ship_comment,ct.batch_comment_test,ct.created_date,ct.created_time,ct.unit_price,ct.ref_po,ct.ref_po_item, " & _
       " ct.country_of_assembly,ct.micron_material,ct.date_code,ct.ship_site,ct.special_process_lot,ct.lot_status, " & _
       " ct.custom_part_no,'T','" & gUserName & "',sysdate,ct.qtech_lastupdate_by,ct.qtech_lastupdate_date,ct.customershortname,ct.downqty,ct.invflag,ct.wafer_visual_inspect, " & _
       " ct.comp_code,ct.eqdatacode,ct.jobno,ct.zx_fromsite,ct.zx_invoice from CustomerOItbl_test ct, MAPPINGDATATEST mt  where ct.id = " & lID & " and to_char(ct.id) = mt.filename "
    cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerOI](id, po_num,po_item,source_batch_id,source_mtrl_num,mtrl_num,mtrl_desc,test_mtrl_num,test_mtrl_desc,mpn,mpn_desc,source_mtrl_sloc, " & _
       " mtrl_num_mtrlgrp,probe_ship_part_type,offshore_asm_company,offshore_test_company,current_wafer_qty,die_qty,design_id,country_of_fab,fab_conv_id,fab_excr_id,reticle_level_71, " & _
       " reticle_level_72,reticle_level_73,wafer_size,imager_customer_rev,chromaticity,micro_lens_shift,temperature_spec,prb_containment_type,fabrication_facility,prb_excr_id,batch_comment_probe, " & _
       " assy_process_id,dark_bond_pad_assy,assy_serial_type,sticky_backs_to_save,optical_quality,encoded_mark_id,planned_laser_scribe,package_lid_type,package_type,pb_free_package,target_waf_thickness, " & _
       " reliability_sampling,lot_priority,wafer_box_type,test_site,assembly_facility,batch_comment_assy,tst_process_id,elec_special_test,box_type,protective_film_apld,shipping_mst_260,shipping_mst_level, " & _
       " t_price,ship_comment,batch_comment_test,created_date,created_time,unit_price,ref_po,ref_po_item,country_of_assembly,micron_material,date_code,ship_site,special_process_lot,lot_status,custom_part_no, " & _
       " flag,qtech_created_by,qtech_created_date,qtech_lastupdate_by,qtech_lastupdate_date,customershortname,downqty,wafer_visual_inspect,comp_code,eqdatacode,jobno,zx_fromsite,zx_invoice)   " & _
       " select distinct  '" & sSeqID & "',ct.po_num,ct.po_item,ct.source_batch_id,ct.source_mtrl_num,ct.mtrl_num,ct.mtrl_desc,ct.test_mtrl_num,ct.test_mtrl_desc,ct.mpn,ct.mpn_desc,ct.source_mtrl_sloc,ct.mtrl_num_mtrlgrp, " & _
       " ct.probe_ship_part_type,ct.offshore_asm_company,ct.offshore_test_company,ct.current_wafer_qty,ct.die_qty,ct.design_id,ct.country_of_fab,ct.fab_conv_id,ct.fab_excr_id,ct.reticle_level_71,ct.reticle_level_72, " & _
       " ct.reticle_level_73,ct.wafer_size,ct.imager_customer_rev,ct.chromaticity,ct.micro_lens_shift,ct.temperature_spec,ct.prb_containment_type,ct.fabrication_facility,ct.prb_excr_id,ct.batch_comment_probe, " & _
       " ct.assy_process_id,ct.dark_bond_pad_assy,ct.assy_serial_type,ct.sticky_backs_to_save,ct.optical_quality,ct.encoded_mark_id,ct.planned_laser_scribe,ct.package_lid_type,ct.package_type,ct.pb_free_package, " & _
       " ct.target_waf_thickness,ct.reliability_sampling,ct.lot_priority,ct.wafer_box_type,ct.test_site,ct.assembly_facility,ct.batch_comment_assy,ct.tst_process_id,ct.elec_special_test,ct.box_type, " & _
       " ct.protective_film_apld,ct.shipping_mst_260,ct.shipping_mst_level,ct.t_price,ct.ship_comment,ct.batch_comment_test,ct.created_date,ct.created_time,ct.unit_price,ct.ref_po,ct.ref_po_item, " & _
       " ct.country_of_assembly,ct.micron_material,ct.date_code,ct.ship_site,ct.special_process_lot,ct.lot_status, " & _
       " ct.custom_part_no,'T','" & gUserName & "',GetDate(),ct.qtech_lastupdate_by,ct.qtech_lastupdate_date,ct.customershortname,ct.downqty,ct.wafer_visual_inspect, " & _
       " ct.comp_code,ct.eqdatacode,ct.jobno,ct.zx_fromsite,ct.zx_invoice from [ERPBASE].[dbo].[tblCustomerOI] ct, [ERPBASE].[dbo].[tblmappingData] mt  where ct.id = " & lID & " and convert(varchar,ct.id) = mt.filename"
    AddSql (cmdStr)
    AddSql2 (cmdStr2)
    cmdStr = "insert into mappingDataTest (id, substrateid, productid, lotid, Wafer_ID, passbincount, failbincount, CustomerShortName, flag, Qtech_Created_By, Qtech_Created_Date,filename) " & " select  mappingData_SEQ.Nextval, '" & strWaferIDNew & "', productid, lotid,Wafer_ID,  '" & lGoodDIe & "', '" & lNGDie & "', CustomerShortName, 'T',  '" & gUserName & "', sysdate, '" & sSeqID & "' " & " from MAPPINGDATATEST  where filename =  '" & lID & "' and to_number(wafer_id) = '" & strWaferNo & "' "
    cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)" & " select  '" & strWaferIDNew & "', productid, lotid,Wafer_ID, '" & lGoodDIe & "', '" & lNGDie & "', CustomerShortName, 'T',  '" & gUserName & "', GETDATE(), '" & sSeqID & "' " & " from [ERPBASE].[dbo].[tblmappingData]  where filename = '" & lID & "' and wafer_id in ('" & strWaferNo & "', '0' + '" & strWaferNo & "') "
    AddSql (cmdStr)
    AddSql2 (cmdStr2)
    Cnn.CommitTrans
    INIadoCon.CommitTrans
    MsgBox "重工WAFER:" & strWaferIDNew & "成功生成", vbInformation, "提示:"
    Exit Sub
Else
    strCusCode = Trim(Frm_ProductionPlan.cbCusCode.text)
    strCusPN = Trim(Frm_ProductionPlan.txtCusPN.text)
    strLotID = strWaferID
    strMark = Right$(strWaferIDNew, 6)
    sSeqID = GetMaxID()
    cmdStr = "insert into CustomerOItbl_test(id,source_batch_id,SHIP_SITE,mpn_desc,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date) values('" & sSeqID & "', '" & strLotID & "', '" & strCusCode & "', '" & strCusPN & "', '" & strCusCode & "','T', '" & gUserName & "', sysdate)"
    cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,source_batch_id,SHIP_SITE,mpn_desc,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date) values('" & sSeqID & "', '" & strLotID & "', '" & strCusCode & "', '" & strCusPN & "', '" & strCusCode & "','T', '" & gUserName & "', GETDATE())"
    AddSql (cmdStr)
    AddSql2 (cmdStr2)
    cmdStr = "insert into mappingDataTest (substrateid, productid, lotid, Wafer_ID, passbincount, failbincount, CustomerShortName, flag, Qtech_Created_By, Qtech_Created_Date,filename) values('" & strWaferIDNew & "', '" & strMark & "', '" & strLotID & "', '" & strWaferNo & "','" & lGoodDIe & "', '" & lNGDie & "', '" & strCusCode & "','T', '" & gUserName & "', sysdate, '" & sSeqID & "') "
    cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)  values('" & strWaferIDNew & "', '" & strMark & "', '" & strLotID & "', '" & strWaferNo & "','" & lGoodDIe & "', '" & lNGDie & "', '" & strCusCode & "','T', '" & gUserName & "', GETDATE(), '" & sSeqID & "') "
    AddSql (cmdStr)
    AddSql2 (cmdStr2)
    If Frm_ProductionPlan.cbWOType = "硅基工单" Then
        sSeqID = GetMaxID()
        strCusPN = Replace$(strCusPN, "-FO", "")
        lGoodDIe = Get_OracleStr("select customerdieqty from tbltsvnpiproduct where customerptno1 = '" & strCusPN & "'")
        cmdStr = "insert into CustomerOItbl_test(id,source_batch_id,SHIP_SITE,mpn_desc,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date) values('" & sSeqID & "', '" & strLotID & "', '" & strCusCode & "', '" & strCusPN & "', '" & strCusCode & "','T', '" & gUserName & "', sysdate)"
        cmdStr2 = "insert into [ERPBASE].[dbo].[tblCustomerOI](id,source_batch_id,SHIP_SITE,mpn_desc,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date) values('" & sSeqID & "', '" & strLotID & "', '" & strCusCode & "', '" & strCusPN & "', '" & strCusCode & "','T', '" & gUserName & "', GETDATE())"
        AddSql (cmdStr)
        AddSql2 (cmdStr2)
        cmdStr = "insert into mappingDataTest (substrateid, productid, lotid, Wafer_ID, passbincount, failbincount, CustomerShortName, flag, Qtech_Created_By, Qtech_Created_Date,filename) values('" & strWaferIDNew & "' || '+', '" & strMark & "', '" & strLotID & "', '" & strWaferNo & "','" & lGoodDIe & "', '" & lNGDie & "', '" & strCusCode & "','T', '" & gUserName & "', sysdate, '" & sSeqID & "') "
        cmdStr2 = "insert into [ERPBASE].[dbo].[tblmappingData] (substrateid,productid,lotid,Wafer_ID,passbincount,failbincount,CustomerShortName,flag,Qtech_Created_By,Qtech_Created_Date,filename)  values('" & strWaferIDNew & "' + '+', '" & strMark & "', '" & strLotID & "', '" & strWaferNo & "','" & lGoodDIe & "', '" & lNGDie & "', '" & strCusCode & "','T', '" & gUserName & "', GETDATE(), '" & sSeqID & "') "
        AddSql (cmdStr)
        AddSql2 (cmdStr2)

    End If

    Cnn.CommitTrans
    INIadoCon.CommitTrans
    Exit Sub

End If

ERRORON:
Cnn.RollbackTrans
INIadoCon.RollbackTrans
MsgBox "订单生成失败:" & Err.DESCRIPTION, vbInformation, "提示:"

End Sub
