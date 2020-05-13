VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_UploadAw 
   Caption         =   "上传艾为客户PO"
   ClientHeight    =   8805
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16770
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
   ScaleHeight     =   19351.65
   ScaleMode       =   0  'User
   ScaleWidth      =   27791.19
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "上传记录查询"
      Height          =   6855
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Width           =   19455
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   10200
         TabIndex        =   14
         Top             =   240
         Width           =   3135
         Begin VB.OptionButton OptS 
            Caption         =   "二次上传"
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton OptF 
            Caption         =   "一次上传"
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查询"
         Height          =   960
         Left            =   7200
         TabIndex        =   11
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtText3 
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtText2 
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   720
         Width           =   2535
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   4815
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   1800
         Width           =   18855
         _Version        =   524288
         _ExtentX        =   33258
         _ExtentY        =   8493
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
         SpreadDesigner  =   "Frm_UploadAw.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lblLotWafer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LotWafer"
         Height          =   195
         Left            =   3360
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblLotID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LotID"
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   480
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PO上传"
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   19455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "上传"
         Height          =   960
         Left            =   5760
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton cmd 
         Caption         =   ".."
         Height          =   360
         Left            =   4560
         TabIndex        =   4
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtText1 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   2895
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   10560
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblXlsx 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择待上传的xlsx："
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   2
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label lblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码 :"
         Height          =   195
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   825
      End
   End
End
Attribute VB_Name = "Frm_UploadAw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Enum E_F_PO

    e_PO = 1
    E_Dev
    E_WaferType
    E_LOTID
    E_END

End Enum

Enum E_F_PO_S

    e_PO = 1
    E_Dev
    e_PKG
    E_partno
    E_WaferType
    E_Tracing_Code
    E_Assembly_LotID
    E_LOTID
    E_END

End Enum

Private Sub cmd_Click()

'GC
On Error Resume Next

Dim FName

'帅选文件
CommonDialog1.Filter = "EXCEL文件(*.xlsx)|*.xlsx|EXCEL文件(*.xls)|*.xls"
CommonDialog1.ShowOpen
'得到文件名
FName = CommonDialog1.filename
If FName <> "" Then
    txtText1.Text = FName

End If

End Sub

Private Sub cmdQuery_Click()
Dim LOTID As String

LOTID = UCase$(Trim$(txtText2))
InitFpsHeader
' 判断Lot号是否提供
If LOTID = "" Then
    MsgBox "请输入Lot号"
    Exit Sub

End If

If OptF.Value Then
    Call ShowDataF(LOTID)
Else
    Call ShowDataS(LOTID)

End If

End Sub

Private Function ShowDataF(LOTID As String)
Dim cmd_ora    As String
Dim mainItemRS As New ADODB.Recordset

cmd_ora = "select ct.po_num, ct.REF_PO, ct.COUNTRY_OF_ASSEMBLY, ct.source_batch_id from customeroitbl_test ct where ct.source_batch_id = '" & LOTID & "'" & "and ct.po_num is not null and rownum = 1"
Set mainItemRS = getStr(cmd_ora)

With fpS(0)
    .MaxRows = 0
    If mainItemRS.RecordCount > 0 Then
        Set .DataSource = mainItemRS

    End If

End With

End Function

Private Function ShowDataS(LOTID As String)
Dim cmd_ora    As String
Dim mainItemRS As New ADODB.Recordset

cmd_ora = "select ct.po_num,ct.REF_PO,ct.comp_code,ct.eqdatacode,ct.country_of_assembly,ct.zx_fromsite,ct.zx_invoice, ct.source_batch_id " & "from customeroitbl_test ct where ct.source_batch_id = '" & LOTID & "' and ct.po_num is not null and ct.comp_code is not null "
Set mainItemRS = getStr(cmd_ora)

With fpS(0)
    .MaxRows = 0
    If mainItemRS.RecordCount > 0 Then
        Set .DataSource = mainItemRS

    End If

End With

End Function

Private Sub cmdUpload_Click()
Dim customerStr As String

If Trim(Combo1.Text) = "" Then
    MsgBox "请先选择客户！"
    Exit Sub

End If

customerStr = UCase(Trim(Combo1.Text))
Call UploadAWPO(customerStr)

End Sub

Private Sub Form_Load()
Combo1.AddItem ("HK037")
Combo1.AddItem ("AC70")
OptF.Value = True
Call InitFpsHeader

End Sub

Private Sub InitFpsHeader()
If OptF.Value Then

    With fpS(0)
        .ReDraw = False
        .MaxCols = E_F_PO.E_END - 1
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
        ' 定义可选框
        ' .Col = sel_box
        ' .CellType = CellTypeCheckBox
        ' .TypeHAlign = TypeHAlignCenter
        ' .TypeVAlign = TypeVAlignCenter
        ' 定义表头名
        .SetText E_F_PO.e_PO, 0, "PO_NO"
        .SetText E_F_PO.E_Dev, 0, "Device产品编码"
        .SetText E_F_PO.E_WaferType, 0, "WaferType"
        .SetText E_F_PO.E_LOTID, 0, "Wafer Lot ID"
        ' 定义宽度
        .ColWidth(1) = 20
        .ColWidth(2) = 30
        .ColWidth(3) = 8
        .ColWidth(4) = 8
        ' 定义高度
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        ' 定义是否可编辑
        '        .Col = 22
        '           .Lock = False
        .ReDraw = True

    End With

Else

    With fpS(0)
        .ReDraw = False
        .MaxCols = E_F_PO_S.E_END - 1
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
        ' 定义可选框
        ' .Col = sel_box
        ' .CellType = CellTypeCheckBox
        ' .TypeHAlign = TypeHAlignCenter
        ' .TypeVAlign = TypeVAlignCenter
        ' 定义表头名
        .SetText E_F_PO_S.e_PO, 0, "PO_NO"
        .SetText E_F_PO_S.E_Dev, 0, "Device产品编码"
        .SetText E_F_PO_S.e_PKG, 0, "Package"
        .SetText E_F_PO_S.E_partno, 0, "Part NO"
        .SetText E_F_PO_S.E_WaferType, 0, "Wafer Type"
        .SetText E_F_PO_S.E_Tracing_Code, 0, "Tracing Code"
        .SetText E_F_PO_S.E_Assembly_LotID, 0, "Assembly Lot ID"
        .SetText E_F_PO_S.E_LOTID, 0, "Wafer Lot ID"
        ' 定义宽度
        .ColWidth(1) = 20
        .ColWidth(2) = 30
        .ColWidth(3) = 18
        .ColWidth(4) = 18
        .ColWidth(5) = 18
        .ColWidth(6) = 18
        .ColWidth(7) = 18
        .ColWidth(8) = 10
        ' 定义高度
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        ' 定义是否可编辑
        '        .Col = 22
        '           .Lock = False
        .ReDraw = True

    End With

End If

End Sub

Private Sub UploadAWPO(customerNameTemp As String)
If txtText1.Text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub

End If

Set VBExcel = CreateObject("excel.application")     '创建Excle对象
VBExcel.Visible = False
Set xlBook = VBExcel.Workbooks.Open(txtText1.Text)    '打开文件
Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
'2)判定最大列Excel中的和设定列是否相同
If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 11 Then
    MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
    Exit Sub

End If

' 定义变量
Dim po_no           As String
Dim dev_code        As String
Dim PACKAGE         As String
Dim part_no         As String
Dim wafer_type      As String
Dim Lot_id          As String
Dim Wafer_id        As String
Dim tracing_code    As String
Dim assembly_lot_id As String
Dim die_qty         As Long
Dim lot_wafer_id    As String
Dim a               As Integer
Dim b               As Integer

' 遍历表格
' 第2行开始,循环更换行号
For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count

    ' 查询一行的值
    ' 第1列开始,循环增加列数
    For J = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + J)
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
        If J = 1 Then
            po_no = Trim(tempVal)

        End If

        If J = 2 Then
            dev_code = Trim(tempVal)

        End If

        If J = 3 Then
            PACKAGE = Trim(tempVal)

        End If

        If J = 4 Then
            part_no = Trim(tempVal)

        End If

        If J = 5 Then
            wafer_type = Trim(tempVal)

        End If

        If J = 6 Then
            Lot_id = Trim(tempVal)

        End If

        If J = 7 Then
            Wafer_id = Trim(tempVal)
            Wafer_id = IIf(Len(Wafer_id) = 1, "0" & Wafer_id, Wafer_id)

        End If

        If J = 8 Then

        End If

        If J = 9 Then
            Dim die_tmp As Long

            If (PACKAGE = "") Then
            Else
                die_qty = Trim(tempVal)

            End If

        End If

        If J = 10 Then
            tracing_code = Trim(tempVal)

        End If

        If J = 11 Then
            assembly_lot_id = Trim(tempVal)

        End If

    Next J

    lot_wafer_id = Trim(Lot_id & Wafer_id)
    ' 一次上传处理
    If (lot_wafer_id <> "" And PACKAGE = "") Then
        If GetHeaderId(lot_wafer_id, po_no, dev_code, wafer_type) Then
            a = a + 1

        End If

    End If

    ' 二次上传处理
    If (lot_wafer_id <> "" And PACKAGE <> "") Then
        If GetHeaderIdPlus(lot_wafer_id, po_no, PACKAGE, part_no, dev_code, wafer_type, die_qty, tracing_code, assembly_lot_id) Then
            b = b + 1

        End If

    End If

Next i

xlBook.Close      '总是提示是否保存   结束Excel
Set xlSheet = Nothing
Set xlBook = Nothing
Set VBExcel = Nothing
'VBExcel.Quit
MsgBox "已成功上传" & a & "笔一次订单PO, 已成功上传" & b & "笔二次订单PO ！", vbInformation, "友情提示"

End Sub

Private Function GetHeaderId(lot_wafer As String, _
                             po_no As String, _
                             dev_code As String, _
                             wafer_type As String) As Boolean
Dim file_name As String
Dim cmd_ora   As String
Dim cmd_sql   As String
Dim strCustPN As String
Dim strPackage As String

GetHeaderId = False

cmd_ora = "select a.filename from mappingdatatest a WHERE a.substrateid IN   RTRIM('" + lot_wafer + "') "
file_name = getStr2(cmd_ora)

strCustPN = Get_OracleStr("select mpn_desc from customeroitbl_test where id = '" & file_name & "'")
strPackage = Get_SqlStr("SELECT PACKAGE FROM erptemp..device_attribute_ac70 where CUST_DEVICE = '" & strCustPN & "'")
If strPackage = "" Then
    MsgBox "对照表没有该机种的PACKAGE", vbInformation, "提示"
    Exit Function
End If


If file_name <> "" Then
    ' 判断是否有一次上传PO
    cmd_ora = "select ct.po_num from customeroitbl_test ct where ct.id = '" & file_name & "' "
    If (QueryStr2(cmd_ora)) Then
        MsgBox ("已存在一次上传PO号,  不需要重复上传一次PO, 请确认是否有误")
        Exit Function

    End If

    'REF_PO - Device 产品编码      COUNTRY_OF_ASSEMBLY - Wafer Type   define by tw 20171023
    cmd_ora = "Update customeroitbl_test ct set ct.PO_NUM='" & po_no & "', ct.REF_PO='" & dev_code & "',ct.reticle_level_72 = '" & strPackage & "', ct.COUNTRY_OF_ASSEMBLY='" & wafer_type & "'  where ct.id='" & file_name & "' and ct.flag='Y'"
    AddSql (cmd_ora)
    cmd_sql = "Update [ERPBASE].[dbo].[tblCustomerOI]  set PO_NUM='" & po_no & "', REF_PO='" & dev_code & "',ct.reticle_level_72 = '" & strPackage & "', COUNTRY_OF_ASSEMBLY='" & wafer_type & "'  where id='" & file_name & "' and flag='Y'"
    AddSql2 (cmd_sql)
    GetHeaderId = True

End If

End Function

Private Function GetHeaderIdPlus(lot_wafer_id As String, _
                                 po_no As String, _
                                 package_flag As String, _
                                 part_no As String, _
                                 dev_code As String, _
                                 wafer_type As String, _
                                 die_qty As Long, _
                                 tracing_code As String, _
                                 assembly_lot_id As String) As Boolean
Dim file_name      As String
Dim cmd_ora        As String
Dim cmd_sql        As String
Dim ID             As Long
Dim customerTemp   As String
Dim strMarkingCode As String
Dim strPackage As String

by_user = gUserName
GetHeaderIdPlus = False
' 判断是否有二次上传PO
cmd_ora = "select a.substrateid  from mappingdatatest a where a.substrateid in ('" & lot_wafer_id & "+')"
If (QueryStr(cmd_ora)) Then
    MsgBox ("已存在二次上传PO号, 不需要重复上传二次PO, 请确认是否有误")

    'Exit Function
End If

If part_no = "AW33805CSR" Then
    strMarkingCode = "3805\\" & tracing_code

End If

strPackage = Get_SqlStr("SELECT PACKAGE FROM erptemp..device_attribute_ac70 where CUST_DEVICE = '" & part_no & "'")
If strPackage = "" Then
    MsgBox "对照表没有该机种的PACKAGE", vbInformation, "提示"
    Exit Function
End If

' 添加二次子表
ID = GetMaxID()
cmd_ora = "insert into mappingdatatest(id,filename, substrateid,lotid,passbincount,failbincount,flag,Qtech_Created_By,Qtech_Created_Date,wafer_id,customershortname,productid) select mappingData_SEQ.Nextval,'" & ID & "', " & "substrateid || '+',lotid, '" & die_qty & "',failbincount,'Y','Auto',sysdate,wafer_id,customershortname,'" & strMarkingCode & "' from mappingdatatest where substrateid = '" & lot_wafer_id & "'  "
AddSql (cmd_ora)
' sql  server
cmd_sql = "insert into [ERPBASE].[dbo].[tblmappingData](filename, substrateid,lotid,passbincount,failbincount,flag,Qtech_Created_By,Qtech_Created_Date,wafer_id,customershortname,productid) select '" & ID & "', " & "substrateid + '+',lotid, '" & die_qty & "',failbincount,'Y','Auto',GETDATE(),wafer_id,customershortname ,'" & strMarkingCode & "' from [ERPBASE].[dbo].[tblmappingData] where substrateid = '" & lot_wafer_id & "'  "
AddSql2 (cmd_sql)
' 添加二次头表
cmd_ora = "select a.filename from mappingdatatest a WHERE a.substrateid IN   RTRIM('" + lot_wafer_id + "') "
file_name = getStr2(cmd_ora)
cmd_ora = " insert into customeroitbl_test(id, po_num, source_batch_id, mtrl_num, mpn_desc, test_site,created_date, ref_po, country_of_assembly, ship_site , flag, qtech_created_by,qtech_created_date, customershortname, downqty, invflag, jobno," & " comp_code, eqdatacode, zx_fromsite, zx_invoice) " & "select '" & ID & "', '" & po_no & "', source_batch_id, mtrl_num, '" & part_no & "','HTKS', created_date,'" & dev_code & "',  '" & wafer_type & "' , ship_site, flag,'" & by_user & "',sysdate ,customershortname, downqty, invflag, jobno, '" & strPackage & "', '" & part_no & "', '" & tracing_code & "', '" & assembly_lot_id & "' from customeroitbl_test where id =  '" & file_name & "' "
AddSql (cmd_ora)
cmd_sql = " insert into [ERPBASE].[dbo].[tblCustomerOI](id, po_num, source_batch_id, mtrl_num, mpn_desc, test_site,created_date, ref_po, country_of_assembly, ship_site , flag, qtech_created_by,qtech_created_date, customershortname, downqty, jobno," & " comp_code, eqdatacode, zx_fromsite, zx_invoice) " & "select '" & ID & "', '" & po_no & "', source_batch_id, mtrl_num, '" & part_no & "','HTKS', created_date,'" & dev_code & "',  '" & wafer_type & "' , ship_site, flag,'" & by_user & "',  GETDATE() ,customershortname, downqty, jobno, '" & strPackage & "', '" & part_no & "', '" & tracing_code & "', '" & assembly_lot_id & "' from [ERPBASE].[dbo].[tblCustomerOI] where id =  '" & file_name & "' "
' sql server
AddSql2 (cmd_sql)
GetHeaderIdPlus = True

End Function
