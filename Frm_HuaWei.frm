VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_HuaWei 
   Caption         =   "国税局发票上传"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18525
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
   ScaleHeight     =   10530
   ScaleWidth      =   18525
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "上传记录"
      Height          =   1095
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   18495
      Begin VB.OptionButton Option2 
         Height          =   195
         Left            =   6000
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLablePrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LablePrint:"
         Height          =   195
         Left            =   5040
         TabIndex        =   12
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblCarton 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carton:"
         Height          =   195
         Left            =   3120
         TabIndex        =   11
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.Frame Fra 
      Height          =   1935
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   18495
      Begin VB.CommandButton cmdExcelIn 
         Caption         =   "华为标签上传"
         Height          =   480
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "华为标签保存"
         Height          =   480
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   1095
         Left            =   3240
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退        出"
         Height          =   1080
         Left            =   9240
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径："
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Fra 
      Height          =   6015
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   3000
      Width           =   18495
      Begin FPSpreadADO.fpSpread fps 
         Height          =   4935
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   17535
         _Version        =   524288
         _ExtentX        =   30930
         _ExtentY        =   8705
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
         MaxCols         =   5
         MaxRows         =   5
         SpreadDesigner  =   "Frm_HuaWei.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   10560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_HuaWei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum FpsDetail
    e_Barcode = 1
    E_PO = 2
    e_PCS = 3
    e_ItemCode = 4
    e_MPN = 5
    e_ItemDesc = 6
    e_09code = 7
    e_09barcode = 8
    e_Vendor_code = 9
    e_ROHS = 10
    e_SuppCode
    e_VendorLot
    e_Country
    E_DATE
    e_Remarks
    e_UOM
    e_PoLine
    e_ShipNo
    e_ItemDescEn
    e_Law
    
    e_MCol = 28
End Enum

Private Sub cmdExcelIn_Click()
On Error GoTo ErrHandler

Dim FName
    '筛选文件
    Com.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
    Com.ShowOpen
    '得到文件名
    FName = Com.filename
    If FName <> "" Then
       txtPath.Text = FName  '路径显示出来
       '将资料写到FPS
       FileExportInFps
    End If
    
Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub
Private Sub InitCtrl()
Dim i               As Integer

    'Fps初始化
    With fps(0)
        .ReDraw = False
        .DAutoSizeCols = DAutoSizeColsBest
        .MaxRows = 0
        .MaxCols = 28
        .ColsFrozen = 1
        .Row = -1
        .Col = -1
        .Lock = True
        
        .UserColAction = UserColActionSort
        
         For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        
        .ZOrder
        .ReDraw = True
    End With

End Sub

Private Sub cmdExit_Click()
    Unload Me

End Sub

Private Sub cmdHistory_Click()

Dim sql As String
Dim mainItemRS As New adodb.Recordset
'Dim mainItemRs As adodb.Recordset

sql = "select * from HUAWEI_CARTON "


Set mainItemRS = getStr(sql)

With fps(0)
   .MaxRows = 0
        
    If mainItemRS.RecordCount > 0 Then
        Set .DataSource = mainItemRS
       
    End If
End With



End Sub

Private Sub cmdSave_Click() '资料保存
On Error GoTo ErrHandle
Dim strSql                          As String
Dim Rs                              As New adodb.Recordset
Dim i                               As Integer
Dim strTmp(FpsDetail.e_MCol)        As String
Dim strOra                          As String
Dim filename As String
    
    '检查资料
    If fps(0).MaxRows <= 0 Then
        MsgBox "没有要保存的资料", vbInformation, "提示"
        Exit Sub
    End If
    
    If MsgBox("是否要保存吗？", vbInformation + vbYesNo, "提示") = vbNo Then Exit Sub
    '如果有资料，开始插入数据库
    '开启事物模式
    MousePointer = 11
    With fps(0)
        .Row = 0
        .Col = 1
        strTmp(0) = .Text
        
        
    
        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 1
            strTmp(1) = Trim$(.Text)
            
            .Col = 2
            strTmp(2) = Trim$(.Text)
            
            .Col = 3
            strTmp(3) = Trim$(.Text)
            
            .Col = 4
            strTmp(4) = Trim$(.Text)
            
            .Col = 5
            strTmp(5) = Trim$(.Text)
            
            .Col = 6
            strTmp(6) = Trim$(.Text)
            
            .Col = 7
            strTmp(7) = Trim$(.Text)
            
            .Col = 8
            strTmp(8) = Trim$(.Text)
            
            .Col = 9
            strTmp(9) = Trim$(.Text)
            
            .Col = 10
            strTmp(10) = Trim$(.Text)
            
            .Col = 11
            strTmp(11) = Trim$(.Text)
            
            .Col = 12
            strTmp(12) = Trim$(.Text)
            
            .Col = 13
            strTmp(13) = Trim$(.Text)
            
            .Col = 14
            strTmp(14) = Trim$(.Text)
            
            .Col = 15
            strTmp(15) = Trim$(.Text)
            
            .Col = 16
            strTmp(16) = Trim$(.Text)
            
            .Col = 17
            strTmp(17) = Trim$(.Text)
            
            .Col = 18
            strTmp(18) = Trim$(.Text)
            
            .Col = 19
            strTmp(19) = Trim$(.Text)
            
            .Col = 20
            strTmp(20) = Trim$(.Text)
            
            .Col = 21
            strTmp(21) = Trim$(.Text)
            
            .Col = 22
            strTmp(22) = Trim$(.Text)
            
            .Col = 23
            strTmp(23) = Trim$(.Text)
            
            .Col = 24
            strTmp(24) = Trim$(.Text)
            
            .Col = 25
            strTmp(25) = Trim$(.Text)
            
            .Col = 26
            strTmp(27) = Trim$(.Text)
            
            .Col = 27
            strTmp(27) = Trim$(.Text)
            
            .Col = 28
            strTmp(28) = Trim$(.Text)
        
        ' 1)校验
        ' 2)插入数据库
        
            If (Left(strTmp(0), 2) = "Ba") Then
            
                strOra = "insert into HUAWEI_CARTON(BAR_CODE, PO_NUMBER, PCS, ITEM_CODE,ITEM_REV,MPN, ITEM_DESC, CODE_09, BARCODE_09, VENDOR_CODE,COMPANY_CODE, INSPECT_FLAG,RESTRICT_FLAG, ROHS, SUPP_CODE, VENDOR_LOT, COUNTRY, PRODUC_DATE, REMARKS,UOM, PO_LINE_NUM, SHIPMENT_NUM, ITEMDESC_EN, LAW_INS_FLAG, HW_M, SN_TN) " & _
                    "values('" & strTmp(1) & "', '" & strTmp(2) & "','" & strTmp(3) & "','" & strTmp(4) & "','" & strTmp(5) & "','" & strTmp(6) & "','" & strTmp(7) & "','" & strTmp(8) & "','" & strTmp(9) & "','" & strTmp(10) & "','" & strTmp(11) & "','" & strTmp(12) & "','" & strTmp(13) & "','" & strTmp(14) & "','" & strTmp(15) & "'," & _
                    "'" & strTmp(16) & "','" & strTmp(17) & "','" & strTmp(18) & "','" & strTmp(19) & "','" & strTmp(20) & "','" & strTmp(21) & "','" & strTmp(22) & "','" & strTmp(23) & "','" & strTmp(24) & "','" & strTmp(25) & "','" & strTmp(26) & "') "
           
            Else
           
                strOra = "insert into HUAWEI_LABLE(PART_NO, VER, CE, FCC, ROHS, CI, P, QTY, UNIT, SN_TN, HW_M, ITEM_DESC_CN, ITEM_DESC_EN, SN, PKG_ID, MPN, MFG_CODE, MAN_DATE, M_LOT, LAW, G_W, CODE_09, PO, REMARK, COUNTRY, TOTAL_QTY, PKG_CODE,ORDER_INFO)" & _
                "values('" & strTmp(1) & "', '" & strTmp(2) & "','" & strTmp(3) & "','" & strTmp(4) & "','" & strTmp(5) & "','" & strTmp(6) & "','" & strTmp(7) & "','" & strTmp(8) & "','" & strTmp(9) & "','" & strTmp(10) & "','" & strTmp(11) & "','" & strTmp(12) & "','" & strTmp(13) & "','" & strTmp(14) & "','" & strTmp(15) & "'," & _
                    "'" & strTmp(16) & "','" & strTmp(17) & "','" & strTmp(18) & "','" & strTmp(19) & "','" & strTmp(20) & "','" & strTmp(21) & "','" & strTmp(22) & "','" & strTmp(23) & "','" & strTmp(24) & "','" & strTmp(25) & "','" & strTmp(26) & "','" & strTmp(27) & "','" & strTmp(28) & "') "
           
           
            End If
            
            AddSql (strOra)
            
        Next
    End With
    
    MousePointer = 0

    filename = txtPath.Text


    MsgBox " " & filename & " 资料保存成功！"
    
Exit Sub
    
ErrHandle:
    MousePointer = 0
    MsgBox Err.Description, vbCritical + vbInformation, "警告"
End Sub

Private Sub Form_Load()
    '初始化
    InitCtrl
End Sub
'Form大小自动调整
Private Sub Form_Resize()
    Fra(0).Move Fra(0).Left, Fra(0).Top, Me.ScaleWidth - 200, Fra(0).Height
    Fra(1).Move Fra(0).Left, Fra(1).Top, Me.ScaleWidth - 200, Me.ScaleHeight - Fra(0).Height - 50
    fps(0).Move fps(0).Left, fps(0).Top, Fra(1).Width - 300, Fra(1).Height - 300
End Sub
'导入资料
Private Sub FileExportInFps()
On Error GoTo ErrHandle
Dim VBExcel                         As Excel.Application
Dim xlBook                          As Excel.Workbook
Dim xlSheet                         As Excel.Worksheet
Dim strFileName                     As String
Dim i                               As Integer
Dim j                               As Integer
Dim strChar                         As String
Dim strTmp(FpsDetail.e_MCol)        As Variant
    
    MousePointer = 11
    'Fps
    fps(0).MaxRows = 0
    '获取文件名
    If InStrRev(Trim(txtPath.Text), "\") > 0 Then
        strFileName = Mid(Trim(txtPath.Text), InStrRev(Trim(txtPath.Text), "\") + 1)
        If InStr(strFileName, ".") > 0 Then
            strFileName = Mid(strFileName, 1, InStr(strFileName, ".") - 1)
        End If
    End If
    'Excel文件处理
    '1)打开Excel
    Set VBExcel = CreateObject("excel.application")      '创建Excle对象
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.open(txtPath.Text)    '打开文件
    Set xlSheet = xlBook.Worksheets(1)            '打开sheet中的表
    '判定最大列Excel中的和设定列是否相同
    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 26 And xlSheet.Range("A1").CurrentRegion.Columns.Count <> 28 Then
        MousePointer = 0
        MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        GoTo ExitPro
        Exit Sub
    End If
    '处理ExcelExcel
    With fps(0)
        For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.Count         '2)得到Excel最大行
            strTmp(0) = Trim(xlSheet.Range("A" & i).Value)
            If Len(strTmp(0)) > 0 Then
                If i <> 1 Then .MaxRows = .MaxRows + 1  '第一行表示标题，不用增加行
                For j = 1 To 28
                    '循环i,j 26(得到A.B.C.)
                    If j > 26 Then
                        strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
                    Else
                        strChar = Chr(96 + j)
                    End If
'                    strTmp(j) = xlSheet.Range(strChar & i).Value   '先屏蔽，换方法写
                    If i = 1 Then '得到第一行
                        .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))  '赋值到FPS
'                        For j = 0 To UBound(strTmp) - 1
'                            .SetText j + 1, i - 1, Trim$(strTmp(j + 1))
'                        Next
                    Else
                        .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))   '赋值到FPS
'                        For j = 0 To UBound(strTmp) - 1
'                            .SetText j + 1, i - 1, Trim$(strTmp(j + 1))
'                        Next
                    End If
                Next
                
            End If
        Next
    End With
    MousePointer = 0  '鼠标状态还原
    
    xlBook.Close      '总是提示是否保存
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
    VBExcel.Quit
    
Exit Sub
ExitPro:
    On Error Resume Next
    MousePointer = 0
    If Not VBExcel Is Nothing Then
        xlBook.Close
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing
        VBExcel.Quit
    End If
    Exit Sub
ErrHandle:
    GoTo ExitPro
End Sub


Private Sub Option1_Click()
Dim sql As String
Dim mainItemRS As New adodb.Recordset
'Dim mainItemRs As adodb.Recordset

sql = "select * from HUAWEI_CARTON "


Set mainItemRS = getStr(sql)

With fps(0)
   .MaxRows = 0
        
    If mainItemRS.RecordCount > 0 Then
        Set .DataSource = mainItemRS
       
    End If
End With
End Sub

Private Sub Option2_Click()
Dim sql As String
Dim mainItemRS As New adodb.Recordset
'Dim mainItemRs As adodb.Recordset

sql = "select * from HUAWEI_LABLE "


Set mainItemRS = getStr(sql)

With fps(0)
   .MaxRows = 0
        
    If mainItemRS.RecordCount > 0 Then
        Set .DataSource = mainItemRS
       
    End If
End With
End Sub
