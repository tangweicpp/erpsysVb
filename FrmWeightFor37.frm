VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmWeightFor37 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "37WaferID称重"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
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
   ScaleHeight     =   7830
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra 
      Height          =   735
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton cmdExportOut 
         Caption         =   "导     出"
         Height          =   360
         Left            =   6360
         TabIndex        =   19
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷     新"
         Height          =   360
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退     出"
         Height          =   360
         Left            =   7800
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H000000FF&
         Caption         =   "保     存"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         TabIndex        =   11
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "查     询"
         Height          =   360
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdExportIn 
         Caption         =   "导     入"
         Height          =   360
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   990
      End
      Begin MSComDlg.CommonDialog Com 
         Left            =   9480
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Fra 
      ForeColor       =   &H00FF0000&
      Height          =   7335
      Index           =   1
      Left            =   3840
      TabIndex        =   3
      Top             =   720
      Width           =   9615
      Begin VB.OptionButton Opt 
         Caption         =   "已维护信息"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton Opt 
         Caption         =   "待维护信息"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6015
         _Version        =   524288
         _ExtentX        =   10610
         _ExtentY        =   3836
         _StockProps     =   64
         EditEnterAction =   5
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
         SpreadDesigner  =   "FrmWeightFor37.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   2175
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6015
         _Version        =   524288
         _ExtentX        =   10610
         _ExtentY        =   3836
         _StockProps     =   64
         EditEnterAction =   4
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
         SpreadDesigner  =   "FrmWeightFor37.frx":050B
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "查询条件"
      ForeColor       =   &H00FF0000&
      Height          =   7335
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3735
      Begin VB.TextBox txtPath 
         Height          =   2490
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "双击选择供应商"
         Top             =   1680
         Visible         =   0   'False
         Width           =   3315
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   15
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   250478593
         CurrentDate     =   42739
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "双击选择供应商"
         Top             =   240
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   17
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   250478593
         CurrentDate     =   42739
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wafer   ID"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   750
      End
   End
End
Attribute VB_Name = "FrmWeightFor37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum FpsDetail
    e_WaferID = 1               'WaferID
    e_Weight                    '重量
    e_Stand                     '标准
    e_NUM                       '数量
    e_Cust                      '客户
    e_MCol
End Enum
'退出
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExportIn_Click()
On Error GoTo ErrHandler

Dim FName
    '筛选文件
    com.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
    com.ShowOpen
    '得到文件名
    FName = com.filename
    If FName <> "" Then
        txtPath.Text = FName
       '将资料写到FPS
       FileExportInFps
    End If
    
Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub
'导入资料
Private Sub FileExportInFps()
On Error GoTo ErrHandle
Dim VBExcel                         As Excel.Application
Dim xlBook                          As Excel.Workbook
Dim xlSheet                         As Excel.Worksheet
Dim strFilename                     As String
Dim I                               As Integer
Dim J                               As Integer
Dim strChar                         As String
Dim strTmp(FpsDetail.e_MCol - 1)    As Variant
    
    MousePointer = 11
    '清空FPS(0)
    Fps(0).ClearRange FpsDetail.e_WaferID, FpsDetail.e_WaferID, Fps(0).MaxCols, Fps(0).MaxRows, True
    '获取文件名
    If InStrRev(Trim(txtPath.Text), "\") > 0 Then
        strFilename = Mid(Trim(txtPath.Text), InStrRev(Trim(txtPath.Text), "\") + 1)
        If InStr(strFilename, ".") > 0 Then
            strFilename = Mid(strFilename, 1, InStr(strFilename, ".") - 1)
        End If
    End If
    'Excel文件处理
    '1)打开Excel
    Set VBExcel = CreateObject("excel.application")      '创建Excle对象
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(txtPath.Text)    '打开文件
    Set xlSheet = xlBook.Worksheets(1)            '打开sheet中的表
    '判定最大列Excel中的和设定列是否相同
    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> FpsDetail.e_MCol - 1 Then
        MousePointer = 0
        MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        GoTo EXITPRO
        Exit Sub
    End If
    '处理ExcelExcel
    With Fps(0)
        For I = 1 To xlSheet.Range("A1").CurrentRegion.Rows.Count         '2)得到Excel最大行
            strTmp(0) = Trim(xlSheet.Range("A" & I).Value)
            If Len(strTmp(0)) > 0 Then
                For J = 1 To FpsDetail.e_MCol - 1
                    '循环i,j 26(得到A.B.C.)
                    If J > 26 Then
                        strChar = Chr(96 + Int(J / 26 - 0.001)) & IIf(J Mod 26 = 0, "Z", Chr(96 + (J Mod 26)))
                    Else
                        strChar = Chr(96 + J)
                    End If
'                    strTmp(j) = xlSheet.Range(strChar & i).Value   '先屏蔽，换方法写
                    If I = 1 Then '得到第一行
'                        .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))  '赋值到FPS
'                        For j = 0 To UBound(strTmp) - 1
'                            .SetText j + 1, i - 1, Trim$(strTmp(j + 1))
'                        Next
                    Else
                        .SetText J, I - 1, Trim$(xlSheet.Range(strChar & I)) '赋值到FPS
'                        For j = 0 To UBound(strTmp) - 1
'                            .SetText j + 1, i - 1, Trim$(strTmp(j + 1))
'                        Next
                    End If
                Next
                
            End If
        Next
    End With
    MousePointer = 0  '鼠标状态还原
    
    MsgBox "导入成功！"
    
    xlBook.Close      '总是提示是否保存
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
    
Exit Sub
EXITPRO:
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
    MousePointer = 0  '鼠标状态还原
    MsgBox "执行失败！" + Chr(13) + "原因:" + Err.Description, vbInformation, Me.Caption
    GoTo EXITPRO
End Sub

Private Sub cmdExportOut_Click()
    If Opt(0).Value = True Then
        Export 0
    Else
        Export 1
    End If
End Sub

'刷新
Private Sub cmdRefresh_Click()
    If Opt(0).Value = True Then
        Fps(0).ClearRange FpsDetail.e_WaferID, FpsDetail.e_WaferID, Fps(0).MaxCols, Fps(0).MaxRows, True
    Else
        Fps(1).MaxRows = 0
    End If
End Sub

Private Sub CmdSave_Click()
    '校验数据
    If Not CheckData Then Exit Sub
    '保存数据到资料库
    saveData
End Sub
'查询
Private Sub cmdSearch_Click()
    Call Search(IIf(Opt(0).Value = True, 0, 1))
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Fra(1).Move Fra(1).Left, Fra(1).Top, Me.ScaleWidth - Fra(0).Width - 120, Me.ScaleHeight - Fra(2).Height - 120
    Fra(2).Move 60, Fra(2).Top, Me.ScaleWidth - 120, Fra(2).Height
    Fra(0).Move 60, Fra(0).Top, Fra(0).Width, Me.ScaleHeight - Fra(2).Height - 120
    Fps(0).Move 60, Fps(0).Top, Fra(1).Width - 120, Me.ScaleHeight - Fra(2).Height - 4 * 120
    Fps(1).Move 60, Fps(0).Top, Fra(1).Width - 120, Me.ScaleHeight - Fra(2).Height - 4 * 120
    
End Sub
Private Sub Form_Load()
    '初始化控件
    InitCtrl
End Sub
'初始化控件
Private Sub InitCtrl()
Dim I                   As Integer
Dim strsql              As String
Dim rs                  As New ADODB.Recordset

    'Fps初始化
    With Fps(0)
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsBest
        .MaxRows = 500
        .ColsFrozen = 1
        .ButtonDrawMode = 1
        .MaxCols = FpsDetail.e_MCol - 1
        .Row = -1
        .Col = -1
        .Lock = True
        '设定列类型
        .Col = FpsDetail.e_WaferID
        .Lock = False
        .Col = FpsDetail.e_Weight
        .CellType = CellTypeNumber
        .TypeNumberDecPlaces = 4
        .Lock = False
        .Col = FpsDetail.e_Stand
        .TypeHAlign = TypeHAlignRight
        .TypeVAlign = TypeVAlignCenter
'        .CellType = CellTypeNumber
'        .TypeNumberDecPlaces = 6
        .Col = FpsDetail.e_NUM
        .CellType = CellTypeNumber
        .TypeNumberShowSep = True
        .TypeNumberSeparator = ","
        .TypeNumberDecPlaces = 0
        .Col = FpsDetail.e_Cust
        .CellType = CellTypeComboBox
        .TypeComboBoxList = "37"
'        .TypeComboBoxList = .TypeComboBoxList & "68"
'        .TypeComboBoxList = .TypeComboBoxList & "95"
        .TypeHAlign = TypeHAlignRight
        .TypeVAlign = TypeVAlignCenter
        .SetText FpsDetail.e_Cust, -1, "37"
        '设定可编辑边框颜色
        .SetCellBorder FpsDetail.e_WaferID, -1, FpsDetail.e_Weight, -1, 15, vbBlue, CellBorderStyleDot
        '设置列头
        .SetText FpsDetail.e_WaferID, 0, "Wafer ID"
        .SetText FpsDetail.e_Weight, 0, "重量"
        .SetText FpsDetail.e_Stand, 0, "标准重量"
        .SetText FpsDetail.e_NUM, 0, "数量"
        .SetText FpsDetail.e_Cust, 0, "客户代码"
        '设置默认值
        .SetText FpsDetail.e_Stand, -1, "0.000106"
        
        '设定列宽
        .ColWidth(-1) = 10
        .RowHeight(-1) = 15
'        '设定是否排序
'        .UserColAction = UserColActionSort
'        For i = 1 To .MaxCols
'            .Col = i
'            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
'        Next
'        .ZOrder
        .ReDraw = True
    End With
    
    With Fps(1)
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsBest
        .MaxRows = 0
        .Row = -1
        .Col = -1
        .Lock = True
        .Col = FpsDetail.e_NUM
        .CellType = CellTypeNumber
        .TypeNumberDecPlaces = 0
        .TypeNumberShowSep = True
        .TypeNumberSeparator = ","
        
        .ColWidth(-1) = 10
        .RowHeight(-1) = 15
        '设定是否排序
        .UserColAction = UserColActionSort
        For I = 1 To .MaxCols
            .Col = I
            .ColUserSortIndicator(I) = ColUserSortIndicatorAscending
        Next
        .Visible = False
        .ZOrder
        .ReDraw = True
    End With
    
    DTP(0).Value = Format(Now(), "YYYY/MM/01")
    DTP(1).Value = Format(Now(), "YYYY/MM/DD")
    
End Sub

Private Sub fps_Change(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim DblWeight       As Double
Dim DblStand        As Double
Dim DblNum          As Double

    If Index = 0 Then
        If Col <> FpsDetail.e_Weight Then Exit Sub '编辑重量触发事件
        With Fps(Index)
            .Row = Row
            .Col = FpsDetail.e_Weight   '重量
            DblWeight = Val(.Text)
            .Col = FpsDetail.e_Stand    '标准重量
            DblStand = IIf(Val(.Text) = 0, 1, Val(.Text))
            DblNum = DblWeight / DblStand '得到数量
            .SetText FpsDetail.e_NUM, Row, DblNum
        End With
    End If
End Sub

'Fps编辑事件
Private Sub Fps_EditChange(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim DblWeight       As Double
Dim DblStand        As Double
Dim DblNum          As Double

    If Index = 0 Then
        If Col <> FpsDetail.e_Weight Then Exit Sub '编辑重量触发事件
        With Fps(Index)
            .Row = Row
            .Col = FpsDetail.e_Weight   '重量
            DblWeight = Val(.Text)
            .Col = FpsDetail.e_Stand    '标准重量
            DblStand = IIf(Val(.Text) = 0, 1, Val(.Text))
            DblNum = DblWeight / DblStand '得到数量
            .SetText FpsDetail.e_NUM, Row, DblNum
        End With
    End If
End Sub

Private Sub Opt_Click(Index As Integer)
    If Index = 0 Then
        Fps(0).Visible = True
        Fps(1).Visible = False
        cmdExportIn.Enabled = True
        cmdSave.Enabled = True
    Else
        Fps(0).Visible = False
        Fps(1).Visible = True
        cmdExportIn.Enabled = False
        cmdSave.Enabled = False
    End If
End Sub

'Private Sub Fps_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Index = 1 Then
'        If KeyCode = 46 Then    '表示点下了del键
'            If MsgBox("是否要删除该行吗？", vbInformation + vbDefaultButton2 + vbYesNo, "提示") = vbNo Then Exit Sub
'            With Fps(0)
'                If .MaxRows <= 0 Then Exit Sub
'                Set .DataSource = Nothing
'                .DeleteRows .ActiveRow, 1
'                .MaxRows = .MaxRows - 1
'            End With
'        End If
'    End If
'End Sub

'导出数据
Public Sub Export(intBJ As Integer)
    If intBJ = 0 Then
        If Not ExportFpspreadToExcel(Fps(intBJ), "待维护信息", "待维护信息") Then Exit Sub
    Else
        If Not ExportFpspreadToExcel(Fps(intBJ), "已维护信息", "已维护信息") Then Exit Sub
    End If
End Sub
'校验数据
Private Function CheckData() As Boolean
On Error GoTo ErrHandle
Dim I               As Integer
Dim J               As Integer
Dim strsql          As String
Dim rs              As New ADODB.Recordset
Dim strTmp(4)       As String
Dim strWaferID      As String
Dim strWaferInfo    As String

    CheckData = False
    strWaferID = ""
    Screen.MousePointer = 11
    With Fps(0)
        If .MaxRows <= 0 Then Exit Function
        For I = 1 To .MaxRows
            .Row = I
            .Col = FpsDetail.e_WaferID          'wafer id
            strTmp(0) = Replace(Replace(Trim$(.Text), vbCrLf, ""), "'", "")
            
            .Col = FpsDetail.e_Weight           '重量
            strTmp(1) = Val(.Text)
            .Col = FpsDetail.e_Stand            '标准
            strTmp(2) = Val(.Text)
            .Col = FpsDetail.e_NUM              '数量
            strTmp(3) = Val(.Text)
            .Col = FpsDetail.e_Cust             '客户
            strTmp(4) = Trim$(.Text)
            'Wafer id 不能为空
            If Len(strTmp(0)) > 0 Then
                '查询数据库进行校验
'                strSql = "select containername from container Where containername='" & strTmp(0) & "'"
'                If rs.State = adStateOpen Then rs.Close
'                rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
'                If rs.EOF Then  '如果Mes没有数据表示不存在 WaferID
'                    Screen.MousePointer = 0
'                    MsgBox "第" & i & "行的WaferID:" & strTmp(0) & "在MES系统中不存在！"
'                    Exit Function
'                End If
'                rs.Close
                '单行是否维护正确
                If strTmp(1) <= 0 Then          '重量
                    Screen.MousePointer = 0
                    MsgBox "第" & I & "行的WaferID:" & strTmp(1) & "没有重量！"
                    Exit Function
                End If
                If strTmp(2) <= 0 Then          '标准
                    Screen.MousePointer = 0
                    MsgBox "第" & I & "行的WaferID:" & strTmp(2) & "没有标准重量！"
                    Exit Function
                End If
                If strTmp(3) <= 0 Then          '数量
                    Screen.MousePointer = 0
                    MsgBox "第" & I & "行的WaferID:" & strTmp(3) & "没有数量！"
                    Exit Function
                End If
                If strTmp(4) <= 0 Then          '客户
                    Screen.MousePointer = 0
                    MsgBox "第" & I & "行的WaferID:" & strTmp(4) & "没有客户！"
                    Exit Function
                End If
                '记录所有的WaferID,方面后面校验是否存在数据库
                strWaferID = strWaferID + strTmp(0) + ","
                '内层循环
                For J = I + 1 To .MaxRows
                    .Row = J
                    .Col = FpsDetail.e_WaferID
                    If strTmp(0) = Replace(Replace(Trim$(.Text), vbCrLf, ""), "'", "") Then    '如果从第一行的WaferID 到下面有重复就提示错误
                        Screen.MousePointer = 0
                        MsgBox "第" & I & "行的WaferID:" & strTmp(0) & "和第" & J & "行的WaferID相同！"
                        Exit Function
                    End If
                Next J
            End If
        Next I
    End With
    '截取WaferID
    If Len(strWaferID) > 0 Then
        strWaferID = Mid$(strWaferID, 1, Len(strWaferID) - 1)
        '查询数据库进行校验
        strsql = "Select WaferID From Weight37 Where WaferID In('" & Replace$(strWaferID, ",", "','") & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not rs.EOF Then  '如果有数据表示有存在的WaferID
            Do While Not rs.EOF
                strWaferInfo = strWaferInfo + Trim$("" & rs!WaferID) + ","
                rs.MoveNext
            Loop
        End If
        rs.Close
        If Len(strWaferInfo) > 0 Then
            Screen.MousePointer = 0
            strWaferInfo = Mid$(strWaferInfo, 1, Len(strWaferInfo) - 1)
            MsgBox "WaferID:" & strWaferInfo & "已经存在数据库中，不能重复保存数据！"
            Exit Function
        End If
    End If
    
    CheckData = True
    Screen.MousePointer = 0
    
Exit Function
ErrHandle:
    CheckData = False
    Screen.MousePointer = 0
    MsgBox "执行失败！" + Chr(13) + "原因:" + Err.Description, vbInformation, Me.Caption
End Function

'保存数据到资料库中
Public Sub saveData()

    On Error GoTo ErrHandle

    Dim I           As Integer

    Dim strsql      As String

    Dim strsql2     As String

    Dim rs          As New ADODB.Recordset

    Dim strTmp(4)   As String

    Dim bln         As Boolean

    Dim strDatecode As String

    bln = False
    Screen.MousePointer = 11

    With Fps(0)

        If .MaxRows <= 0 Then Exit Sub

        For I = 1 To .MaxRows
            .Row = I
            .Col = FpsDetail.e_WaferID          'wafer id
            strTmp(0) = Trim$(.Text)
            .Col = FpsDetail.e_Weight           '重量
            strTmp(1) = Val(.Text)
            .Col = FpsDetail.e_Stand            '标准
            strTmp(2) = Val(.Text)
            .Col = FpsDetail.e_NUM              '数量
            strTmp(3) = Val(Replace$(.Text, ",", ""))
            .Col = FpsDetail.e_Cust             '客户
            strTmp(4) = Trim$(.Text)

            'Wafer id 不能为空
            If Len(strTmp(0)) > 0 Then
                bln = True
                '插入资料库
                strsql = "Insert Into weight37(WAFERID,WEIGHT,STANDWEIGHT,DIE,CUSTOMER) Values('" & strTmp(0) & "','" & strTmp(1) & "','" & strTmp(2) & "','" & strTmp(3) & "','" & strTmp(4) & "')"
            
                Cnn.Execute strsql
                
                Dim sOra As String
        
                sOra = "select mes_dn_pkg.MES_WEIGHT_37('" & strTmp(0) & "') from dual"
                AddSql (sOra)
        
                strDatecode = Get_OracleStr("select case when create_date >= to_date(to_char(create_date,'yyyy') || '-12-31','yyyy-mm-dd') - mod(to_char(create_date,'YYYY'),7) - 5 " & " then to_char(create_date,'yyww') else to_char(create_date +  mod(mod(to_char(create_date,'YYYY'),7) + 5,7),'yyww') end as PODATECODE from weight37 where waferid = '" & strTmp(0) & "'")
                
                ' strDateCode = Get_OracleStr("select to_char(create_date+1,'YYWW') from weight37 where WAFERID = '" & strTmp(0) & "'")
                strsql2 = "insert into erpbase..WEIGHT37(WAFERID,CREATE_DATE) values('" & strTmp(0) & "', '" & strDatecode & "') "
                AddSql2 (strsql2)
            
            End If

        Next

    End With

    If bln = True Then
        MsgBox "资料保存成功！", vbInformation, "提示"

    End If
    
    Screen.MousePointer = 0
    Exit Sub
ErrHandle:
    Screen.MousePointer = 0
    MsgBox "执行失败！" + Chr(13) + "原因:" + Err.Description, vbInformation, Me.Caption

End Sub

'查询数据
Public Sub Search(intBJ As Integer)
On Error GoTo ErrHandle
Dim I               As Long
Dim J               As Integer
Dim rs              As New ADODB.Recordset
Dim strsql          As String

    Screen.MousePointer = 11
    Fps(intBJ).MaxRows = 0
    If intBJ = 0 Then '待维护信息
        Screen.MousePointer = 0
        MsgBox "不需查询，直接输入信息，保存即可！"
        Exit Sub
    Else
        strsql = "Select * From weight37 Where Create_Date>=to_date('" & DTP(0).Value & "','YYYY/MM/DD') And  Create_Date<to_date('" & DTP(1).Value + 1 & "','YYYY/MM/DD') "
    End If
    If txt(0).Text <> "" Then
        strsql = strsql & " And WAFERID like '" & Trim(txt(0).Text) & "%'"
    End If
    
    rs.Open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        With Fps(intBJ)
            Set .DataSource = rs
            .MaxRows = rs.RecordCount
        End With
    Else
        Screen.MousePointer = 0
        MsgBox "无资料信息！", vbInformation, "提示"
        Exit Sub
    End If
    rs.Close
    Screen.MousePointer = 0
    
Exit Sub
ErrHandle:
    Screen.MousePointer = 0
    MsgBox "执行失败！" + Chr(13) + "原因:" + Err.Description, vbInformation, Me.Caption
End Sub






