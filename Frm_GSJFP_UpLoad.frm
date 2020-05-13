VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_GSJFP_UpLoad 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "国税局发票上传"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
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
   ScaleHeight     =   7140
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra 
      Height          =   1455
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton cmdExcelIn 
         Caption         =   "国税发票上传"
         Height          =   480
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "国税发票保存"
         Height          =   480
         Left            =   480
         TabIndex        =   5
         Top             =   840
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
         Height          =   480
         Left            =   8280
         TabIndex        =   3
         Top             =   480
         Width           =   1695
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
      Height          =   4095
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   10455
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6975
         _Version        =   524288
         _ExtentX        =   12303
         _ExtentY        =   3413
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
         SpreadDesigner  =   "Frm_GSJFP_UpLoad.frx":0000
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
Attribute VB_Name = "Frm_GSJFP_UpLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum FpsDetail
    e_Invoice = 1       '国税局发票
    
    e_FHDH = 2          '发货单号
    e_Qty = 3          '数量
    e_Unit = 4         '单位
    e_Price = 5        '单价
    e_JE = 6            '金额
    e_BB = 7            '币别
    e_HL = 8            '汇率
    e_SL = 9            '税率
    e_Invoice1 = 10       '销售发票
    
    e_MCol = 10
End Enum

Private Sub cmdExcelIn_Click()
On Error GoTo ErrHandler

Dim FName
    '筛选文件
    Com.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
    Com.ShowOpen
    '得到文件名
    FName = Com.FileName
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
    With Fps(0)
        .ReDraw = False
        .DAutoSizeCols = DAutoSizeColsBest
        .MaxRows = 0
        .MaxCols = FpsDetail.e_MCol
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

Private Sub cmdSave_Click() '资料保存
On Error GoTo ErrHandle
Dim strSql                          As String
Dim Rs                              As New adodb.Recordset
Dim i                               As Integer
Dim strTmp(FpsDetail.e_MCol)        As String


    '检查资料
    If Fps(0).MaxRows <= 0 Then
        MsgBox "没有要保存的资料", vbInformation, "提示"
        Exit Sub
    End If
    
    If MsgBox("是否要保存吗？", vbInformation + vbYesNo, "提示") = vbNo Then Exit Sub
    '如果有资料，开始插入数据库
    '开启事物模式
    MousePointer = 11
    With Fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsDetail.e_Invoice          'Invoice Number
            strTmp(0) = Trim$(.Text)
             .Col = FpsDetail.e_Invoice1          'Invoice Number
            strTmp(9) = Trim$(.Text)
            .Col = FpsDetail.e_FHDH             '发货单号
            strTmp(1) = Trim$(.Text)
            .Col = FpsDetail.e_Qty              '数量
            strTmp(2) = Val(Trim$(.Text))
            .Col = FpsDetail.e_Unit             '单位
            strTmp(3) = Trim$(.Text)
            .Col = FpsDetail.e_Price            '单价
            strTmp(4) = Val(.Text)
            .Col = FpsDetail.e_JE               '金额
            strTmp(5) = Val(.Text)
            .Col = FpsDetail.e_BB               '币别
            strTmp(6) = Trim(.Text)
            .Col = FpsDetail.e_HL               '汇率
            strTmp(7) = Val(.Text)
            .Col = FpsDetail.e_SL               '税率
            strTmp(8) = Val(.Text)
            '判定有没有数量-----------------------
            If strTmp(2) <= 0 Then
                MousePointer = 0
                MsgBox "第" + Trim(i) + "行数量为0，不能保存！", vbInformation, "提示"
                Exit Sub
            End If
            
            '------------------------------------------------

            '根据箱号查询是否已经上传过此箱号了
            strSql = "Select * From erptemp..tblBB_GSJFP Where Invoice_Number='" & Trim(strTmp(0)) & "' And Send_Number='" & Trim(strTmp(1)) & "'"
            If Rs.State = adStateOpen Then Rs.Close
            Rs.open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            If Not Rs.EOF Then  '表示有数据了
                MousePointer = 0
                MsgBox "第" + Trim(i) + "行Invoice_Number:" + Trim(strTmp(0)) + ",发货单号:" + Trim(strTmp(1)) + "已经上传过了，不能重复上传！", vbInformation, "提示"
                Exit Sub
            End If
            Rs.Close
            
            '校验完毕，插入数据库
            If Val(strTmp(2)) > 0 Then
                strSql = "Insert Into erptemp..tblBB_GSJFP(Invoice_Number,Send_Number,数量,单位,单价,金额,币别,汇率,税率, Create_by,sale_Invoice_Number) " & _
                         " Values('" & strTmp(0) & "','" & strTmp(1) & "'," & strTmp(2) & ",'" & strTmp(3) & "'," & strTmp(4) & "," & strTmp(5) & ",'" & strTmp(6) & "'," & strTmp(7) & "," & strTmp(8) & ",'" & gUserName & "','" & strTmp(9) & "')"
                INIadoCon2.Execute strSql
            End If
            
        Next
    End With
    MousePointer = 0
    
    MsgBox "资料保存成功！"
    
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
    Fps(0).Move Fps(0).Left, Fps(0).Top, Fra(1).Width - 300, Fra(1).Height - 300
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
    Fps(0).MaxRows = 0
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
    Set xlSheet = xlBook.Worksheets("Sheet1")            '打开sheet中的表
    '判定最大列Excel中的和设定列是否相同
    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> FpsDetail.e_MCol Then
        MousePointer = 0
        MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        GoTo ExitPro
        Exit Sub
    End If
    '处理ExcelExcel
    With Fps(0)
        For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.Count         '2)得到Excel最大行
            strTmp(0) = Trim(xlSheet.Range("A" & i).Value)
            If Len(strTmp(0)) > 0 Then
                If i <> 1 Then .MaxRows = .MaxRows + 1  '第一行表示标题，不用增加行
                For j = 1 To FpsDetail.e_MCol
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


