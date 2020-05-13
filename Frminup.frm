VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frminup 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11970
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
   ScaleHeight     =   6000
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Fra 
      Height          =   4095
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   11895
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   5
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
         SpreadDesigner  =   "Frminup.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Fra 
      Height          =   1455
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txtINV 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton com2 
         Caption         =   "保存发票"
         Height          =   360
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   990
      End
      Begin VB.CommandButton com1 
         Caption         =   "上传发票"
         Height          =   360
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   990
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   1095
         Left            =   6240
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   3495
      End
      Begin VB.CommandButton cmd 
         Caption         =   "退        出"
         Height          =   480
         Index           =   0
         Left            =   9960
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog Com 
         Left            =   11400
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发票号"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径："
         Height          =   195
         Index           =   1
         Left            =   5400
         TabIndex        =   3
         Top             =   360
         Width           =   540
      End
   End
End
Attribute VB_Name = "Frminup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum FpsDetail
    e_PO_NO = 1
    e_PO_LINE = 2
    e_PKG = 3
    e_DEVICE = 4
    e_Lot = 5
    E_Batch = 6
    e_Sublot = 7
    e_DaCode = 8
    e_Qty = 9
    e_Price = 10
    e_USD = 11
    e_MCol = 11
End Enum



Private Sub com1_Click()

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

Private Sub com2_Click()
On Error GoTo ErrHandle
Dim strSql                          As String
Dim strupin                         As String
Dim Rs                              As New adodb.Recordset
Dim i                               As Integer
Dim strTmp(FpsDetail.e_MCol)        As String
Dim inv  As String
Dim upby  As String

upby = gUserName

    '检查资料
    If fps(0).MaxRows <= 0 Then
        MsgBox "没有要保存的资料", vbInformation, "提示"
        Exit Sub
    End If
    
    If MsgBox("是否要保存吗？", vbInformation + vbYesNo, "提示") = vbNo Then Exit Sub
    '如果有资料，开始插入数据库
    '开启事物模式
    MousePointer = 11
    On Error GoTo ErrRollback
    INIadoCon2.BeginTrans
    With fps(0)
        For i = 2 To .MaxRows
            .Row = i
            .Col = FpsDetail.e_PO_NO
            strTmp(0) = Trim$(.Text)
            .Col = FpsDetail.e_PO_LINE
            strTmp(1) = Trim$(.Text)
            .Col = FpsDetail.e_PKG
            strTmp(2) = Trim$(.Text)
            .Col = FpsDetail.e_DEVICE
            strTmp(3) = Trim$(.Text)
            .Col = FpsDetail.e_Lot
            strTmp(4) = Trim$(.Text)
            .Col = FpsDetail.E_Batch
            strTmp(5) = Trim$(.Text)
            .Col = FpsDetail.e_Sublot
            strTmp(6) = Trim$(.Text)
             .Col = FpsDetail.e_DaCode
            strTmp(7) = Trim$(.Text)
             .Col = FpsDetail.e_Qty
            strTmp(8) = Trim$(.Text)
             .Col = FpsDetail.e_Price
            strTmp(9) = Trim$(.Text)
             .Col = FpsDetail.e_USD
            strTmp(10) = Trim$(.Text)
            inv = Trim(txtINV.Text)
            
            
            
            strSql = "select * from erptemp..tblbb_invoice a  where a.invoice_no = '" & inv & "' and a.po_num = '" & strTmp(0) & "' "
            If Rs.State = adStateOpen Then Rs.Close
            Rs.open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            If Not Rs.EOF Then
            strupin = "update erptemp..tblbb_invoice set  po_num = '" & strTmp(0) & "' ,po_line = '" & strTmp(1) & "',sublot = '" & strTmp(6) & "' , PKG = '" & strTmp(2) & "' , device = '" & strTmp(3) & "',  " & _
             " lot = '" & strTmp(4) & "',date_code = '" & strTmp(7) & "', qty = '" & strTmp(8) & "' ,price = '" & strTmp(9) & "' ,usd = '" & strTmp(10) & "' , batch_id = '" & strTmp(5) & "' " & _
            " where invoice_no = '" & inv & "' "
            Else
            strupin = " insert erptemp..tblbb_invoice values ('" & inv & "','" & strTmp(0) & "' ,'" & strTmp(1) & "' ,'" & strTmp(6) & "' ,'" & strTmp(2) & "' ,'" & strTmp(3) & "' ,'" & strTmp(4) & "' ,   " & _
            " '" & strTmp(7) & "' ,'" & strTmp(8) & "' ,'" & strTmp(9) & "' ,'" & strTmp(10) & "',getdate(),'" & upby & "',getdate(),'" & upby & "','" & strTmp(5) & "')  "
            
            End If
            Rs.Close
            
            INIadoCon2.Execute strupin
        Next
    End With
    INIadoCon2.CommitTrans
    MousePointer = 0
    
    MsgBox "资料保存成功！"
    
Exit Sub
ErrRollback:
    '出现错误，事物回滚
    MousePointer = 0
    INIadoCon2.RollbackTrans
    
ErrHandle:
    MousePointer = 0
    MsgBox Err.Description, vbCritical + vbInformation, "警告"

End Sub

Private Sub Form_Load()
    '初始化
    InitCtrl
End Sub

Private Sub InitCtrl()
Dim i               As Integer

    'Fps初始化
    With fps(0)
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


'Form大小自动调整
Private Sub Form_Resize()
    'SSTFPC.Move SSTFPC.Left, SSTFPC.Top, Me.ScaleWidth, Me.ScaleHeight
    Fra(0).Move Fra(0).Left, Fra(0).Top, Me.ScaleWidth - 200, Fra(0).Height
    Fra(1).Move Fra(0).Left, Fra(1).Top, Me.ScaleWidth - 200, Me.ScaleHeight - Fra(0).Height - 500
    fps(0).Move fps(0).Left, fps(0).Top, Fra(1).Width - 300, Fra(1).Height - 300
End Sub



Private Sub FileExportInFps()
On Error GoTo ErrHandle
Dim VBExcel                         As Excel.Application
Dim xlBook                          As Excel.Workbook
Dim xlSheet                         As Excel.Worksheet
Dim strFileName                     As String
Dim inv                             As String
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
    Set xlSheet = xlBook.Worksheets("Sheet1")            '打开sheet中的表
    '判定最大列Excel中的和设定列是否相同
    If xlSheet.Range("A15").CurrentRegion.Columns.Count <> FpsDetail.e_MCol Then
        MousePointer = 0
        MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        GoTo ExitPro
        Exit Sub
    End If
    
    
    txtINV.Text = Trim(xlSheet.Range("K9").Value)
    
    
    '处理ExcelExcel
    With fps(0)
       '  For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.Count  '2)得到Excel最大行
        
        For i = 16 To xlSheet.Range("A65536").End(xlUp).Row
        
        
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



