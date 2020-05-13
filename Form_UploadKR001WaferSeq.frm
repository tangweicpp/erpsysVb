VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Form_UploadKR001WaferSeq 
   Caption         =   "KR001Wafer序号维护"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12540
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
   ScaleHeight     =   9600
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   8055
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   12495
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   7455
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   11655
         _Version        =   524288
         _ExtentX        =   20558
         _ExtentY        =   13150
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
         SpreadDesigner  =   "Form_UploadKR001WaferSeq.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin MSComDlg.CommonDialog com 
         Left            =   2880
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   10440
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtPath 
         Height          =   1215
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Width           =   6255
      End
      Begin VB.CommandButton cmdSaveWAFER 
         Caption         =   "保存WAFER序号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton cmdAddWAFER 
         Caption         =   "上传WAFER序号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form_UploadKR001WaferSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum FpsDetail
    e_Wafer = 1       'LOT号
    e_NO = 2         'NCMR
    e_Total = 3         'WAFER

    e_MCol = 3
End Enum

Private Sub cmdAddWAFER_Click()
On Error GoTo ErrHandler

Dim FName
    '筛选文件
    com.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
    com.ShowOpen
    '得到文件名
    FName = com.filename
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

Private Sub cmdExit_Click()
Unload Me
End Sub

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
    Set xlBook = VBExcel.Workbooks.Open(txtPath.Text)    '打开文件
    Set xlSheet = xlBook.Worksheets("Sheet1")            '打开sheet中的表
    '判定最大列Excel中的和设定列是否相同
    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> FpsDetail.e_MCol Then
        MousePointer = 0
        MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        GoTo EXITPRO
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

                    Else
                        .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))   '赋值到FPS
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
    GoTo EXITPRO
End Sub

Private Sub cmdSaveWAFER_Click()
On Error GoTo ErrHandle
Dim strSql                          As String
Dim rs                              As New ADODB.Recordset
Dim i                               As Integer
Dim strTmp(FpsDetail.e_MCol)        As String
Dim strsqlup1 As String
Dim strsqlup2 As String
Dim strsqlin1 As String
Dim strsqlin2 As String



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
            .Col = FpsDetail.e_Wafer
            strTmp(0) = Trim$(.Text)
             .Col = FpsDetail.e_NO
            strTmp(1) = Trim$(.Text)
            .Col = FpsDetail.e_Total
            strTmp(2) = Trim$(.Text)
          
            
            '------------------------------------------------

            '根据箱号查询是否已经上传过此箱号了
            strSql = "select * from  mes_reference a where a.KEY1 = '" & strTmp(0) & "'"
            If rs.State = adStateOpen Then rs.Close
            rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs.EOF Then  '表示有数据了
    
            strsqlup2 = " update mes_reference set PROPERTYVALUE =  '" & strTmp(1) & "' || '/' || '" & strTmp(2) & "'  where KEY1 = '" & strTmp(0) & "' "
        
            Cnn.Execute strsqlup2
            
            Else
    
            strsqlin2 = "  insert into mes_reference (IDENTIFIER,KEY1,KEY2, KEY3,PROPERTYNAME,PROPERTYVALUE,FLAG,CREAT_BY, CREAT_DATE ) values ('US026_NO_QBOX_WAFER' ,'" & strTmp(0) & "' ,'NULL','NULL','NO_QBOX_WAFER', '" & strTmp(1) & "' || '/' || '" & strTmp(2) & "','0','" & gUserName & "', '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "')"
          
            Cnn.Execute strsqlin2
            
            End If
            rs.Close
    
        Next
    End With
    MousePointer = 0
    
    MsgBox "资料保存成功！"
    
Exit Sub
    
ErrHandle:
    MousePointer = 0
    MsgBox Err.Description, vbCritical + vbInformation, "警告"
End Sub
