VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FormWAFERNO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
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
   ScaleHeight     =   6990
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Fra 
      Height          =   1455
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11535
      Begin VB.TextBox txtPath 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   1095
         Left            =   3240
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton cmd 
         Caption         =   "退        出"
         Height          =   480
         Index           =   2
         Left            =   8280
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdWAFER 
         Caption         =   "上传WAFER序号"
         Height          =   480
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "保存WAFER序号"
         Height          =   480
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog Com 
         Left            =   9960
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
      Height          =   5535
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   11415
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
         SpreadDesigner  =   "FormWAFERNO.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "FormWAFERNO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Enum FpsDetail
    e_ID = 1       'wafer
    e_NO = 2         'no
    E_TOTAL = 3         'total

    
    e_MCol = 3
End Enum




Private Sub cmdSAVE_Click(Index As Integer)
On Error GoTo ErrHandle
Dim strSql                          As String
Dim rs                              As New ADODB.Recordset
Dim i                               As Integer
Dim strTmp(FpsDetail.e_MCol)        As String
Dim strsqlup1 As String
Dim strsqlup2 As String
Dim strsqlin1 As String
Dim strsqlin2 As String
Dim qboxvalue As String


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
            .Col = FpsDetail.e_ID
            strTmp(0) = Trim$(.Text)
             .Col = FpsDetail.e_NO
            strTmp(1) = Trim$(.Text)
            .Col = FpsDetail.E_TOTAL
            strTmp(2) = Trim$(.Text)
          
            
            qboxvalue = strTmp(1) + "/" + strTmp(2)
            
            '------------------------------------------------

            '根据箱号查询是否已经上传过此箱号了
            strSql = "select * from mes_reference  a where a.identifier = 'US026_NO_QBOX_WAFER' and a.key1 = '" & strTmp(0) & "' and a.propertyname = 'NO_QBOX_WAFER'"
            If rs.State = adStateOpen Then rs.Close
            rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs.EOF Then  '表示有数据了
            
            MsgBox "数据已存在,请通知IT变更！"
            
            Else
          
            strsqlin2 = "  insert into mes_reference (identifier,key1,key2,key3,propertyname,propertyvalue,flag,creat_by,creat_date ) " & _
            " values ('US026_NO_QBOX_WAFER' ,'" & strTmp(0) & "' ,'NULL','NULL','NO_QBOX_WAFER','" & qboxvalue & "','0','" & gUserName & "',sysdate)"
            Cnn.Execute strsqlin2
            
            End If
            rs.Close
            
'            Dim sOra As String
'            sOra = "select mes_dn_pkg.MES_NCMR_37('" & strTmp(2) & "') from dual"
'
'            Get_OracleRs (sOra)
           
        Next
    End With
    MousePointer = 0
    
    MsgBox "资料保存成功！"
    
Exit Sub
    
ErrHandle:
    MousePointer = 0
    MsgBox Err.Description, vbCritical + vbInformation, "警告"

End Sub

Private Sub cmdWAFER_Click(Index As Integer)

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

Private Sub Form_Resize()
    Fra(0).Move Fra(0).Left, Fra(0).Top, Me.ScaleWidth - 200, Fra(0).Height
    Fra(1).Move Fra(0).Left, Fra(1).Top, Me.ScaleWidth - 200, Me.ScaleHeight - Fra(0).Height - 50
    Fps(0).Move Fps(0).Left, Fps(0).Top, Fra(1).Width - 300, Fra(1).Height - 300
End Sub



