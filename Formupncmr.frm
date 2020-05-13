VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Formupncmr 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10770
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
   ScaleHeight     =   5625
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Fra 
      Height          =   4095
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   10455
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
         SpreadDesigner  =   "Formupncmr.frx":0000
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
      Width           =   10455
      Begin VB.CommandButton cmdNCMR2 
         Caption         =   "保存NCMR"
         Height          =   480
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdNCMR1 
         Caption         =   "上传NCMR"
         Height          =   480
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         Caption         =   "退        出"
         Height          =   480
         Index           =   2
         Left            =   8280
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   1095
         Left            =   3240
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4935
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
         TabIndex        =   3
         Top             =   360
         Width           =   540
      End
   End
End
Attribute VB_Name = "Formupncmr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum fpSDetail
    e_Lot = 1       'LOT号
    e_NCMR = 2         'NCMR
    E_Wafer = 3         'WAFER

    
    e_MCol = 3
End Enum

Private Sub cmdNCMR1_Click()
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

Private Sub FileExportInFps()
On Error GoTo ErrHandle
Dim VBExcel                         As Excel.Application
Dim xlBook                          As Excel.Workbook
Dim xlSheet                         As Excel.Worksheet
Dim StrFileName                     As String
Dim i                               As Integer
Dim J                               As Integer
Dim strChar                         As String
Dim strTmp(fpSDetail.e_MCol)        As Variant
    
    MousePointer = 11
    'Fps
    Fps(0).MaxRows = 0
    '获取文件名
    If InStrRev(Trim(txtPath.Text), "\") > 0 Then
        StrFileName = Mid(Trim(txtPath.Text), InStrRev(Trim(txtPath.Text), "\") + 1)
        If InStr(StrFileName, ".") > 0 Then
            StrFileName = Mid(StrFileName, 1, InStr(StrFileName, ".") - 1)
        End If
    End If
    'Excel文件处理
    '1)打开Excel
    Set VBExcel = CreateObject("excel.application")      '创建Excle对象
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(txtPath.Text)    '打开文件
    Set xlSheet = xlBook.Worksheets("Sheet1")            '打开sheet中的表
    '判定最大列Excel中的和设定列是否相同
    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> fpSDetail.e_MCol Then
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
                For J = 1 To fpSDetail.e_MCol
                    '循环i,j 26(得到A.B.C.)
                    If J > 26 Then
                        strChar = Chr(96 + Int(J / 26 - 0.001)) & IIf(J Mod 26 = 0, "Z", Chr(96 + (J Mod 26)))
                    Else
                        strChar = Chr(96 + J)
                    End If
'                    strTmp(j) = xlSheet.Range(strChar & i).Value   '先屏蔽，换方法写
                    If i = 1 Then '得到第一行
                        .SetText J, .MaxRows, Trim$(xlSheet.Range(strChar & i))  '赋值到FPS
'                        For j = 0 To UBound(strTmp) - 1
'                            .SetText j + 1, i - 1, Trim$(strTmp(j + 1))
'                        Next
                    Else
                        .SetText J, .MaxRows, Trim$(xlSheet.Range(strChar & i))   '赋值到FPS
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

Private Sub cmdNCMR2_Click()
On Error GoTo ErrHandle
Dim strSql                          As String
Dim rs                              As New ADODB.Recordset
Dim i                               As Integer
Dim strTmp(fpSDetail.e_MCol)        As String
Dim strsqlup1 As String
Dim strsqlup2 As String
Dim strSqlin1 As String
Dim strsqlin2 As String
Dim strSqlUp  As String


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
            .Col = fpSDetail.e_Lot
            strTmp(0) = Trim$(.Text)
             .Col = fpSDetail.e_NCMR
            strTmp(1) = Trim$(.Text)
            .Col = fpSDetail.E_Wafer
            strTmp(2) = Trim$(.Text)
          
            
            '------------------------------------------------

            '根据箱号查询是否已经上传过此箱号了
            strSql = "select * from ERPBASE..TBLWAREHOUSEDB_INFO  a where a.wafer_id = '" & strTmp(2) & "'"
            If rs.State = adStateOpen Then rs.Close
            rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs.EOF Then  '表示有数据了
'           strsqlup1 = " Update ERPBASE..TBLWAREHOUSEDB_INFO set Comment = '" & strTmp(1) & "' where wafer_id = '" & strTmp(2) & "'"
            
            strsqlup1 = " Update ERPBASE..TBLWAREHOUSEDB_INFO set Comment = '" & strTmp(1) & "' + ';' +  replace(Comment,'" & strTmp(1) & "','')   where wafer_id = '" & strTmp(2) & "'"
            strsqlup2 = " update pj_ncmr set ncmr =  '" & strTmp(1) & "'  where wafer_id = '" & strTmp(2) & "' "
          
    
            
            strSqlUp = "update ERPBASE..TBLWAREHOUSEDB_INFO set Comment = REPLACE(Comment,';;',';')  where wafer_id = '" & strTmp(2) & "' "
            AddSql2 (strSqlUp)
           
            INIadoCon2.Execute strsqlup1
            Cnn.Execute strsqlup2
            
            Else
            strSqlin1 = "insert into ERPBASE..TBLWAREHOUSEDB_INFO ( HTLOTID, Comment,wafer_id ,flag)  values ('" & strTmp(0) & "' ,'" & strTmp(1) & "' ,'" & strTmp(2) & "','Y')"
            strsqlin2 = "  insert into pj_ncmr (lot_id,ncmr,wafer_id,flag ) values ('" & strTmp(0) & "' ,'" & strTmp(1) & "' ,'" & strTmp(2) & "','Y')"
            INIadoCon2.Execute strSqlin1
            Cnn.Execute strsqlin2
            
            End If
            rs.Close
            
            Dim sOra As String
            sOra = "select mes_dn_pkg.MES_NCMR_37('" & strTmp(2) & "') from dual"
            
            AddSql (sOra)
           
        Next
    End With
    MousePointer = 0
    
    MsgBox "资料保存成功！"
    
Exit Sub
    
ErrHandle:
    MousePointer = 0
    MsgBox Err.DESCRIPTION, vbCritical + vbInformation, "警告"
End Sub


Private Sub Form_Resize()
    Fra(0).Move Fra(0).Left, Fra(0).Top, Me.ScaleWidth - 200, Fra(0).Height
    Fra(1).Move Fra(0).Left, Fra(1).Top, Me.ScaleWidth - 200, Me.ScaleHeight - Fra(0).Height - 50
    Fps(0).Move Fps(0).Left, Fps(0).Top, Fra(1).Width - 300, Fra(1).Height - 300
End Sub


