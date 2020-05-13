VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Frm_B2BUpload 
   Caption         =   "上传US026工单信息"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16080
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Frm_B2BUpload.frx":0000
      Left            =   1920
      List            =   "Frm_B2BUpload.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1043
      Width           =   1455
   End
   Begin VB.Frame Fra 
      Caption         =   "上传"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   960
      TabIndex        =   0
      Top             =   2880
      Width           =   11415
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H0080C0FF&
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpload 
         BackColor       =   &H00FF8080&
         Caption         =   "上传"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtText1 
         Height          =   405
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   4935
      End
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H0080FF80&
         Caption         =   "导出报表"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00FF00FF&
         Caption         =   ".."
         Height          =   405
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   495
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8640
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   5
         Top             =   682
         Width           =   555
      End
   End
   Begin VB.Label lblCusType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户类型:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Width           =   1035
   End
End
Attribute VB_Name = "Frm_B2BUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()

    On Error Resume Next

    Dim FName

    '帅选文件
    CommonDialog1.Filter = "CSV文件(*.csv)|*.csv|XLS文件(*.xls)|*.xls"
    
    CommonDialog1.ShowOpen
    '得到文件名
    FName = CommonDialog1.filename

    If FName <> "" Then
        txtText1.Text = FName

    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me

End Sub

Private Sub cmdExport_Click()

    If Combo1.Text = "" Then
        MsgBox "请选择客户类型", vbInformation, "友情提示"
        Exit Sub

    End If

    If Combo1.Text = "US026工单模板" Then
        ExporToExcel ("SELECT * " & "FROM B2B_ORDERTBL order by upload_time desc ")
    Else
        ExporToExcel ("SELECT * " & "FROM US026_INVOICE ")

    End If

End Sub

Private Sub UpUS026FP()

    Dim source_batch_id_Temp As String

    Dim dirName              As String

    Dim filename             As String

    Dim xlApp                As Excel.Application

    Dim xlBook               As Excel.Workbook

    Dim xlSheet              As Excel.Worksheet

    Dim bT                   As US026INVOICE

    Dim I                    As Integer

    Dim J                    As Integer

    Dim id                   As Long

    Dim temp                 As String

    Dim temp2                As String

    Dim tempVal              As String

    Dim strChar              As String

    Set xlApp = CreateObject("excel.application")     '创建Excle对象

    xlApp.Visible = False

    Set xlBook = xlApp.Workbooks.Open(txtText1.Text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 9 Then
        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    SumCount = 0
    BCResultFlag = False

    For I = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count

        temp = ""
        source_batch_id_Temp = ""

        For J = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count

            If J > 26 Then
                strChar = Chr(96 + Int(J / 26 - 0.001)) & IIf(J Mod 26 = 0, "Z", Chr(96 + (J Mod 26)))
            Else
                strChar = Chr(96 + J)

            End If
        
            tempVal = xlSheet.Range(strChar & I).Value   '临时保存值

            If J = 1 Then
                bT.SHIPPING_INVOICE = Trim$(tempVal)
            ElseIf J = 2 Then
                bT.SHIPPING_DESTINATION = Trim$(tempVal)
            ElseIf J = 3 Then
                bT.GROSS_WEIGHT = Trim$(tempVal)
            ElseIf J = 4 Then
                bT.NET_WEIGHT = Trim(tempVal)
            ElseIf J = 5 Then
                bT.FORWAROER = Trim(tempVal)
            ElseIf J = 6 Then
                bT.AIR_WAYBILL = Trim(tempVal)
            ElseIf J = 7 Then
                bT.CARTON_QTY = Trim(tempVal)
            ElseIf J = 8 Then
                bT.BILLING_INVOICE = Trim$(tempVal)
            ElseIf J = 9 Then
                bT.FLAG = Trim$(tempVal)

            End If

        Next J
    
        Call AddB2b2(bT)
        SumCount = SumCount + 1

        '上传到DB
NextRecord2:

    Next I

    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set xlApp = Nothing

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"

    Else

        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"
            Exit Sub

        End If

    End If

End Sub

Private Sub cmdUpload_Click()

    Dim source_batch_id_Temp As String

    Dim dirName              As String

    Dim filename             As String

    Dim xlApp                As Excel.Application

    Dim xlBook               As Excel.Workbook

    Dim xlSheet              As Excel.Worksheet

    Dim bT                   As B2B

    Dim I                    As Integer

    Dim J                    As Integer

    Dim id                   As Long

    Dim temp                 As String

    Dim temp2                As String

    Dim tempVal              As String

    Dim strChar              As String

    If Combo1.Text = "" Then
        MsgBox "请选择客户类型", vbInformation, "友情提示:"
        Exit Sub

    End If

    If txtText1.Text = "" Then
        MsgBox "先选择待上传的文件", vbInformation, "友情提示:"
        Exit Sub

    End If

    If Combo1.Text = "US026发票" Then
        Call UpUS026FP
        Exit Sub

    End If

    Set xlApp = CreateObject("excel.application")     '创建Excle对象

    xlApp.Visible = False

    Set xlBook = xlApp.Workbooks.Open(txtText1.Text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 49 Then
        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If

    SumCount = 0
    BCResultFlag = False

    For I = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count

        temp = ""
        source_batch_id_Temp = ""

        For J = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count

            If J > 26 Then
                strChar = Chr(96 + Int(J / 26 - 0.001)) & IIf(J Mod 26 = 0, "Z", Chr(96 + (J Mod 26)))
            Else
                strChar = Chr(96 + J)

            End If
        
            tempVal = xlSheet.Range(strChar & I).Value   '临时保存值

            If J = 1 Then
                bT.EVENT_DATE = Trim$(tempVal)
            ElseIf J = 2 Then
                bT.EVENT = Trim$(tempVal)
            ElseIf J = 3 Then
                bT.OVT_COMPANY = Trim$(tempVal)
            ElseIf J = 4 Then
                bT.OVT_ORG = Trim(tempVal)
            ElseIf J = 5 Then
                bT.SUB_NAME = Trim(tempVal)
            ElseIf J = 6 Then
                bT.Stage = Trim(tempVal)
            ElseIf J = 7 Then
                bT.WAFER_LOT = Trim(tempVal)
            ElseIf J = 8 Then
                bT.OVT_JOB = Trim(tempVal)
            ElseIf J = 9 Then
                bT.SUB_LOT = Trim(tempVal)
            ElseIf J = 10 Then
                bT.FAB_Device = Trim(tempVal)
            ElseIf J = 11 Then
                bT.SOURCE_DEVICE = Trim(tempVal)
            ElseIf J = 12 Then
                bT.TARGET_DEVICE = Trim(tempVal)
            ElseIf J = 13 Then
                bT.WAFER_QTY = Trim(tempVal)
            ElseIf J = 14 Then
                bT.WAFER_DIE = Trim(tempVal)
            ElseIf J = 15 Then
                bT.WAFER_ID = Trim(tempVal)
            ElseIf J = 16 Then
                bT.PO = Trim(tempVal)
            ElseIf J = 17 Then
                bT.PO_RELEASE = Trim(tempVal)
            ElseIf J = 18 Then
                bT.OPERATION_CODE = Trim(tempVal)
            ElseIf J = 19 Then
                bT.OPERATION_DESCRIPTION = Trim(tempVal)
            ElseIf J = 20 Then
                bT.RECEIVE_QTY = Trim(tempVal)
            ElseIf J = 21 Then
                bT.JOB_ISSUE_QTY = Trim(tempVal)
            ElseIf J = 22 Then
                bT.START_QTY = Trim(tempVal)
            ElseIf J = 23 Then
                bT.COMPLETED_QTY = Trim(tempVal)
            ElseIf J = 24 Then
                bT.GRADE_RECORD = Trim(tempVal)
            ElseIf J = 25 Then
                bT.HOLD_CODE = Trim(tempVal)
            ElseIf J = 26 Then
                bT.HOLD_QTY = Trim(tempVal)
            ElseIf J = 27 Then
                bT.SCRAP_CODE = Trim(tempVal)
            ElseIf J = 28 Then
                bT.SCRAP_QTY = Trim(tempVal)
            ElseIf J = 29 Then
                bT.SCRAP_WAFER_ID = Trim(tempVal)
            ElseIf J = 30 Then
                bT.Priority = Trim(tempVal)
            ElseIf J = 31 Then
                bT.DATECODE = Trim(tempVal)
            ElseIf J = 32 Then
                bT.ENG_NO = Trim(tempVal)
            ElseIf J = 33 Then
                bT.NEXT_SUB_LOCATION = Trim(tempVal)
            ElseIf J = 34 Then
                bT.LOT_TYPE = Trim(tempVal)
            ElseIf J = 35 Then
                bT.E_SOD = Trim(tempVal)
            ElseIf J = 36 Then
                bT.TEST_PROGRAM = Trim(tempVal)
            ElseIf J = 37 Then
                bT.RMA_NO = Trim(tempVal)
            ElseIf J = 38 Then
                bT.BILLING_INVOICE = Trim(tempVal)
            ElseIf J = 39 Then
                bT.SHIPPING_INVOICE = Trim(tempVal)
            ElseIf J = 40 Then
                bT.SHIPPING_DESTINATION = Trim(tempVal)
            ElseIf J = 41 Then
                bT.JOB_FLAG = Trim(tempVal)
            ElseIf J = 42 Then
                bT.Remark = Trim(tempVal)
            ElseIf J = 43 Then
                bT.SO = Trim(tempVal)
            ElseIf J = 44 Then
                bT.SO_LINE = Trim(tempVal)
            ElseIf J = 45 Then
                bT.GROSS_WEIGHT = Trim(tempVal)
            ElseIf J = 46 Then
                bT.NET_WEIGHT = Trim(tempVal)
            ElseIf J = 47 Then
                bT.FORWARDER = Trim(tempVal)
            ElseIf J = 48 Then
                bT.AIR_WAYBILL = Trim(tempVal)
            ElseIf J = 49 Then
                bT.CARTON_QTY = Trim(tempVal)

            End If

        Next J
       
        ' 判断是否重复
        If Get_OracleStr("select * from B2B_ORDERTBL where wafer_lot = '" & bT.WAFER_LOT & "' and wafer_id = '" & bT.WAFER_ID & "'") <> "" Then
            MsgBox "LOT:" & bT.WAFER_LOT & " WAFER:" & bT.WAFER_ID & "已经存在, 即将更新此笔", vbInformation
            
            AddSql ("delete from B2B_ORDERTBL where wafer_lot = '" & bT.WAFER_LOT & "' and wafer_id = '" & bT.WAFER_ID & "'")
            AddSql2 ("delete from [ERPBASE].[dbo].[B2B_ORDERTBL] where wafer_lot = '" & bT.WAFER_LOT & "' and wafer_id = '" & bT.WAFER_ID & "'")

        End If

        ' 判断和WO是否一致
        Dim p
        Dim k As Integer
        Dim strWaferID As String
        
        p = Split(bT.WAFER_ID, "_")
        
        For k = 1 To UBound(p)
            strWaferID = p(k)
        Next
        
        Call AddB2b(bT)
        SumCount = SumCount + 1

        '上传到DB
NextRecord2:

    Next I

    xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set xlApp = Nothing

    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"

    Else

        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"
            Exit Sub

        End If

    End If

End Sub

Private Sub Form_Load()

    Combo1.AddItem ("US026发票")
    Combo1.AddItem ("US026工单模板")

End Sub
