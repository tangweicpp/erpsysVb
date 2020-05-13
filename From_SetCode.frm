VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form From_SetCode 
   Caption         =   "阴极线维护"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15060
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
   ScaleHeight     =   10590
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   10335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      Begin VB.TextBox txtSeq 
         Height          =   375
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         TabIndex        =   17
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "导出"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   6600
         TabIndex        =   16
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "删除"
         Height          =   360
         Index           =   3
         Left            =   5280
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "修改"
         Height          =   360
         Index           =   2
         Left            =   3840
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "增加"
         Height          =   360
         Index           =   1
         Left            =   2400
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "查找"
         Height          =   360
         Index           =   0
         Left            =   960
         TabIndex        =   12
         Top             =   1680
         Width           =   975
      End
      Begin MSDataListLib.DataCombo CmbCustomer 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtStatus 
         Height          =   285
         Left            =   8400
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   4320
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtBline 
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtDevice 
         Height          =   285
         Left            =   4320
         TabIndex        =   4
         Top             =   840
         Width           =   2415
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   7815
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   2400
         Width           =   12015
         _Version        =   524288
         _ExtentX        =   21193
         _ExtentY        =   13785
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
         SpreadDesigner  =   "From_SetCode.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7560
         TabIndex        =   9
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODE:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3720
         TabIndex        =   7
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BLINE:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "机种:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3720
         TabIndex        =   3
         Top             =   870
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   960
         TabIndex        =   1
         Top             =   870
         Width           =   945
      End
   End
End
Attribute VB_Name = "From_SetCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0

    E_SeqId = 1
    E_Customer
    e_DEVICE
    E_BLINE
    E_CODE
    E_STATUS
    E_End
    
End Enum

Private Sub cmd_Click(Index As Integer)

    Select Case Index

        Case 0  ' 查找
            ForSearch

        Case 1  ' 增加
            ForAdd

        Case 2  ' 修改
            ForMod

        Case 3  ' 删除
            ForDel

        Case 4  ' 退出
            ForExit

        Case 5  ' 导出
            ForExport

    End Select

End Sub

Private Sub ForSearch()

    Dim strCuscode As String

    Dim strDevice  As String

    Dim strSql     As String
    
    Dim rs         As New ADODB.Recordset

    If CmbCustomer.Text = "" Then
        MsgBox "请输入客户代码", vbInformation, "友情提示"
        Exit Sub

    End If

    strCuscode = Trim$(UCase$(CmbCustomer.Text))

    If txtDevice.Text = "" Then
        strSql = "select seq,customer,device, bline, code,status from code37 where customer = '" & strCuscode & "'"
    Else
        strDevice = Trim$(txtDevice.Text)
        strSql = "select seq,customer,device, bline, code,status from code37 where customer = '" & strCuscode & "' and device = '" & strDevice & "'"
    
    End If

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
  
    With fps(0)
        .MaxRows = 0

        If rs.RecordCount > 0 Then
            Set .DataSource = rs
       
        End If

    End With

End Sub

Private Sub ForAdd()
Dim strINfo As CODE37

If CmbCustomer.Text = "" Then
    MsgBox "请输入客户代码", vbInformation, "提示"
    Exit Sub
End If

If txtDevice.Text = "" Then
    MsgBox "请输入机种", vbInformation, "提示"
    Exit Sub
End If

If txtBline.Text = "" Then
    MsgBox "请输入BLINE", vbInformation, "提示"
    Exit Sub
End If

If txtCode.Text = "" Then
    MsgBox "请输入CODE", vbInformation, "提示"
    Exit Sub
End If

If txtStatus.Text = "" Then
    MsgBox "请输入STATUS", vbInformation, "提示"
    Exit Sub
End If

strINfo.strCus = Replace(UCase(Trim(CmbCustomer.Text)), Chr(13) + Chr(10), "")
strINfo.strDev = Replace(Trim(txtDevice.Text), Chr(13) + Chr(10), "")
strINfo.strBline = Replace(Trim(txtBline.Text), Chr(13) + Chr(10), "")
strINfo.strcode = Replace(Trim(txtCode.Text), Chr(13) + Chr(10), "")
strINfo.strStatus = Replace(Trim(txtStatus.Text), Chr(13) + Chr(10), "")

Call AddCode37(strINfo)

MsgBox "新增成功", vbInformation, "提示"

ShowDataAll

End Sub

Private Sub ForMod()
Dim strINfo As CODE37

strINfo.strBline = Replace(Trim(txtBline.Text), Chr(13) + Chr(10), "")
strINfo.strcode = Replace(Trim(txtCode.Text), Chr(13) + Chr(10), "")
strINfo.strStatus = Replace(Trim(txtStatus.Text), Chr(13) + Chr(10), "")

Call ModCode37(strINfo, CLng(txtSeq.Text))

MsgBox "修改成功", vbInformation, "提示"

ShowDataAll

End Sub

Private Sub ForDel()

    If CLng(txtSeq.Text) >= 1 Then

        Call DelCode37(CLng(txtSeq.Text))

        MsgBox "删除成功!", vbInformation, "友情提示"

        ShowDataAll

    Else

        MsgBox "请先双击要删除的行!", vbInformation, "友情提示"

    End If

End Sub

Private Sub ForExit()

    Unload Me

End Sub

Private Sub ForExport()

    Dim sqlTemp As String

    sqlTemp = "select * from code37"
         
    ExporToExcel (sqlTemp)

End Sub

Private Sub Form_Load()
    Init

End Sub

Private Sub Init()
    InitCuscode
    InitFps
    ShowDataAll

End Sub

Private Sub ShowDataAll()

    Dim rs     As New ADODB.Recordset

    Dim strSql As String

    strSql = "select seq,customer,device, bline, code,status from code37"

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
  
    With fps(0)
        .MaxRows = 0

        If rs.RecordCount > 0 Then
            Set .DataSource = rs
       
        End If

    End With

End Sub

Private Sub ShowData(i As Long)

    Dim rs     As New ADODB.Recordset

    Dim strSql As String

    strSql = "select seq,customer,device, bline, code,status from code37 where seq = '" & i & "'"

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
 
        CmbCustomer.Text = rs.Fields("customer").Value & ""
        txtDevice.Text = rs.Fields("device").Value & ""
        txtBline.Text = rs.Fields("bline").Value & ""
        txtCode.Text = rs.Fields("code").Value & ""
        txtStatus.Text = rs.Fields("status").Value & ""
        txtSeq.Text = rs.Fields("seq").Value & ""
         
    End If

End Sub

Private Sub InitFps()

    With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone

        .Col = -1
        .Row = -1
        .Lock = True
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .SetText E_FPS0.E_SeqId, 0, "记录号"
        .SetText E_FPS0.E_Customer, 0, "客户代码"
        .SetText E_FPS0.e_DEVICE, 0, "机种名"
        .SetText E_FPS0.E_BLINE, 0, "BLINE"
        .SetText E_FPS0.E_CODE, 0, "CODE"
        .SetText E_FPS0.E_STATUS, 0, "STATUS"
                
        .ColWidth(E_FPS0.E_SeqId) = 10
        .ColWidth(E_FPS0.E_Customer) = 10
        .ColWidth(E_FPS0.e_DEVICE) = 20
        .ColWidth(E_FPS0.E_BLINE) = 10
        .ColWidth(E_FPS0.E_CODE) = 10
        .ColWidth(E_FPS0.E_STATUS) = 8
       
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .ReDraw = True

    End With

End Sub

Private Sub InitCuscode()

    Dim mainItemRS As ADODB.Recordset

    Set mainItemRS = GetJDCustomerName()
    Set CmbCustomer.RowSource = mainItemRS
    CmbCustomer.ListField = mainItemRS("productname").Name
    CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub

Private Sub fps_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)

    Dim i As Long

    With fps(0)
        .Row = Row
        .Col = 1
        i = Trim(.Text)

    End With

    ShowData (i)

End Sub
