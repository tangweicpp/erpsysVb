VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FormWLP_DN 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "DN_UP"
   ClientHeight    =   11505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "WL"
   MDIChild        =   -1  'True
   ScaleHeight     =   11505
   ScaleWidth      =   15675
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTTab0 
      Height          =   11415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   20135
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "DN_UP"
      TabPicture(0)   =   "FormWLP_DN.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtPath"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblcust"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLOT"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbldevice"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblstart"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblend"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDN"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CommonDialog1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Fps(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdup"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdquery"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtcust"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtlot"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtdevice"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtdn"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtstart"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtend"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "DTPicker1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "DTPicker2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdCommand1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FormWLP_DN.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdCommand1 
         Caption         =   "导出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   9960
         TabIndex        =   20
         Top             =   2400
         Visible         =   0   'False
         Width           =   990
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   9120
         TabIndex        =   19
         Top             =   2520
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Format          =   261423105
         CurrentDate     =   43719
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   9120
         TabIndex        =   18
         Top             =   1920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Format          =   127467521
         CurrentDate     =   43719
         MinDate         =   43101
      End
      Begin VB.TextBox txtend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   17
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox txtstart 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   16
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox txtdn 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   15
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtdevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtlot 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   13
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtcust 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1680
         TabIndex        =   12
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton cmdquery 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   990
      End
      Begin VB.CommandButton cmdup 
         Caption         =   "上传"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   990
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   7815
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   3480
         Width           =   15615
         _Version        =   524288
         _ExtentX        =   27543
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
         MaxCols         =   3
         MaxRows         =   0
         SpreadDesigner  =   "FormWLP_DN.frx":0038
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4080
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   11
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label lblend 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4560
         TabIndex        =   10
         Top             =   2520
         Width           =   1305
      End
      Begin VB.Label lblstart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4560
         TabIndex        =   9
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label lbldevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户机种:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   1305
      End
      Begin VB.Label lblLOT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOT:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   840
         TabIndex        =   7
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label lblcust 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件路径"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   5040
         TabIndex        =   4
         Top             =   840
         Width           =   900
      End
      Begin MSForms.TextBox txtPath 
         Height          =   315
         Left            =   6120
         TabIndex        =   3
         Top             =   840
         Width           =   5655
         VariousPropertyBits=   746604563
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "9975;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
   End
End
Attribute VB_Name = "FormWLP_DN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdup_Click()

If Replace(txtcust.Text, Chr(13) + Chr(10), "") = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub
End If

'If chkMsgAppend.Value = 1 And Trim(txtMsg.Text) = "" Then
'    MsgBox "您已勾选了邮件补充, 请填写正文 补充内容" & vbCrLf & "否则请取消勾选再上传", vbInformation, "提示"
'    Exit Sub
'End If

CommonDialog1.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
CommonDialog1.ShowOpen
If CommonDialog1.filename = "" Then
    Exit Sub

End If

txtPath.Text = CommonDialog1.filename
CommonDialog1.filename = ""
If txtPath.Text = "" Then
    MsgBox "请选择要上传的文件", vbInformation, "提示"
    Exit Sub

End If

Call Upload_0


End Sub

Private Sub Upload_0()
  On Error GoTo ErrHandle

    Dim VBExcel     As Excel.Application

    Dim xlBook      As Excel.Workbook

    Dim xlSheet     As Excel.Worksheet
    
    Dim CUST_ID  As String

    Dim CUST_DEVICE As String

    Dim CUST_LOT  As String
    
    Dim DATE_CODE  As String
    
    Dim QTY  As String
    
    Dim SHIP_PO  As String
    
    Dim SHIP_TO  As String
    
    Dim DN_NUM  As String
    
    Dim User As String
    
    Dim rs         As New ADODB.Recordset

    Dim strSql     As String
    
    Dim i As Integer
    
    
    User = gUserName
    
    strSql = ""
    
        
    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(txtPath.Text)

    Set xlSheet = xlBook.Worksheets(1)
 
  
    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 7 Then
        
        MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        GoTo EXITPRO
        Exit Sub

    End If
    
    
    DN_NUM = "DN" + Format(DATE, "yyyy") & Format(DATE, "mm") & Format(DATE, "dd") + Format(Get_OracleStr("select DN_SEQ.NEXTVAL FROM DUAL"), "0000")
    
    Fps(0).MaxRows = 0
    
    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        CUST_ID = Replace(Trim(xlSheet.Range("A" & i)), Chr(13) + Chr(10), "")
        CUST_DEVICE = Replace(Trim(xlSheet.Range("B" & i)), Chr(13) + Chr(10), "")
        CUST_LOT = Replace(Trim(xlSheet.Range("C" & i)), Chr(13) + Chr(10), "")
        DATE_CODE = Replace(Trim(xlSheet.Range("D" & i)), Chr(13) + Chr(10), "")
        QTY = Replace(Trim(xlSheet.Range("E" & i)), Chr(13) + Chr(10), "")
        SHIP_PO = Replace(Trim(xlSheet.Range("F" & i)), Chr(13) + Chr(10), "")
        SHIP_TO = Replace(Trim(xlSheet.Range("G" & i)), Chr(13) + Chr(10), "")
        
        strSql = "INSERT INTO erptemp..ht_dn(DN_NUM,CUST_ID,CUST_DEVICE,CUST_LOT,DC,QTY,SHIP_PO,SHIP_AD,create_by,create_date) " & _
                " VALUES ('" & DN_NUM & "','" & CUST_ID & "','" & CUST_DEVICE & "','" & CUST_LOT & "','" & DATE_CODE & "','" & QTY & "','" & SHIP_PO & "','" & SHIP_TO & "','" & User & "',GETDATE())"
        
        AddSql2 (strSql)
       

    Next
    
  
    
    Query (DN_NUM)
   
EXITUPLOAD:

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
   
    Exit Sub
EXITPRO:

    On Error Resume Next

    MousePointer = 0
    
     MsgBox CUST_DEVICE & "$" & CUST_LOT & "上传失败", vbInformation, "提示"

    If Not VBExcel Is Nothing Then

        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing

    End If
    Exit Sub
ErrHandle:
    GoTo EXITPRO

End Sub


Private Sub Query(dn As String)


    Dim rs         As New ADODB.Recordset

    Dim strSql     As String


     If Replace(txtcust.Text, Chr(13) + Chr(10), "") = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub
      End If


   strSql = " SELECT a.dn_num,a.ship_po,a.cust_device,a.cust_lot,a.dc,a.qty,ISNULL(SUM(b.数量),0) AS 库存DIE " & _
            "    ,ISNULL(COUNT(DISTINCT c.箱号 ),0) AS 库存卷,RTRIM(ISNULL(c.仓位,'')) AS 仓位 FROM erptemp..ht_dn a " & _
            "     LEFT JOIN erpdata..tblStockNumSub b ON a.CUST_LOT = b.工单号 AND b.库房编号 = '99' LEFT JOIN erpdata..tblStockNumTree c ON c.箱号 = b.箱号 " & _
            "     WHERE a.dn_num = '" & dn & "'  GROUP BY a.dn_num,a.cust_device,a.cust_lot,a.dc,a.qty,a.ship_po,c.仓位     "


    Fps(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType(rs)
     MsgBox "上传完成", vbInformation, "提示"
    Else

        MsgBox "上传失败", vbInformation, "提示"
        Exit Sub

    End If

End Sub

Private Sub cmdQuery_Click()

If Replace(txtcust.Text, Chr(13) + Chr(10), "") = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub
End If

Query1

cmdCommand1.Visible = True


End Sub


Private Sub ListDataType(rs As ADODB.Recordset)

    Dim i As Long

    With Fps(0)

        .MaxRows = 0

        Set .DataSource = rs

    End With

    With Fps(0)

        For i = 1 To .MaxRows

        Next

    End With

End Sub







Private Sub Form_Load()
With Fps(0)
    .Col = -1
    .Row = -1
    .Lock = True

End With
End Sub


Private Sub DTPicker1_CHANGE()

txtstart.Text = Format(Trim(DTPicker1.Value), "YYYY-MM-DD")

End Sub




Private Sub DTPicker2_CHANGE()
txtend.Text = Format(Trim(DTPicker2.Value), "YYYY-MM-DD")

End Sub



Private Sub Query1()


    Dim rs         As New ADODB.Recordset

    Dim strSql     As String

        strSql = " SELECT a.dn_num,a.ship_po,a.cust_device,a.cust_lot,a.dc,a.qty,ISNULL(SUM(b.数量),0) AS 库存DIE " & _
                "    ,ISNULL(COUNT(DISTINCT c.箱号 ),0) AS 库存卷,RTRIM(ISNULL(c.仓位,'')) AS 仓位 FROM erptemp..ht_dn a " & _
                "     LEFT JOIN erpdata..tblStockNumSub b ON a.CUST_LOT = b.工单号 AND b.库房编号 = '99' LEFT JOIN erpdata..tblStockNumTree c ON c.箱号 = b.箱号 " & _
                "     WHERE  a.cust_id = '" & Replace(txtcust.Text, Chr(13) + Chr(10), "") & "'    "
                
                 
    If Replace(txtlot.Text, Chr(13) + Chr(10), "") <> "" Then
    
    strSql = strSql + "and  a.cust_lot =  '" & Replace(txtlot.Text, Chr(13) + Chr(10), "") & "'  "
    
    End If
    
   If Replace(txtdevice.Text, Chr(13) + Chr(10), "") <> "" Then
    
    strSql = strSql + "and  a.cust_device =  '" & Replace(txtdevice.Text, Chr(13) + Chr(10), "") & "'  "
    
    End If
    
     If Replace(txtdn.Text, Chr(13) + Chr(10), "") <> "" Then
    
    strSql = strSql + "and  a.dn_num =  '" & Replace(txtdn.Text, Chr(13) + Chr(10), "") & "'  "
    
    End If
    
     If Replace(txtstart.Text, Chr(13) + Chr(10), "") <> "" Then
    
    strSql = strSql + "and  a.cust_lot >=  '" & Replace(txtstart.Text, Chr(13) + Chr(10), "") & "'  "
    
    End If
    
     If Replace(txtstart.Text, Chr(13) + Chr(10), "") <> "" Then
    
    strSql = strSql + "and  a.cust_lot <=  '" & Replace(txtstart.Text, Chr(13) + Chr(10), "") & "'  "
    
    End If
    
    strSql = strSql + "  GROUP BY a.dn_num,a.cust_device,a.cust_lot,a.dc,a.qty,a.ship_po,c.仓位  "
   

    Fps(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType(rs)
     MsgBox "查询完成", vbInformation, "提示"
    Else

        MsgBox "查询失败", vbInformation, "提示"
        Exit Sub

    End If

End Sub

Private Sub cmdCommand1_Click()

 Dim strSql     As String

        strSql = " SELECT a.dn_num,a.ship_po,a.cust_device,a.cust_lot,a.dc,a.qty,ISNULL(SUM(b.数量),0) AS 库存DIE " & _
                "    ,ISNULL(COUNT(DISTINCT c.箱号 ),0) AS 库存卷,RTRIM(ISNULL(c.仓位,'')) AS 仓位 FROM erptemp..ht_dn a " & _
                "     LEFT JOIN erpdata..tblStockNumSub b ON a.CUST_LOT = b.工单号 AND b.库房编号 = '99' LEFT JOIN erpdata..tblStockNumTree c ON c.箱号 = b.箱号 " & _
                "     WHERE  a.cust_id = '" & Replace(txtcust.Text, Chr(13) + Chr(10), "") & "'      "
                 
    If Replace(txtlot.Text, Chr(13) + Chr(10), "") <> "" Then
    
    strSql = strSql + "and  a.cust_lot =  '" & Replace(txtlot.Text, Chr(13) + Chr(10), "") & "'  "
    
    End If
    
   If Replace(txtdevice.Text, Chr(13) + Chr(10), "") <> "" Then
    
    strSql = strSql + "and  a.cust_device =  '" & Replace(txtdevice.Text, Chr(13) + Chr(10), "") & "'  "
    
    End If
    
     If Replace(txtdn.Text, Chr(13) + Chr(10), "") <> "" Then
    
    strSql = strSql + "and  a.dn_num =  '" & Replace(txtdn.Text, Chr(13) + Chr(10), "") & "'  "
    
    End If
    
     If Replace(txtstart.Text, Chr(13) + Chr(10), "") <> "" Then
    
    strSql = strSql + "and  a.cust_lot >=  '" & Replace(txtstart.Text, Chr(13) + Chr(10), "") & "'  "
    
    End If
    
     If Replace(txtstart.Text, Chr(13) + Chr(10), "") <> "" Then
    
    strSql = strSql + "and  a.cust_lot <=  '" & Replace(txtstart.Text, Chr(13) + Chr(10), "") & "'  "
    
    End If
    
    strSql = strSql + "  GROUP BY a.dn_num,a.cust_device,a.cust_lot,a.dc,a.qty,a.ship_po,c.仓位  "
   
    SqlServerExporToExcel ("" & strSql & "")
    
    
End Sub










