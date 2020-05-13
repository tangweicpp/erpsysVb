VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmGet 
   Caption         =   "工单入库信息"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   22725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   22725
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin FPSpreadADO.fpSpread fpSpread2 
      Height          =   7695
      Left            =   11400
      TabIndex        =   13
      Top             =   1800
      Width           =   11175
      _Version        =   524288
      _ExtentX        =   19711
      _ExtentY        =   13573
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
      SpreadDesigner  =   "FrmGetAll.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "工单明细查询"
      Height          =   360
      Left            =   9600
      TabIndex        =   12
      Top             =   840
      Width           =   1590
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   7695
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   10935
      _Version        =   524288
      _ExtentX        =   19288
      _ExtentY        =   13573
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
      SpreadDesigner  =   "FrmGetAll.frx":03EA
      AppearanceStyle =   0
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   360
      Left            =   10200
      TabIndex        =   6
      Top             =   240
      Width           =   990
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   195952641
      CurrentDate     =   43620
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   195952641
      CurrentDate     =   43620
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "支持双击工单查询"
      Height          =   195
      Left            =   9720
      TabIndex        =   14
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工单号"
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "成品料号"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工单结束时间"
      Height          =   195
      Left            =   6840
      TabIndex        =   2
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工单开始时间"
      Height          =   195
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户代码"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "FrmGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    '客户代码
    Dim kh As String
    '工单开始日期
    Dim beginTime As String
    '工单结束日期
    Dim endTime As String
    '成品料号
    Dim cl As String
    '工单号
    Dim gd As String
    
    Dim strSql1 As String
    
    Dim strSql2 As String
    
     Dim strSql3 As String
    
    Dim rs As New ADODB.Recordset
    
    kh = UCase$(Trim$(Text1.Text))
    beginTime = Format(DTPicker1.Value, "YYYY/MM/DD")
    endTime = Format(DTPicker2.Value, "YYYY/MM/DD")
    cl = UCase$(Trim$(Text2.Text))
    gd = UCase$(Trim$(Text3.Text))
    strSql2 = " and b.PRODUCT='" & cl & "'"
    strSql3 = " where a.大工单='" & gd & "'"
    
    If kh = "" Then
        MsgBox "请输入客户代码", vbInformation, "提示"
        Exit Sub
    ElseIf DateDiff("m", beginTime, endTime) > 6 Then
         MsgBox "间隔时间不能超过半年", vbInformation, "提示"
         Exit Sub
    ElseIf beginTime >= DATE Or endTime > DATE Then
        MsgBox "日期输入错误，请重新输入", vbInformation, "提示"
        Exit Sub
    Else
        strSql1 = "SELECT x.*,ISNULL(y.入库数,0) AS 入库数,x.工单数量 - ISNULL(y.入库数,0) AS 差异,'' as '√' FROM (select  b.ORDERNAME,SUM(CONVERT(INT, c.DIEQTY)) AS 工单数量,CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) AS 工单时间 FROM erpdata .. tblTSVworkorder b  LEFT JOIN  erpdata .. tblTSVwaferlist c ON c.ORDERNAME = b.ORDERNAME WHERE CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) > '" & beginTime & "' and CONVERT(VARCHAR(100), b.ERPCREATEDATE,23)<'" & endTime & "' and b.CUSTOMER ='" & kh & "' GROUP BY b.ORDERNAME ,CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) )x LEFT JOIN  (SELECT a.大工单,SUM(a.入库数)  AS 入库数 FROM erpdata..tblPackToHouseRec a  GROUP BY a.大工单 ) y ON y.大工单 = x.ORDERNAME ORDER BY x.工单时间"
        If cl <> "" Then
            strSql1 = "SELECT x.*,ISNULL(y.入库数,0) AS 入库数,x.工单数量 - ISNULL(y.入库数,0) AS 差异,'' as '√' FROM (select  b.ORDERNAME,SUM(CONVERT(INT, c.DIEQTY)) AS 工单数量,CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) AS 工单时间 FROM erpdata .. tblTSVworkorder b  LEFT JOIN  erpdata .. tblTSVwaferlist c ON c.ORDERNAME = b.ORDERNAME WHERE CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) > '" & beginTime & "' and CONVERT(VARCHAR(100), b.ERPCREATEDATE,23)<'" & endTime & "' and b.CUSTOMER ='" & kh & "' " & strSql2 & " GROUP BY b.ORDERNAME ,CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) )x LEFT JOIN  (SELECT a.大工单,SUM(a.入库数)  AS 入库数 FROM erpdata..tblPackToHouseRec a  GROUP BY a.大工单 ) y ON y.大工单 = x.ORDERNAME ORDER BY x.工单时间"
           
            If gd <> "" Then
                 strSql1 = "SELECT x.*,ISNULL(y.入库数,0) AS 入库数,x.工单数量 - ISNULL(y.入库数,0) AS 差异,'' as '√' FROM (select  b.ORDERNAME,SUM(CONVERT(INT, c.DIEQTY)) AS 工单数量,CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) AS 工单时间 FROM erpdata .. tblTSVworkorder b  LEFT JOIN  erpdata .. tblTSVwaferlist c ON c.ORDERNAME = b.ORDERNAME WHERE CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) > '" & beginTime & "' and CONVERT(VARCHAR(100), b.ERPCREATEDATE,23)<'" & endTime & "' and b.CUSTOMER ='" & kh & "' " & strSql2 & "  GROUP BY b.ORDERNAME ,CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) )x LEFT JOIN  (SELECT a.大工单,SUM(a.入库数)  AS 入库数 FROM erpdata..tblPackToHouseRec a  " & strSql3 & " GROUP BY a.大工单 ) y ON y.大工单 = x.ORDERNAME ORDER BY x.工单时间"
            End If
            
        ElseIf gd <> "" Then
            strSql1 = "SELECT x.*,ISNULL(y.入库数,0) AS 入库数,x.工单数量 - ISNULL(y.入库数,0) AS 差异,'' as '√' FROM (select  b.ORDERNAME,SUM(CONVERT(INT, c.DIEQTY)) AS 工单数量,CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) AS 工单时间 FROM erpdata .. tblTSVworkorder b  LEFT JOIN  erpdata .. tblTSVwaferlist c ON c.ORDERNAME = b.ORDERNAME WHERE CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) > '" & beginTime & "' and CONVERT(VARCHAR(100), b.ERPCREATEDATE,23)<'" & endTime & "' and b.CUSTOMER ='" & kh & "'  GROUP BY b.ORDERNAME ,CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) )x LEFT JOIN  (SELECT a.大工单,SUM(a.入库数)  AS 入库数 FROM erpdata..tblPackToHouseRec a  " & strSql3 & " GROUP BY a.大工单 ) y ON y.大工单 = x.ORDERNAME ORDER BY x.工单时间"
               
            ''If cl <> "" Then
                 ''strSql1 = "SELECT x.*,ISNULL(y.入库数,0) AS 入库数,x.工单数量 - ISNULL(y.入库数,0) AS 差异 FROM (select  b.ORDERNAME,SUM(CONVERT(INT, c.DIEQTY)) AS 工单数量,CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) AS 工单时间 FROM erpdata .. tblTSVworkorder b  LEFT JOIN  erpdata .. tblTSVwaferlist c ON c.ORDERNAME = b.ORDERNAME WHERE CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) > '" & beginTime & "' and CONVERT(VARCHAR(100), b.ERPCREATEDATE,23)<'" & endTime & "' and b.CUSTOMER ='" & kh & "' " & strSql3 & " " & strSql2 & " GROUP BY b.ORDERNAME ,CONVERT(VARCHAR(100), b.ERPCREATEDATE,23) )x LEFT JOIN  (SELECT a.大工单,SUM(a.入库数)  AS 入库数 FROM erpdata..tblPackToHouseRec a  GROUP BY a.大工单 ) y ON y.大工单 = x.ORDERNAME ORDER BY x.工单时间"
                 ''MsgBox strSql1
            ''End If
        End If
    End If
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql1, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
     If Not (rs.EOF And rs.BOF) Then '表示有数据了
        Call ListDataType(rs)
    Else
        MsgBox "查询不到信息", vbInformation, "提示"
        Exit Sub

    End If
End Sub

Private Sub ListDataType(rs As ADODB.Recordset)
    Dim I As Long
   
   With fpSpread1(1)
        .MaxRows = 0

        Set .DataSource = rs

    End With
   
    
    Rem 标示不同的差异颜色
    With fpSpread1(1)

        For I = 1 To .MaxRows
            .Row = I
            .Col = 5
            If .Text > 0 Then
                .BackColor = &HFFFF&
            ElseIf .Text < 0 Then
                .BackColor = &HFF&
            End If
            
        Next

    End With
 
    Rem 加上勾选框
    With fpSpread1(1)

        For I = 1 To .MaxRows
            .Row = I
            .Col = 6
            .CellType = CellTypeCheckBox
            
        Next

    End With
 
End Sub

    
Private Sub Command2_Click()
    
    Dim I As Long
    
    Dim J As Long
    
    Dim gd As String
    
    Dim flag As Boolean
    
    Dim rs As New ADODB.Recordset
    
    Dim strSql As String
    
    flag = False
    
    With fpSpread1(1)

        For I = 1 To .MaxRows
            .Row = I
            .Col = 6
            If .Text = "1" Then
                flag = True

                .Col = 1
               gd = .Text
               
               strSql = "SELECT a.*  FROM erpdata..tblPackToHouseRec a  WHERE a.大工单 = '" & gd & "'"
               
               
               If rs.State = adStateOpen Then rs.Close
               rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
               
               
               If Not (rs.BOF And rs.EOF) Then
                 With fpSpread2
        
                .MaxRows = 0

                Set .DataSource = rs
                End With
                
                With fpSpread2

                For J = 1 To rs.RecordCount
                  .Row = J
                  .Col = 9
                  .CellType = CellTypeCheckBox
                 Next

                End With
                Else

                MsgBox "没有该工单的详细信息"
                Exit Sub
                End If
                
            End If
        Next

    End With
    
    If flag = False Then
         MsgBox "请勾选需要查询工单号"
         Exit Sub
    End If

End Sub

Private Sub Form_Load()

    DTPicker1.Value = Now() - 1
    DTPicker2.Value = Now()
End Sub

Private Sub ShowData(gd As String)
       Dim rs As New ADODB.Recordset
       
       Dim strSql As String
       
       Dim I As Long
       
       strSql = "SELECT a.*  FROM erpdata..tblPackToHouseRec a  WHERE a.大工单 = '" & gd & "'"
       
       If rs.State = adStateOpen Then rs.Close
       
       rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
       
       If rs.RecordCount > 0 Then
       
            With fpSpread2
        
                .MaxRows = 0
                Set .DataSource = rs

            End With
            With fpSpread2
                For I = 1 To rs.RecordCount
                    .Row = I
                    .Col = 9
                       
                Next
            End With
        Else
           MsgBox "没有该工单的详细信息"
           Exit Sub
        End If
             
End Sub

Private Sub fpSpread1_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)

    Dim I As String
    
    Dim rso As ADODB.Recordset
    
    
    
      
    
    With fpSpread1(1)
        .Row = Row
        .Col = 1
        If .Row <> 0 Then
            I = .Text
        End If
    End With

   ShowData (I)
    
   

End Sub

