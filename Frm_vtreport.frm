VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_vtreport 
   Caption         =   "委外报表"
   ClientHeight    =   9915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20070
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   20070
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "仅辅助"
      Height          =   495
      Left            =   15960
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Cob_cust 
      Height          =   300
      ItemData        =   "Frm_vtreport.frx":0000
      Left            =   1080
      List            =   "Frm_vtreport.frx":000D
      TabIndex        =   13
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "仅辅助"
      Height          =   495
      Left            =   14520
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "查询"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.CheckBox Chk_bonded 
      Caption         =   "仅保税"
      Height          =   255
      Left            =   11280
      TabIndex        =   10
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "委外未完全回货"
      Height          =   375
      Left            =   9360
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "退出"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton CmdOutput 
      Caption         =   "导出"
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   106823681
      CurrentDate     =   43822
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   106823681
      CurrentDate     =   43822
   End
   Begin FPSpreadADO.fpSpread fpS 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   19575
      _Version        =   524288
      _ExtentX        =   34528
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
      SpreadDesigner  =   "Frm_vtreport.frx":0020
   End
   Begin VB.Label Label5 
      Caption         =   "合计"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   9360
      Width           =   6135
   End
   Begin VB.Label Label4 
      Caption         =   "客户代码"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   735
   End
   Begin MSForms.ComboBox CobPn 
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3625;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "料号"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "结束时间"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "开始时间"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Frm_vtreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOutput_Click()
  CmdOutput.Caption = "导出中..."
  CmdOutput.Enabled = False
  
  FpsToExcel
  CmdOutput.Caption = "导出"
  CmdOutput.Enabled = True
End Sub







Private Sub Command2_Click()
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer

    CobPn.Clear
    
    If SMR.State = adStateOpen Then SMR.Close

 
strSql = " SELECT distinct rtrim(t3.料号) AS 料号 ,rtrim(t1.流程卡编号) AS 流程卡编号 ,t1.申请时间 AS 第一次发出时间,t1.合格数 AS 发出数量,t2.申请时间 AS 最后一次回货时间,t2.合格数 AS 回货数量 FROM  " & _
" (SELECT b.流程卡编号,sum(b.合格数+b.制程不良数+b.来料不良数) AS 合格数 ,min(a.申请时间) AS 申请时间  " & _
" FROM erpdata..tblstockdb  a  " & _
" LEFT JOIN erpdata..tblstockdbsub b ON a.调拨编号=b.调拨编号 AND a.序号=b.序号  " & _
" INNER JOIN erpbase..tblstock e ON a.原仓库=e.库房代码 " & _
" WHERE a.原仓库<>'72' AND a.目标仓库='72'  "
If Chk_bonded.Value = 1 Then
    strSql = strSql & " and  e.库房类型<>'非保税' "
End If

strSql = strSql & " GROUP BY b.流程卡编号 ) t1  " & _
" LEFT JOIN  " & _
" (SELECT d.流程卡编号,sum(d.合格数+d.制程不良数+d.来料不良数)  AS 合格数,max(c.申请时间) AS 申请时间  " & _
" FROM erpdata..tblstockdb  c   " & _
" LEFT JOIN erpdata..tblstockdbsub d ON c.调拨编号=d.调拨编号 AND c.序号=d.序号   " & _
" INNER JOIN erpbase..tblstock f ON c.目标仓库=f.库房代码 " & _
" WHERE c.原仓库='72' AND c.目标仓库<>'72'"
If Chk_bonded.Value = 1 Then
    strSql = strSql & " and  f.库房类型<>'非保税' "
End If

strSql = strSql & "  GROUP BY d.流程卡编号  ) t2 ON t1.流程卡编号=t2.流程卡编号  " & _
" LEFT JOIN erpdata ..tblPackMainInfSub t3   ON t3.流程卡编号=t1.流程卡编号  " & _
" WHERE isnull(t1.合格数,0)>isnull(t2.合格数,0)   " & _
" ORDER BY rtrim(t3.料号),t1.申请时间"


    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If SMR.RecordCount > 0 Then
        With Fps
           .MaxRows = 0
           Set .DataSource = SMR
           
       End With

    Else
   
      With Fps
           .MaxRows = 0
          
        End With

   End If

    MsgBox "共查询出" & Fps.MaxRows & "条记录", vbInformation, "提示"
    SMR.Close
    
    Set SMR = Nothing
End Sub

Private Sub Command3_Click()
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer

    Dim pnlist     As String
    Command3.Enabled = False
    If SMR.State = adStateOpen Then SMR.Close
    
    pnlist = ""
   getdatafromstockdb

 ' If Trim(CobPn.Text) = "" Then
    ' strSql = "SELECT distinct c.料号 FROM erpdata..tblstockdbsub a " & _
    ' " LEFT JOIN erpdata..tblstockdb b ON a.调拨编号=b.调拨编号 AND a.序号=b.序号 " & _
    ' " LEFT JOIN erpdata ..tblPackMainInfSub c   ON a.流程卡编号=c.流程卡编号 " & _
    ' " INNER JOIN erpbase..tblstock e ON b.目标仓库=e.库房代码 " & _
    ' " INNER JOIN erpbase..tblstock f ON b.原仓库=f.库房代码 " & _
    ' " WHERE  e.库房类型<>'非保税' AND f.库房类型<>'非保税' " & _
    ' " AND  c.料号 IS NOT NULL AND  (b.目标仓库='72'  OR b.原仓库='72') " & _
    ' " And b.申请时间<'" & Format(DTP2.Value + 1, "yyyy/mm/dd") & "' and  b.申请时间>'" & Format(DTP1.Value, "yyyy/mm/dd") & "'"
        
        ' SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        ' If SMR.RecordCount > 0 Then
            ' SMR.MoveFirst
            ' For i = 1 To SMR.RecordCount
            ' If pnlist = "" Then
                ' pnlist = "'" & Trim(SMR("料号")) & "'"
            ' Else
                ' pnlist = pnlist & "," & "'" & Trim(SMR("料号")) & "'"
            ' End If
                ' SMR.MoveNext
            ' Next
        ' End If
        ' SMR.Close
        
        ' Set SMR = Nothing
    ' Else
        ' pnlist = "'" & Trim(CobPn.Text) & "'"
    ' End If
    date_start = Format(DTP1.Value, "yyyy/mm/dd")
    date_end = Format(DTP2.Value, "yyyy/mm/dd")
    date_MID = "2020/01/01 00:00:00"
    If DateDiff("D", date_end, date_MID) > 0 Then
    '只看2020/1/1之前的数据
        strSql = "  select 客户机种 , 料号, qty外发, qty回货, 金额, 申请时间 , 申请人员, 工单号, rtrim(流程卡编号) AS 流程卡编号,委托方企业名称, 调拨编号, 序号, 原仓库, 目标仓库 FROM  erpdata..zh_tblVT_modify   where  (flag=1 or flag=2 or flag=9) and 申请时间<'" & Format(DTP2.Value + 1, "yyyy/mm/dd") & "' and  申请时间>'" & Format(DTP1.Value, "yyyy/mm/dd") & "'"
        
        If UCase(Trim(Cob_cust.text)) = "GC" Then
            strSql = strSql & " and 客户代码='GC'"
       ElseIf UCase(Trim(Cob_cust.text)) = "GD108" Then
            strSql = strSql & " and  料号 Like '%GD108%'"
        ElseIf UCase(Trim(Cob_cust.text)) = "KR" Then
            strSql = strSql & " and  料号 Like '%KR%'"
        End If
        
        If Trim(CobPn.text) <> "" Then
            strSql = strSql & " AND 料号 ='" & Trim(CobPn.text) & "'"
        End If
        
        strSql = strSql & "  ORDER BY 料号,rtrim(流程卡编号) ,申请时间 "
            
        
        
        
    ElseIf DateDiff("D", date_start, date_MID) > 0 Then
    '跨越2020/1/1
         
         
        strSql = " select 客户机种 , 料号, qty外发, qty回货, 金额, 申请时间 , 申请人员, 工单号, rtrim(流程卡编号) AS 流程卡编号,委托方企业名称, 调拨编号, 序号, 原仓库, 目标仓库 from  erpdata..zh_tblVT_new where flag=1  and 申请时间>='" & date_MID & "' and  申请时间<'" & Format(DTP2.Value, "yyyy/mm/dd") & "'"
        If UCase(Trim(Cob_cust.text)) = "GC" Then
            strSql = strSql & " and 客户代码='GC'"
        ElseIf UCase(Trim(Cob_cust.text)) = "GD108" Then
            strSql = strSql & " and  料号 Like '%GD108%'"
        ElseIf UCase(Trim(Cob_cust.text)) = "KR" Then
            strSql = strSql & " and  料号 Like '%KR%'"
        End If
        If Trim(CobPn.text) <> "" Then
            strSql = strSql & " AND 料号 ='" & Trim(CobPn.text) & "'"
        End If
        
    
         
         strSql = strSql & " union select 客户机种 , 料号, qty外发, qty回货, 金额, 申请时间 , 申请人员, 工单号, rtrim(流程卡编号) AS 流程卡编号,委托方企业名称, 调拨编号, 序号, 原仓库, 目标仓库 FROM  erpdata..zh_tblVT_modify   where  (flag=1 or flag=2 or flag=9) and 申请时间<'" & date_MID & "' and  申请时间>'" & Format(DTP1.Value, "yyyy/mm/dd") & "'"
        
        If UCase(Trim(Cob_cust.text)) = "GC" Then
            strSql = strSql & " and 客户代码='GC'"
        ElseIf UCase(Trim(Cob_cust.text)) = "GD108" Then
            strSql = strSql & " and  料号 Like '%GD108%'"
        ElseIf UCase(Trim(Cob_cust.text)) = "KR" Then
            strSql = strSql & " and  料号 Like '%KR%'"
        End If
        If Trim(CobPn.text) <> "" Then
            strSql = strSql & " AND 料号 ='" & Trim(CobPn.text) & "'"
        End If
  


        strSql = strSql & "   ORDER BY 料号,rtrim(流程卡编号) ,申请时间 "
   
    
    Else
    '只看2020/1/1之后的数据
    
         
        strSql = " select 客户机种 , 料号, qty外发, qty回货, 金额, 申请时间 , 申请人员, 工单号, rtrim(流程卡编号) AS 流程卡编号,委托方企业名称, 调拨编号, 序号, 原仓库, 目标仓库  from  erpdata..zh_tblVT_new where flag=1  and 申请时间>='" & date_start & "' and  申请时间<'" & Format(DTP2.Value, "yyyy/mm/dd") & "'"
        If UCase(Trim(Cob_cust.text)) = "GC" Then
            strSql = strSql & " and 客户代码='GC'"
        ElseIf UCase(Trim(Cob_cust.text)) = "GD108" Then
            strSql = strSql & " and  料号 Like '%GD108%'"
        ElseIf UCase(Trim(Cob_cust.text)) = "KR" Then
            strSql = strSql & " and  料号 Like '%KR%'"
        End If
        If Trim(CobPn.text) <> "" Then
            strSql = strSql & " AND 料号 ='" & Trim(CobPn.text) & "'"
        End If
        
        
        strSql = strSql & "  ORDER BY 料号,rtrim(流程卡编号) ,申请时间 "
        
     
    End If
    
 

  If SMR.State = adStateOpen Then SMR.Close
  SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  If SMR.RecordCount > 0 Then
        With Fps
           .MaxRows = 0
           Set .DataSource = SMR
           
       End With

    Else
   
      With Fps
           .MaxRows = 0
'
          
        End With

   End If

 MsgBox "共查询出" & Fps.MaxRows & "条记录", vbInformation, "提示"
   
    sumQty
    Command3.Enabled = True

End Sub

Private Sub Command4_Click()
    
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    
    Dim stritem As String
    Dim strlot As String
    Dim strWafer As String
    Dim Strqty1 As String
    Dim Strqty2 As String
    Dim strtime1 As String
    Dim strperson1 As String
    Dim strtime2 As String
    Dim strperson2 As String

    

    
    strSql = " select item,lot,wafer,cast(qty1 as int) as qty1,cast(qty2 as int) as qty2,time1,person1,time2,person2   from  erpdata..zh_tblVT_temp where isnull(item,'')<>'' and isdate(time1)=1  and isdate(time2)=1      order by wafer  "
    If SMR.State = adStateOpen Then SMR.Close

    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
             stritem = Trim(SMR("item"))
             strlot = Trim(SMR("lot"))
             strWafer = Trim(SMR("wafer"))
             Strqty1 = SMR("qty1")
             Strqty2 = SMR("qty2")
             strtime1 = Trim(SMR("time1"))
             strperson1 = Trim(SMR("person1"))
             strtime2 = Trim(SMR("time2"))
             strperson2 = Trim(SMR("person2"))
             strSql = "insert into erpdata..zh_tblVT_modify(料号,qty外发,qty回货,申请人员,申请时间,工单号,流程卡编号,目标仓库) values('" & stritem & "'," & Strqty1 & ",0,'" & strperson1 & "','" & strtime1 & "','" & strlot & "','" & strWafer & "','72')"
           
             AddSql2 (strSql)
             strSql = "insert into erpdata..zh_tblVT_modify(料号,qty外发,qty回货,申请人员,申请时间,工单号,流程卡编号,原仓库) values('" & stritem & "',0," & Strqty1 & ",'" & strperson2 & "','" & strtime2 & "','" & strlot & "','" & strWafer & "','72')"
          
             AddSql2 (strSql)
             
             
             SMR.MoveNext
        Next
    End If
    SMR.Close
    
    Set SMR = Nothing
    
    
    
    
End Sub

Private Sub Command5_Click()
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    
    Dim stritem As String
    Dim strlot As String
    Dim strWafer As String
    Dim Strqty1 As String
    Dim Strqty2 As String
    Dim strtime1 As String
    Dim strperson1 As String
    Dim strtime2 As String
    Dim strperson2 As String

    

    
    strSql = " SELECT distinct a.流程卡编号,c.CUSTOMERSHORTNAME,c.MPN_DESC FROM erpdata..zh_tblVT_modify  a left JOIN ERPBASE..tblmappingData b ON  rtrim(a.流程卡编号)=rtrim(b.SUBSTRATEID) left JOIN ERPBASE..tblCustomerOI   c ON c.SOURCE_BATCH_ID =b.LOTID and CONVERT(nvarchar(20),c.id)=b.FILENAME   "
    If SMR.State = adStateOpen Then SMR.Close

    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount

             strSql = "update  erpdata..zh_tblVT_modify set 客户代码='" & Trim(SMR("CUSTOMERSHORTNAME")) & "'  where  流程卡编号 ='" & SMR("流程卡编号") & "'"
           
             AddSql2 (strSql)

             
             SMR.MoveNext
        Next
    End If
    SMR.Close
    
    Set SMR = Nothing
    

    
End Sub



Private Sub DTP1_Change()
 Call DTP2_Change
End Sub

Private Sub DTP2_Change()

    
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    
    Exit Sub
    
    CobPn.Clear
    
    If SMR.State = adStateOpen Then SMR.Close

 
strSql = "SELECT distinct c.料号 FROM erpdata..tblstockdbsub a " & _
" LEFT JOIN erpdata..tblstockdb b ON a.调拨编号=b.调拨编号 AND a.序号=b.序号 " & _
" LEFT JOIN erpdata ..tblPackMainInfSub c   ON a.流程卡编号=c.流程卡编号 " & _
" INNER JOIN erpbase..tblstock e ON b.目标仓库=e.库房代码 " & _
" INNER JOIN erpbase..tblstock f ON b.原仓库=f.库房代码 " & _
" WHERE  e.库房类型<>'非保税' AND f.库房类型<>'非保税' " & _
" AND  c.料号 IS NOT NULL AND  (b.目标仓库='72'  OR b.原仓库='72') " & _
" And b.申请时间<'" & Format(DTP2.Value + 1, "yyyy/mm/dd") & "' and  b.申请时间>'" & Format(DTP1.Value, "yyyy/mm/dd") & "'"
    
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            CobPn.AddItem (Trim(SMR("料号")))
            SMR.MoveNext
        Next
    End If
    SMR.Close
    
    Set SMR = Nothing
End Sub



Private Sub FpsToExcel()
    If Fps.MaxRows = 0 Then
        MsgBox "没有数据可以导出", vbInformation, "提示"
        Exit Sub
    End If

    Dim i As Long
    Dim j As Long
    
    Dim xlApp      As Excel.Application
    Dim xlBook     As Excel.Workbook
    Dim xlSheet    As Excel.Worksheet
    

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

    With xlApp
        .Rows(1).Font.Bold = True
    End With
    
' On Error GoTo Ert
    With Fps
        If .MaxRows > 5000 Then MsgBox "共有" & .MaxRows & "条记录需要导出，可能需要几分钟，请耐心等待", vbInformation, "提示"
        If .MaxCols > 10 Then
            For i = 0 To .MaxRows
                For j = 1 To .MaxCols - 4
                    .Col = j
                    .Row = i
                    xlSheet.Cells(i + 1, j) = Trim$(("'" & .text))
                        
                Next j
           
            Next i
        Else
            For i = 0 To .MaxRows
                For j = 1 To .MaxCols
                    .Col = j
                    .Row = i
                    xlSheet.Cells(i + 1, j) = Trim$(("'" & .text))
                Next j
           
            Next i
        End If
    End With

    '数字列格式调整
   For j = 1 To Fps.MaxCols
       If Trim(xlSheet.Cells(1, j)) = "qty外发" Or Trim(xlSheet.Cells(1, j)) = "qty回货" Or Trim(xlSheet.Cells(1, j)) = "金额" Or Trim(xlSheet.Cells(1, j)) = "发出数量" Or Trim(xlSheet.Cells(1, j)) = "回货数量" Then
            For i = 2 To Fps.MaxRows + 1
                xlSheet.Cells(i, j) = Replace(xlSheet.Cells(i, j), "'", "")
            Next
        End If
    Next
    With xlSheet.Range("2:" & Fps.MaxRows + 1)
        .horizontalAlignment = xlLeft
    End With
    xlSheet.Range("A1").Select
    xlApp.Columns.AutoFit
    
    xlApp.Application.Visible = True
    
    
    Set xlApp = Nothing  '"ユ临北倒Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing
'Ert:

 '   If Not (xlApp Is Nothing) Then
        
 '   Set xlApp = Nothing  '"ユ临北倒Excel
  '  Set xlBook = Nothing
  '  Set xlSheet = Nothing
  '  End If
    
    
End Sub



Private Sub Form_Load()
DTP1.Value = Now - 1
DTP2.Value = Now
'Call DTP2_Change
End Sub

Private Sub sumQty()
    Dim qty_out As Long
    Dim qty_back As Long
    Dim i As Long
    qty_out = 0
    qty_back = 0
    

    With Fps

            For i = 0 To .MaxRows
                .Row = i
                .Col = 3
                If IsNumeric(.text) = True Then
                    qty_out = qty_out + .text
                End If
                 .Row = i
                .Col = 4
                
                If IsNumeric(.text) = True Then
                    qty_back = qty_back + .text
                End If
                
            Next
     End With
     Label5.Caption = "合计：委外" & qty_out & "; 回货 " & qty_back
            
            
End Sub

Private Sub getdatafromstockdb()

    
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    Dim strWafer     As String
    Dim maxtime1 As String
    Dim time2 As String

strSql = " SELECT isnull(max(申请时间),'') FROM erpdata..zh_tblvt_new "
maxtime1 = GetSqlServerStr(strSql)

'同步最新的委外回货调拨数据
 strSql = " INSERT INTO erpdata..zh_tblvt_new " & _
" SELECT  DISTINCT j.MPN_DESC as 客户机种, c.料号, " & _
" CASE b.目标仓库 WHEN '72' THEN a.合格数+a.制程不良数+a.来料不良数   ELSE 0 END AS 'qty(外发)'  ,CASE b.目标仓库 WHEN '72' THEN 0   ELSE a.合格数+a.制程不良数+a.来料不良数 END AS  'qty(回货)' , " & _
" b.申请人员 ,b.申请时间,rtrim(a.工单号) AS 工单号 ,rtrim(a.流程卡编号) AS 流程卡编号 ,'', b.调拨编号,b.序号,b.原仓库 ,b.目标仓库 ,j.customershortname,1,'',0 " & _
" FROM erpdata..tblstockdbsub a " & _
" LEFT JOIN erpdata..tblstockdb b ON a.调拨编号=b.调拨编号 AND a.序号=b.序号 " & _
" LEFT JOIN erpdata ..tblPackMainInfSub c   ON a.流程卡编号=c.流程卡编号 " & _
" left JOIN ERPBASE..tblmappingData i ON  a.流程卡编号=i.SUBSTRATEID " & _
" left JOIN ERPBASE..tblCustomerOI   j ON j.SOURCE_BATCH_ID =i.LOTID and CONVERT(nvarchar(20),j.id)=i.FILENAME " & _
" WHERE isnull(c.料号,'')<>'' and left(c.大工单,1)='A' AND  ( b.原仓库='72' OR b.目标仓库='72' ) " & _
" AND b.申请时间>='" & maxtime1 & "' and a.调拨编号 not in (SELECT rtrim(调拨编号) FROM erptemp..InvalidStockDb) and a.调拨编号 not in (SELECT rtrim(关联调拨编号) FROM erptemp..InvalidStockDb) " & _
" ORDER BY RTRIM(a.流程卡编号),b.申请时间"


  AddSql2 (strSql)
  '同步回货调拨对应的外发时间
 strSql = " UPDATE t1 SET t1.外发时间=t2.申请时间 ,t1.flag =case year(t2.申请时间) when '2018' then 0 when '2019' then '0' else 1 end " & _
" FROM erpdata..zh_tblvt_new t1 " & _
" INNER JOIN (SELECT a.流程卡编号,CONVERT(varchar(100),max(c.申请时间) ,23) AS 申请时间 " & _
" FROM erpdata..zh_tblvt_new a " & _
" INNER JOIN erpdata..tblstockdbsub b ON a.流程卡编号=b.流程卡编号 " & _
" INNER JOIN erpdata..tblstockdb c ON b.调拨编号=c.调拨编号 AND b.序号=c.序号 " & _
" WHERE c.目标仓库 ='72' AND a.原仓库='72'   and  b.调拨编号 not in (SELECT rtrim(调拨编号) FROM erptemp..InvalidStockDb) and b.调拨编号 not in (SELECT rtrim(关联调拨编号) FROM erptemp..InvalidStockDb )" & _
" GROUP BY a.流程卡编号 )t2 ON t1.流程卡编号=t2.流程卡编号 " & _
" WHERE t1.原仓库='72' and t1.申请时间>='" & maxtime1 & "' "

 AddSql2 (strSql)
 
 
  '同步金额
 strSql = " UPDATE t1 set t1.金额=h.含税单价 " & _
" FROM erpdata..zh_tblvt_new t1 " & _
" inner JOIN ERPBASE..tblToInRec_Wafer g ON  t1.流程卡编号=g.晶圆ID  " & _
" inner JOIN ERPBASE..TblToInSub  h ON g.入库单编号 =h.入库单编号 and g.批号 =h.到货批号  " & _
" where h.入库单编号 not in ( select 关联入库单编号 from  ERPBASE..TblToInRec ) and  t1.申请时间>='" & maxtime1 & "'"

  AddSql2 (strSql)

'同步收货地址对应的企业名称
strSql = " UPDATE t1 set t1.委托方企业名称=CASE charindex('@',m.SHIP_TO_AD) WHEN 0 THEN m.SHIP_TO_AD  ELSE LEFT (m.SHIP_TO_AD,charindex('@',m.SHIP_TO_AD)-1) END   " & _
" FROM erpdata..zh_tblvt_new t1  " & _
" inner JOIN erptemp..tblstockdb_temp d ON d.remark2=t1.调拨编号  " & _
" inner join erptemp..customer_information m on t1.客户代码 =m.CUSTOMER and m.SHIP_TO=isnull(d.remark1,'') " & _
" where  t1.申请时间>='" & maxtime1 & "'"
  AddSql2 (strSql)
  
End Sub




