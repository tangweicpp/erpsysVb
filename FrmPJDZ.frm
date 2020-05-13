VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmPJDZ 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   11865
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21795
   FillColor       =   &H000000FF&
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
   ScaleHeight     =   11865
   ScaleWidth      =   21795
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtText5 
      Height          =   495
      Left            =   16560
      TabIndex        =   18
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtText4 
      Height          =   495
      Left            =   13080
      TabIndex        =   15
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "退出"
      Height          =   360
      Left            =   16320
      MaskColor       =   &H008080FF&
      MousePointer    =   1  'Arrow
      TabIndex        =   12
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "清空控件值"
      Height          =   600
      Left            =   2640
      TabIndex        =   11
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtText3 
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton command4 
      Caption         =   "修改信息"
      Height          =   600
      Left            =   13200
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton command3 
      Caption         =   "删除"
      Height          =   600
      Left            =   16200
      TabIndex        =   7
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton command2 
      Caption         =   "新增"
      Height          =   600
      Left            =   9360
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton command1 
      Caption         =   "查询"
      Height          =   600
      Left            =   5880
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtText2 
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtText1 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin FPSpreadADO.fpSpread Fps 
      Height          =   9135
      Index           =   0
      Left            =   960
      TabIndex        =   13
      Top             =   2760
      Width           =   18615
      _Version        =   524288
      _ExtentX        =   32835
      _ExtentY        =   16113
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
      SpreadDesigner  =   "FrmPJDZ.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label lblXH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "箱号"
      Height          =   195
      Left            =   16080
      TabIndex        =   17
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label lblWafer_lot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "wafer_lot"
      Height          =   195
      Left            =   12240
      TabIndex        =   16
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label laebl_head 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ht_wafer_mark"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   1485
   End
   Begin VB.Label label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "wafer_ID"
      Height          =   195
      Left            =   4680
      TabIndex        =   10
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "product（必填）"
      Height          =   195
      Left            =   8040
      TabIndex        =   4
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cust_id（必填）"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1230
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   45
   End
End
Attribute VB_Name = "FrmPJDZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheckData_Click()
   CheckData
'MsgBox "功能未开放"
End Sub

Private Sub CmdClear_Click()
Initial
End Sub

Private Sub CmdQuit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
  Query
End Sub
Private Sub Command2_Click()
    '新增
   '查询

    Dim strSql       As String

    Dim rs           As New ADODB.Recordset
    Dim rs_stock           As New ADODB.Recordset
    Dim cust_id         As String

    Dim product      As String

    Dim WAFER          As String
    
    Dim create_by       As String
  
    Dim CREATE_DATE     As String
    Dim REMARK1              As String
    Dim REMARK2              As String
    Dim flag_wafer As Integer
    Dim flag_lot As Integer
    Dim flag_boxid As Integer
    Dim wafer_exists       As String
    Dim intCount As Integer
    
    wafer_exists = ""
    intCount = 0
    cust_id = Trim(txtText1.text)
    product = Trim(txtText2.text)
    WAFER = Trim(txtText3.text)
    REMARK1 = Trim(txtText4.text)
    REMARK2 = Trim(txtText5.text)

    CREATE_DATE = DATE
    create_by = gUserName
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2
    End If
    
    
'客户代码，料号必填
    If cust_id = "" Or product = "" Then
        MsgBox "cust_id，product信息为必填！"
        Exit Sub
    End If
'流程卡编号，lot，箱号三选一，必须且只能填一项
    If WAFER <> "" Then
       flag_wafer = 1
    Else
       flag_wafer = 0
    End If
    
    If REMARK1 <> "" Then
       flag_lot = 1
    Else
       flag_lot = 0
    End If
    
    If REMARK2 <> "" Then
       flag_boxid = 1
    Else
       flag_boxid = 0
    End If
    If flag_wafer + flag_lot + flag_boxid > 1 Then
        MsgBox "wafer_id,wafer_lot,箱号 三个栏位只能有一个栏位有值", vbInformation, "提示"
        Exit Sub
    End If
    strSql = "select distinct a.流程卡编号,a.工单号 from erpdata..tblstocknumsub a , erpdata..tblstocknum  b where a.id=b.id and b.客户代码='" & cust_id & "' and a.料号='" & product & "'"
    If WAFER <> "" Then
        strSql = strSql & " and a.流程卡编号='" & WAFER & "'"
    End If
    
    If REMARK1 <> "" Then
        strSql = strSql & " and a.工单号='" & REMARK1 & "'"
    End If
    intCount = 0
    If REMARK2 <> "" Then
        strSql = strSql & " and a.箱号='" & REMARK2 & "'"
    End If
    strSql = strSql & " order by a.流程卡编号"
    Set rs_stock = Get_SqlserveRs(strSql)
    
    If rs_stock.RecordCount = 0 Then
        MsgBox "库存中查询不到您输入的信息，请确认", vbInformation, "提示"
        Exit Sub
        
    Else
        rs_stock.MoveFirst
        For i = 1 To rs_stock.RecordCount
           
           If Get_SqlserverCnt("select 1 from erptemp..ht_wafer_mark where 1=1 and wafer = '" & rs_stock("流程卡编号") & "'") > 0 Then
               wafer_exists = wafer_exists & "," & Trim(rs_stock("流程卡编号"))
           Else
                '信息插入表中
                strSql = "INSERT INTO erptemp..ht_wafer_mark (cust_id,  product, wafer,create_by,Create_date,REMARK1,REMARK2)" & _
                "values('" & cust_id & "','" & product & "','" & RTrim(rs_stock("流程卡编号")) & "','" & create_by & "','" & CREATE_DATE & "','" & rs_stock("工单号") & "','" & REMARK2 & "')"
                Exec_Sql (strSql)
                intCount = intCount + 1
           End If
                   
            rs_stock.MoveNext
        Next
    
    End If
    If intCount = 0 Then
        MsgBox "上传失败", vbInformation, "提示"
    Else
        If wafer_exists <> "" Then
            MsgBox "成功上传 " & intCount & " 笔，如下wafer已上传过" & wafer_exists, vbInformation, "提示"
         Else
         
            MsgBox "成功上传 " & intCount & " 笔", vbInformation, "提示"
         End If
    End If
    Query2

End Sub



Private Sub Command3_Click()

    Dim i      As Integer
    Dim strSql As String
    Dim strSql2 As String
    Dim cust_id         As String
    Dim count As Integer

    Dim product      As String

    Dim WAFER          As String
    
    Dim create_by       As String
  
    Dim CREATE_DATE     As String
    
    count = 0
    strSql2 = "delete from erptemp..ht_wafer_mark  where cust_id  = "

    With fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 2
                If Trim(.text) <> "" Then
                    strSql = strSql2 + "'" & Trim(.text) & "'  "
                End If
                .Col = 3
                If Trim(.text) <> "" Then
                    strSql = strSql + " and product ='" & Trim(.text) & "'  "
                End If
                .Col = 4
                If Trim(.text) <> "" Then
                    strSql = strSql + "and wafer = '" & Trim(.text) & "'  "
                End If
                
                 .Col = 5
                If Trim(.text) <> "" Then
                    strSql = strSql + "and REMARK1 = '" & Trim(.text) & "' "
                End If
                
                
                .Col = 6
                If Trim(.text) <> "" Then
                    strSql = strSql + "and create_by = '" & Trim(.text) & "'  "
                End If
                .Col = 7
                If Trim(.text) <> "" Then
                    strSql = strSql + "and create_date = '" & Trim(.text) & "'  "
                End If



               ' .Col = 9
              '  If Trim(.text) <> "" Then
               '     strSql = strSql + "and REMARK2 = '" & Trim(.text) & "' "
              '  End If
                
                If AddSql2(strSql) > -1 Then
                        count = count + 1
                End If
            End If
            
            strSql = strSql2
        Next i
    End With

        If count = 0 Then
            MsgBox "删除失败"
        Else
            MsgBox "删除成功" & "记录数" & count & "! "
        End If
    Query2

End Sub

Private Sub Command4_Click()

    Dim rs        As New ADODB.Recordset

    '修改
    Dim i         As Integer

    Dim strSql    As String

    Dim cust_id         As String

    Dim product      As String
    
    Dim REMARK1   As String
    
    Dim REMARK2  As String
    
    Dim REMARK3  As String
    
    Dim REMARK4  As String

    Dim REMARK5  As String

    Dim count As Integer
    
    count = 0
    
    With fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .Value = 1 Then
                .Col = 2
                If Trim(.text) <> "" Then
                    cust_id = Trim(.text)
                End If
                
                .Col = 3
                If Trim(.text) <> "" Then
                    product = Trim(.text)
                End If
                
                .Col = 4
                If Trim(.text) = "" Then
                    MsgBox "error in 修改!"
                    Exit Sub
                Else
                    WAFER = Trim(.text)
                End If
  
                .Col = 5
                If Trim(.text) <> "" Then
                    REMARK1 = Trim(.text)
                End If

              '  .Col = 9
               ' If Trim(.text) <> "" Then
               '     REMARK2 = Trim(.text)
              '  End If
               REMARK3 = ""
               REMARK4 = ""
               REMARK5 = ""
                .Col = 9
                If Trim(.text) <> "" Then
                    REMARK3 = Trim(.text)
                End If
                
                .Col = 10
                If Trim(.text) <> "" Then
                    REMARK4 = Trim(.text)
                End If

                .Col = 11
                If Trim(.text) <> "" Then
                    REMARK5 = Trim(.text)
                End If
              

  
                If cust_id = "" Or product = "" Then
                    MsgBox "主要信息都为必填！"
                Else
                   strSql = "select * from erptemp..ht_wafer_mark where wafer = '" & WAFER & "'"

                    If INIadoCon.State <> adStateOpen Then
                        INIConnectSTART2
                    End If

                    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                    
                    If rs.RecordCount > 0 Then
                        rs.Close
                     '  strSql = "UPDATE erptemp..ht_wafer_mark SET cust_id ='" & cust_id & "'," & " product ='" & product & "'," & " REMARK1 ='" & REMARK1 & _
                        "',REMARK2 ='" & REMARK2 & "',REMARK3 ='" & REMARK3 & "', REMARK4 = '" & REMARK4 & "',REMARK5 = '" & REMARK5 & "' where wafer='" & WAFER & "'"
                       strSql = "UPDATE erptemp..ht_wafer_mark SET REMARK3 ='" & REMARK3 & "', REMARK4 = '" & REMARK4 & "',REMARK5 = '" & REMARK5 & "' where wafer='" & WAFER & "'"
                        AddSql2 (strSql)
                        
                        count = count + 1

                    End If

                End If
            End If

        Next i

    End With

    If count = 0 Then
        MsgBox "修改失败"
    Else
        MsgBox "修改成功" & "修改记录数" & count & "! "
    
    End If

Query2

End Sub
Private Sub ListDataType(rs As ADODB.Recordset)
   Dim i As Long

    With fps(0)
        .MaxRows = 0
        Set .DataSource = rs

    End With

    With fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .BackColor = &HFFFF&
            .ColWidth(1) = 10
            .CellType = CellTypeCheckBox
            .text = 0
            .Col = 4
            .Lock = True
            .Col = 5
            .Lock = True
            .Col = 6
            .Lock = True
            .Col = 7
            .Lock = True
        
        Next
        
    End With
    rs.Close
End Sub
Private Function Query()
   '查询

    Dim strSql       As String

    Dim rs           As New ADODB.Recordset
    
    Dim cust_id         As String

    Dim product      As String

    Dim WAFER          As String
    
    Dim create_by       As String
  
    Dim Create_time     As String
    
    cust_id = Trim(txtText1.text)
    product = Trim(txtText2.text)
    WAFER = Trim(txtText3.text)


    strSql = "select '' AS '选择',a.cust_id as 'cust_id', a.product as 'product', a.wafer  as 'wafer',a.remark1 as wafer_lot ,a.create_by as '创建人'," & _
    " a.Create_date as '创建时间', a.last_update_by  ,a.remark3,a.remark4,a.remark5 from erptemp..ht_wafer_mark  a where 1 = 1"

    If Trim(txtText1.text) <> "" Then
        strSql = strSql + " AND  a.cust_id   = '" & Trim(txtText1.text) & "'  "

    End If

    If Trim(txtText2.text) <> "" Then
        strSql = strSql + " AND a.product = '" & Trim(txtText2.text) & "'  "

    End If

    If Trim(txtText3.text) <> "" Then
        strSql = strSql + " AND a.wafer  = '" & Trim(txtText3.text) & "'  "

    End If
      
    If Trim(txtText4.text) <> "" Then
        strSql = strSql + " AND REMARK1  = '" & Trim(txtText4.text) & "'  "

    End If
    
   If Trim(txtText5.text) <> "" Then
       'strSql = strSql + " AND REMARK2  = '" & Trim(txtText5.text) & "'  "
        strSql = strSql + " AND  a.wafer in (select distinct 流程卡编号 from erpdata..tblstocknumsub where 箱号='" & Trim(txtText5.text) & "')"
    End If
 

    strSql = strSql + " order by Create_date desc"
  

    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If
   
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType(rs)
        Exit Function
    End If

End Function

'初始化
Private Function Initial()

    txtText1.text = ""
    txtText2.text = ""
    txtText3.text = ""

End Function
Private Function Query2()
    Dim strSql       As String

    Dim rs           As New ADODB.Recordset
    strSql = "select '' AS '选择',cust_id as 'cust_id', product as 'product', wafer  as 'wafer',remark1 as 'Wafer_lot',create_by as '创建人'," & _
    "Create_date as '创建时间', last_update_by,remark3,remark4,remark5 from erptemp..ht_wafer_mark where 1 = 1 " & _
    "order by create_date desc"

    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType(rs)
        rs.Close
        Exit Function

    End If
End Function
Private Function CheckData()

MsgBox "功能尚未开放"

'    Dim i As Long
'
'    With Fps(0)
'        .MaxRows = 0
'        Set .DataSource = rs
'
'    End With
'
'    With Fps(0)
'
'        For i = 1 To .MaxRows
'            .Row = i
'            .Col = 1
'            .BackColor = &HFFFF&
'            .ColWidth(1) = 10
'            .CellType = CellTypeCheckBox
'            .Text = 0
'            .Col = 10
'            .Lock = False
'
'        Next
'
'    End With


End Function



