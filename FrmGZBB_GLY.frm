VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmGZBB_GLY 
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
   Begin VB.TextBox txtText4 
      Height          =   435
      Left            =   1800
      TabIndex        =   16
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CheckBox chkCheck1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.ComboBox BMcode 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      ItemData        =   "FrmGZBB_GLY.frx":0000
      Left            =   1800
      List            =   "FrmGZBB_GLY.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2760
      Width           =   1335
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
      Left            =   1800
      TabIndex        =   11
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox txtText3 
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton command4 
      Caption         =   "修改信息"
      Height          =   600
      Left            =   10440
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton command3 
      Caption         =   "取消权限"
      Height          =   600
      Left            =   7800
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton command2 
      Caption         =   "新增"
      Height          =   600
      Left            =   2760
      TabIndex        =   6
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton command1 
      Caption         =   "查询"
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1935
   End
   Begin VB.TextBox txtText2 
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txtText1 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   2775
   End
   Begin FPSpreadADO.fpSpread Fps 
      Height          =   8175
      Index           =   0
      Left            =   5160
      TabIndex        =   18
      Top             =   840
      Width           =   14175
      _Version        =   524288
      _ExtentX        =   25003
      _ExtentY        =   14420
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
      SpreadDesigner  =   "FrmGZBB_GLY.frx":003D
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "部门"
      Height          =   195
      Left            =   1320
      TabIndex        =   17
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "申请原因"
      Height          =   195
      Left            =   960
      TabIndex        =   13
      Top             =   3960
      Width           =   720
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4920
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      X1              =   4920
      X2              =   4920
      Y1              =   0
      Y2              =   5160
   End
   Begin VB.Label label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "申请人"
      Height          =   195
      Left            =   1200
      TabIndex        =   10
      Top             =   3360
      Width           =   540
   End
   Begin VB.Label label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工号"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   360
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
Attribute VB_Name = "FrmGZBB_GLY"
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

    Dim strSql       As String

    Dim rs           As New ADODB.Recordset
    
    Dim GH         As String

    Dim XM      As String

    Dim BM          As String
    
    Dim SQR       As String
    
    Dim GZBBQX   As String
    
    Dim reason     As String
    
    GH = Trim(txtText1.Text)
    XM = Trim(txtText2.Text)
    BM = Trim(BMcode.Text)
    SQR = Trim(txtText3.Text)
    reason = Trim(txtText4.Text)
    GZBBQX = "1"
    
    Create_time = DATE

    strSql = "select * from erptemp.dbo.KJQX where 1=1 and GH = '" & GH & "'"
    
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2
    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If GH = "" Or XM = "" Or BM = "" Or SQR = "" Then
        MsgBox "信息都为必填！"
        Exit Sub
    ElseIf rs.RecordCount > 0 Then
        If rs.Fields("GZBBQX") = 1 Then
            MsgBox "数据已存在！"
            Exit Sub
        ElseIf rs.Fields("GZBBQX") = 0 Then
            'Dim GH As String
            'Dim reason As String
            GH = rs.Fields("GH")
         '   reason = rs.Fields("reason")
            strSql = "UPDATE erptemp.dbo.KJQX SET GZBBQX = 1 where  GH = '" & GH & "' and XM = '" & XM & "' and BM = '" & BM & "' and SQR = '" & SQR & "' and reason = '" & reason & "'"
            Exec_Sql (strSql)
            'MsgBox "权限已更新！时间为" & YStime
        End If
    Else
        '信息插入表中
        strSql = "INSERT INTO erptemp.dbo.KJQX (GH,  XM, BM,SQR,GZBBQX,Create_time,reason)" & _
        "values('" & GH & "','" & XM & "','" & BM & "','" & SQR & "','" & GZBBQX & "','" & Create_time & "','" & reason & "')"
        Exec_Sql (strSql)
    End If
   query2

End Sub
Private Sub Command3_Click()

    '取消权限
    Dim i      As Integer
    Dim strSql As String
    Dim strsql2 As String
    Dim GH As String
    Dim XM As String
    Dim count  As Integer
    
    count = 0
    strsql2 = "update erptemp.dbo.KJQX set GZBBQX = '0' where GH  = "

    With Fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 2
                If Trim(.Text) <> "" Then
                    strSql = strsql2 + "'" & Trim(.Text) & "'  "
                End If
'                If Trim(.Text) <> "" Then
'                   XM = Trim(.Text)
'                End If
'
                If AddSql2(strSql) > -1 Then
                        count = count + 1
                End If
            End If
            '.Col = 1
            strSql = strsql2
        Next i
    End With

        If count = 0 Then
            MsgBox "取消权限失败"
        Else
            MsgBox "取消权限成功" & "记录数" & count & "! "
        End If
    query2

End Sub

Private Sub Command4_Click()

    Dim rs        As New ADODB.Recordset

    '修改
    Dim i         As Integer

    Dim strSql    As String

    Dim GH         As String
    Dim XM      As String
    Dim BM          As String
    Dim SQR       As String
    Dim reason     As String
    
    Dim count As Integer
    
    count = 0
    
    With Fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 2
                If Trim(.Text) <> "" Then
                    GH = Trim(.Text)
                End If
                
                .Col = 3
                If Trim(.Text) <> "" Then
                    XM = Trim(.Text)
                End If

                .Col = 4
                If Trim(.Text) <> "" Then
                    BM = Trim(.Text)
                End If

                .Col = 5
                If Trim(.Text) <> "" Then
                    SQR = Trim(.Text)
                End If

                .Col = 6
                If Trim(.Text) <> "" Then
                    reason = Trim(.Text)
                End If

              End If

  
                If XM = "" Or BM = "" Or SQR = "" Or reason = "" Then
                    MsgBox "主要信息都为必填！"
                Else
                   strSql = "select * from erptemp.dbo.KJQX where GH = '" & GH & "'"

                    If INIadoCon.State <> adStateOpen Then
                        INIConnectSTART2
                    End If

                    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                    
                    If rs.RecordCount > 0 Then
                        rs.Close
                       strSql = "UPDATE erptemp.dbo.KJQX SET XM='" & XM & "'," & " BM ='" & BM & "'," & " SQR ='" & SQR & _
                        "',reason ='" & reason & "' where GH='" & GH & "'"
                        AddSql2 (strSql)
                        
                        count = count + 1


                    End If

                End If


        Next i

    End With

    If count = 0 Then
        MsgBox "修改失败"
    Else
        MsgBox "修改成功" & "修改记录数" & count & "! "
    
    End If

query2

End Sub
Private Sub ListDataType(rs As ADODB.Recordset)
   Dim i As Long

    With Fps(0)
        .MaxRows = 0
        Set .DataSource = rs

    End With

    With Fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .BackColor = &HFFFF&
            .ColWidth(1) = 10
            .CellType = CellTypeCheckBox
            .Text = 0
            .Col = 2
            .Lock = True
        
        Next
        
    End With
    rs.Close
End Sub
Private Function Query()
   '查询
 
    Dim strSql       As String

    Dim rs           As New ADODB.Recordset
    
    Dim GH         As String

    Dim XM      As String

    Dim BM          As String
    
    Dim SQR       As String
  
    Dim reason     As String
    
    GH = Trim(txtText1.Text)
    XM = Trim(txtText2.Text)
    BM = Trim(BMcode.Text)
    SQR = Trim(txtText3.Text)
    reason = Trim(txtText4.Text)
    
    Create_time = DATE
  

    strSql = "select '' AS '选择',GH as '工号', XM as '姓名', BM as '部门',SQR as '申请人'," & _
    "reason as '原因' from erptemp.dbo.KJQX where GZBBQX = '1' "

    If Trim(txtText1.Text) <> "" Then
        strSql = strSql + " AND  GH  = '" & Trim(txtText1.Text) & "'  "

    End If

    If Trim(txtText2.Text) <> "" Then
        strSql = strSql + " AND XM = '" & Trim(txtText2.Text) & "'  "

    End If

    If Trim(txtText3.Text) <> "" Then
        strSql = strSql + " AND SQR  = '" & Trim(txtText3.Text) & "'  "

    End If
    
    If chkCheck1 = 1 Then
        If Trim(BMcode.Text) <> "" Then
            strSql = strSql + " AND BM  = '" & Trim(BMcode.Text) & "'  "
    
        End If
    End If
    
    If Trim(txtText3.Text) <> "" Then
        strSql = strSql + " AND reason  = '" & Trim(txtText3.Text) & "'  "

    End If
        

    strSql = strSql + " order by Create_time desc"
  

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

    txtText1.Text = ""
    txtText2.Text = ""
    txtText3.Text = ""
    txtText4.Text = ""

End Function
Private Function query2()
    Dim strSql       As String

    Dim rs           As New ADODB.Recordset
     strSql = "select '' AS '选择',GH as '工号', XM as '姓名', BM as '部门',SQR as '申请人'," & _
    "reason as '原因',Create_time as '创建日期' from erptemp.dbo.KJQX where GZBBQX = '1' order by Create_time desc"

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

