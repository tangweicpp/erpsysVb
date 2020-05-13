VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmOnMPN 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MPN Attributes"
   ClientHeight    =   12780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   24810
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
   ScaleHeight     =   12780
   ScaleWidth      =   24810
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txt12_Cbo 
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
      ItemData        =   "FrmOnMPN.frx":0000
      Left            =   15360
      List            =   "FrmOnMPN.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   1560
      Width           =   2535
   End
   Begin VB.ComboBox txt8_Cbo 
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
      ItemData        =   "FrmOnMPN.frx":0016
      Left            =   10440
      List            =   "FrmOnMPN.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox txt7_Cbo 
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
      ItemData        =   "FrmOnMPN.frx":002C
      Left            =   10440
      List            =   "FrmOnMPN.frx":0036
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox txt3_cbo 
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
      ItemData        =   "FrmOnMPN.frx":0042
      Left            =   2400
      List            =   "FrmOnMPN.frx":004C
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton cmd_report 
      BackColor       =   &H0000FF00&
      Caption         =   "指定导出"
      Height          =   600
      Left            =   14880
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmd_clear 
      BackColor       =   &H008080FF&
      Caption         =   "清空控件值"
      Height          =   600
      Left            =   16680
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "退出"
      Height          =   600
      Left            =   19200
      TabIndex        =   28
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Modify 
      Caption         =   "修改"
      Height          =   600
      Left            =   10440
      TabIndex        =   27
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton CmdDel 
      BackColor       =   &H000000FF&
      Caption         =   "删除"
      Height          =   240
      Left            =   8880
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "新增"
      Height          =   600
      Left            =   6840
      TabIndex        =   25
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtText11 
      Height          =   375
      Left            =   15360
      TabIndex        =   24
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtText9 
      Height          =   375
      Left            =   10440
      TabIndex        =   22
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtText10 
      Height          =   375
      Left            =   15360
      TabIndex        =   21
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtText6 
      Height          =   375
      Left            =   6120
      TabIndex        =   20
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtText5 
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtText4 
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtText2 
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtText1 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton CmdOutReport 
      BackColor       =   &H0000FF00&
      Caption         =   "导出全部"
      Height          =   600
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查询"
      Height          =   600
      Left            =   5040
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   9855
      Index           =   0
      Left            =   -240
      TabIndex        =   0
      Top             =   2880
      Width           =   24975
      _Version        =   524288
      _ExtentX        =   44053
      _ExtentY        =   17383
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
      SpreadDesigner  =   "FrmOnMPN.frx":0058
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label lblUL 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UL_LISTED_FLAG"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13440
      TabIndex        =   23
      Top             =   1560
      Width           =   1470
   End
   Begin VB.Label lbl13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PKG_GRP_CD "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13440
      TabIndex        =   19
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label lbl12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PACKAGING_TYPE"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13440
      TabIndex        =   18
      Top             =   600
      Width           =   1470
   End
   Begin VB.Label lbl9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MPQ_QTY  "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   8880
      TabIndex        =   17
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label lbl8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PBF_DIE_ATTACH   "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   16
      Top             =   1080
      Width           =   1800
   End
   Begin VB.Label lbl7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HALIDE_FREF  "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   15
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label lb6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TEMP  "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   5400
      TabIndex        =   14
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MSL "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   5400
      TabIndex        =   13
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FrontSideMarking  "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   12
      Top             =   1080
      Width           =   1890
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1890
   End
   Begin VB.Label lblFrontSideMarking 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ECAT  "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5400
      TabIndex        =   10
      Top             =   600
      Width           =   630
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1890
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LEAD_FREF  "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   1560
      Width           =   1170
   End
   Begin VB.Label lblLabel1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PART"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   420
   End
End
Attribute VB_Name = "FrmOnMPN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_clear_Click()
  Initial
End Sub

Private Sub cmd_report_Click()
  report_1
End Sub

Private Sub CmdQuit_Click()
  Unload Me
End Sub

Private Sub cmdQuery_Click()
  Query
End Sub
'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CmdAdd_Click
' Description:       MPN attributes orancle表 CUSTOMERMPNAttributes
' Created by :       祝祎凡
' Machine    :       DESKTOP-F6L8S2V
' Date-Time  :       2019/10/15-9:45:43
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub CmdAdd_Click()
    '新增

    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    
    Dim ID       As Integer
    Dim PART         As String
    Dim MARKINGCODEFIRST      As String
    Dim LEAD_FREE          As String
    Dim ECAT        As String
    Dim MSL      As String
    Dim TEMP       As String
    Dim HALIDE_FREE    As String
    Dim PBF_DIE_ATTACH   As String
    Dim MPQ_QTY         As String
    Dim PACKAGING_TYPE  As String
    Dim PKG_GRP_CD     As String
    Dim UL_LISTED_FLAG    As String
    
    '参数判断
    
    PART = Trim(txtText1.Text)
    MARKINGCODEFIRST = Trim(txtText2.Text)    'FrontSideMarking
    LEAD_FREE = Trim(txt3_cbo.Text)
    ECAT = Trim(txtText4.Text)
    MSL = Trim(txtText5.Text)
    TEMP = Trim(txtText6.Text)
    HALIDE_FREE = Trim(txt7_Cbo.Text)
    PBF_DIE_ATTACH = Trim(txt8_Cbo.Text)
    MPQ_QTY = Trim(txtText9.Text)
    PACKAGING_TYPE = Trim(txtText10.Text)
    PKG_GRP_CD = Trim(txtText11.Text)
    UL_LISTED_FLAG = Trim(txt12_Cbo.Text)

    If MSL <> "" Then
       If Not IsNumeric(MSL) Then
         MsgBox "MSL不是数值型数据！"
         Exit Sub
       ElseIf (CStr(CInt(MSL)) <> MSL) Then
         MsgBox "MSL不是整数！"
         Exit Sub
       Else
         MSL = "'" + MSL + "'"
       End If
    Else
         MSL = "NULL"
    End If
    
    If TEMP <> "" Then
        If Not IsNumeric(TEMP) Then
            MsgBox "TEMP不是数值型数据！"
            Exit Sub
        ElseIf (CStr(CInt(TEMP)) <> TEMP) Then
            MsgBox "TEMP不是整数！"
            Exit Sub
        Else
            TEMP = "'" + TEMP + "'"
        End If
    Else
        TEMP = "NULL"
    End If
    
    If MPQ_QTY <> "" Then
        If Not IsNumeric(MPQ_QTY) Then
            MsgBox "MPQ_QTY不是数值型数据！"
            Exit Sub
        ElseIf MPQ_QTY <> "" And (CStr(CInt(MPQ_QTY)) <> MPQ_QTY) Then
            MsgBox "MPQ_QTY不是整数！"
            Exit Sub
        Else
            MPQ_QTY = "'" + MPQ_QTY + "'"
        End If
    Else
        MPQ_QTY = "NULL"
    End If

    


    '获取最大ID
    strSql = "select max(ID) as ""ID""  from  CUSTOMERMPNAttributes"
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    ID = rs.Fields("ID")
    ID = ID + 1
    rs.Close
    
    strSql = "select * from  CUSTOMERMPNAttributes where 1=1 and PART = '" & PART & "' and MARKINGCODEFIRST = '" & MARKINGCODEFIRST & _
    "' and LEAD_FREE = '" & LEAD_FREE & "' and ECAT = '" & ECAT & "' and MSL = " & MSL & " and TEMP = " & TEMP & " and " & _
    " HALIDE_FREE = '" & HALIDE_FREE & "' and PBF_DIE_ATTACH = '" & PBF_DIE_ATTACH & "' and MPQ_QTY = " & MPQ_QTY & " and " & _
    " PACKAGING_TYPE = '" & PACKAGING_TYPE & "' and PKG_GRP_CD = '" & PKG_GRP_CD & "' and UL_LISTED_FLAG = '" & UL_LISTED_FLAG & _
    "'"
    If Cnn.State = 0 Then
      ConOracle
    End If
    
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
  
  
'    If PART = "" Or MARKINGCODEFIRST = "" Or LEAD_FREE = "" Or ECAT = "" Or TEMP = "" Or MSL = "" Or HALIDE_FREE = "" Or PBF_DIE_ATTACH = "" Then
    If LEAD_FREE = "" Or ECAT = "" Or TEMP = "" Or MSL = "" Or HALIDE_FREE = "" Or PBF_DIE_ATTACH = "" Then
        MsgBox "LEAD_FREE,ECAT,TEMP,MSL,HALIDE_FREE,PBF_DIE_ATTACH信息都为必填！"
        Exit Sub
    ElseIf rs.RecordCount > 0 Then
        If rs.Fields("flag") = 0 Then
            MsgBox "数据已存在！"
            Exit Sub
        ElseIf rs.Fields("flag") = 1 Then
            'Dim ID As String
            ID = rs.Fields("ID")
            strSql = "UPDATE erptemp.dbo.MASK_CODE SET flag = 'Y' where  ID = '" & ID & "'"
'            strSql = "UPDATE CUSTOMERMPNAttributes SET PART = '" & PART & "',MARKINGCODEFIRST='" & MARKINGCODEFIRST & "'" & ", LEAD_FREE ='" & LEAD_FREE & "'" & ",ECAT ='" & ECAT & _
'                       "'" & ",MSL = " & MSL & " ,TEMP = " & TEMP & " ,HALIDE_FREE ='" & HALIDE_FREE & "',PBF_DIE_ATTACH ='" & PBF_DIE_ATTACH & "',MPQ_QTY = " & MPQ_QTY & _
'                       " ,PACKAGING_TYPE ='" & PACKAGING_TYPE & "',PKG_GRP_CD ='" & PKG_GRP_CD & "',UL_LISTED_FLAG ='" & UL_LISTED_FLAG & "' where ID='" & ID & "'"
            Exec_Ora (strSql)
            MsgBox "数据已恢复"
        End If
    Else
        '信息插入表中
       
    strSql = "INSERT INTO  CUSTOMERMPNAttributes (ID,LOC, PART,MarkingCodeFirst,LEAD_FREE,ECAT,MSL,TEMP,HALIDE_FREE," & _
    "PBF_DIE_ATTACH,MPQ_QTY,PACKAGING_TYPE,PKG_GRP_CD,UL_LISTED_FLAG,flag)" & _
    "values('" & ID & "','P2','" & PART & "','" & MARKINGCODEFIRST & "','" & LEAD_FREE & "','" & ECAT & "'," & MSL & "," & TEMP & ",'" & HALIDE_FREE & _
    "','" & PBF_DIE_ATTACH & "'," & MPQ_QTY & ",'" & PACKAGING_TYPE & "','" & PKG_GRP_CD & "','" & UL_LISTED_FLAG & "','Y')"
        
      Exec_Ora (strSql)
      rs.Close
    End If
   query2

End Sub
Private Sub cmdDel_Click()

    '删除
    Dim i      As Integer
    Dim strSql As String
    Dim strsql2 As String
    Dim count  As Integer

    count = 0
    strsql2 = "update CUSTOMERMPNAttributes set flag = 'N' where ID  = "


    With fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 2
                If Trim(.Text) <> "" Then
                    strSql = strsql2 + "'" & Trim(.Text) & "'  "
                End If


                If AddSql(strSql) > -1 Then
                        count = count + 1
                End If
            End If
            '.Col = 1
            strSql = strsql2
        Next i
    End With

        If count = 0 Then
            MsgBox "删除失败"
        Else
            MsgBox "删除成功" & "删除记录数" & count & "! "
        End If
    query2

End Sub

Private Sub cmd_Modify_Click()

    Dim rs        As New ADODB.Recordset

    '修改
    Dim i         As Integer

    Dim strSql    As String
    
    Dim PART         As String
    Dim MARKINGCODEFIRST      As String
    Dim LEAD_FREE          As String
    Dim ECAT        As String
    Dim MSL      As String
    Dim TEMP       As String
    Dim HALIDE_FREE    As String
    Dim PBF_DIE_ATTACH   As String
    Dim MPQ_QTY         As String
    Dim PACKAGING_TYPE  As String
    Dim PKG_GRP_CD     As String
    Dim UL_LISTED_FLAG    As String
    
    Dim count As Integer
    
    count = 0
    
    With fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 2
                If Trim(.Text) <> "" Then
                    ID = Trim(.Text)
                End If
                
                .Col = 4
                If Trim(.Text) <> "" Then
                    PART = Trim(.Text)
                End If

                .Col = 5
                If Trim(.Text) <> "" Then
                    MARKINGCODEFIRST = Trim(.Text)
                End If

                .Col = 6
                If Trim(.Text) <> "" Then
                    LEAD_FREE = Trim(.Text)
                End If

                .Col = 7
                If Trim(.Text) <> "" Then
                    ECAT = Trim(.Text)
                End If

                .Col = 8
                If Trim(.Text) <> "" Then
                    MSL = Trim(.Text)
                End If

                .Col = 9
                If Trim(.Text) <> "" Then
                    TEMP = Trim(.Text)
                End If

                .Col = 10
                If Trim(.Text) <> "" Then
                    HALIDE_FREE = Trim(.Text)
                End If
                
                .Col = 11
                If Trim(.Text) <> "" Then
                    PBF_DIE_ATTACH = Trim(.Text)
                End If
                
                .Col = 12
                If Trim(.Text) <> "" Then
                    MPQ_QTY = Trim(.Text)
                End If
                
                .Col = 13
                If Trim(.Text) <> "" Then
                    PACKAGING_TYPE = Trim(.Text)
                End If
                
                .Col = 14
                If Trim(.Text) <> "" Then
                    PKG_GRP_CD = Trim(.Text)
                End If
                
                .Col = 15
                If Trim(.Text) <> "" Then
                    UL_LISTED_FLAG = Trim(.Text)
                End If
                
                '查看修改后数值参数是否填错
                   If MSL <> "" Then
                        If Not IsNumeric(MSL) Then
                            MsgBox "MSL不是数值型数据！"
                            Exit Sub
                        ElseIf (CStr(CInt(MSL)) <> MSL) Then
                            MsgBox "MSL不是整数！"
                            Exit Sub
                        Else
                            MSL = "'" + MSL + "'"
                        End If
                   Else
                        MSL = "NULL"
                   End If
                
                    If TEMP <> "" Then
                        If Not IsNumeric(TEMP) Then
                            MsgBox "TEMP不是数值型数据！"
                            Exit Sub
                        ElseIf (CStr(CInt(TEMP)) <> TEMP) Then
                            MsgBox "TEMP不是整数！"
                            Exit Sub
                        Else
                            TEMP = "'" + TEMP + "'"
                        End If
                    Else
                        TEMP = "NULL"
                    End If
                
                    If MPQ_QTY <> "" Then
                        If Not IsNumeric(MPQ_QTY) Then
                            MsgBox "MPQ_QTY不是数值型数据！"
                            Exit Sub
                        ElseIf MPQ_QTY <> "" And (CStr(CInt(MPQ_QTY)) <> MPQ_QTY) Then
                            MsgBox "MPQ_QTY不是整数！"
                            Exit Sub
                        Else
                            MPQ_QTY = "'" + MPQ_QTY + "'"
                        End If
                    Else
                        MPQ_QTY = "NULL"
                    End If

            
'                If PART = "" Then
'    '              If PART = "" Or LEAD_FREE = "" Or ECAT = "" Or HALIDE_FREE = "" Or PBF_DIE_ATTACH = "" Or PACKING_TYPE = "" Or PKG_GRP_CD = "" Or UL_LISTED_FLAG = "" Then
'                    MsgBox "PART为必填！"
'                Else
                   strSql = "select * from CUSTOMERMPNAttributes where ID = '" & ID & "'"

                     If Cnn.State = 0 Then
                        ConOracle
                     End If
        
                    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
                    If rs.RecordCount > 0 Then
                        rs.Close
                       strSql = "UPDATE CUSTOMERMPNAttributes SET PART = '" & PART & "',MARKINGCODEFIRST='" & MARKINGCODEFIRST & "'" & ", LEAD_FREE ='" & LEAD_FREE & "'" & ",ECAT ='" & ECAT & _
                       "'" & ",MSL = " & MSL & " ,TEMP = " & TEMP & " ,HALIDE_FREE ='" & HALIDE_FREE & "',PBF_DIE_ATTACH ='" & PBF_DIE_ATTACH & "',MPQ_QTY = " & MPQ_QTY & _
                       " ,PACKAGING_TYPE ='" & PACKAGING_TYPE & "',PKG_GRP_CD ='" & PKG_GRP_CD & "',UL_LISTED_FLAG ='" & UL_LISTED_FLAG & "' where ID='" & ID & "'"
                        AddSql (strSql)
                        count = count + 1
                    Else
                        '修改了主键光照版本
                        MsgBox "修改失败,不能修改ID 第" & .Row & "行"

                    End If

'                End If

                End If

        Next i

    End With

    If count = 0 Then
        MsgBox "修改失败"
    Else
        MsgBox "修改成功" & "修改记录数" & count & "! "
    
    End If

Query

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
            .Text = 0
            .Col = 2
            .Lock = True
            .Col = 3
            .Lock = True
        
        Next
        
    End With
    rs.Close
End Sub
Private Function Query()
   '查询

    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    
    Dim PART         As String
    Dim MARKINGCODEFIRST      As String
    Dim LEAD_FREE          As String
    Dim ECAT        As String
    Dim MSL      As String
    Dim TEMP       As String
    Dim HALIDE_FREE    As String
    Dim PBF_DIE_ATTACH   As String
    Dim MPQ_QTY         As String
    Dim PACKAGING_TYPE  As String
    Dim PKG_GRP_CD     As String
    Dim UL_LISTED_FLAG    As String

    PART = Trim(txtText1.Text)
    MARKINGCODEFIRST = Trim(txtText2.Text)    'FrontSideMarking
    LEAD_FREE = Trim(txt3_cbo.Text)
    ECAT = Trim(txtText4.Text)
    MSL = Trim(txtText5.Text)
    TEMP = Trim(txtText6.Text)
    HALIDE_FREE = Trim(txt7_Cbo.Text)
    PBF_DIE_ATTACH = Trim(txt8_Cbo.Text)
    MPQ_QTY = Trim(txtText9.Text)
    PACKAGING_TYPE = Trim(txtText10.Text)
    PKG_GRP_CD = Trim(txtText11.Text)
    UL_LISTED_FLAG = Trim(txt12_Cbo.Text)
    
    strSql = "select ''as ""选择"", id, LOC,PART,MarkingCodeFirst,LEAD_FREE,ECAT,MSL,TEMP,HALIDE_FREE,PBF_DIE_ATTACH,MPQ_QTY," & _
  "PACKAGING_TYPE,PKG_GRP_CD,UL_LISTED_FLAG  from  CUSTOMERMPNAttributes where flag='Y'"
  
    If Trim(txtText1.Text) <> "" Then
        strSql = strSql + " AND PART  = '" & Trim(txtText1.Text) & "'  "

    End If

    If Trim(txtText2.Text) <> "" Then
        strSql = strSql + " AND MarkingCodeFirst = '" & Trim(txtText2.Text) & "'  "

    End If

    If Trim(txt3_cbo.Text) <> "" Then
        strSql = strSql + " AND LEAD_FREE  = '" & Trim(txt3_cbo.Text) & "'  "

    End If
    
    If Trim(txtText4.Text) <> "" Then
       strSql = strSql + " AND ECAT  = '" & Trim(txtText4.Text) & "'  "
    
    End If
    
    If Trim(txtText5.Text) <> "" Then
        strSql = strSql + " AND MSL  = '" & Trim(txtText5.Text) & "'  "

    End If

    If Trim(txt7_Cbo.Text) <> "" Then
        strSql = strSql + " AND TEMP   = '" & Trim(txt7_Cbo.Text) & "'  "

    End If
    
    If Trim(txtText6.Text) <> "" Then
        strSql = strSql + " AND HALIDE_FREE   = '" & Trim(txtText6.Text) & "'  "

    End If
    
    If Trim(txt8_Cbo.Text) <> "" Then
        strSql = strSql + " AND PBF_DIE_ATTACH   = '" & Trim(txt8_Cbo.Text) & "'  "

    End If
    
    If Trim(txtText9.Text) <> "" Then
        strSql = strSql + " AND MPQ_QTY   = '" & Trim(txtText9.Text) & "'  "

    End If

    If Trim(txtText10.Text) <> "" Then
        strSql = strSql + " AND PACKAGING_TYPE   = '" & Trim(txtText10.Text) & "'  "

    End If

    If Trim(txtText11.Text) <> "" Then
        strSql = strSql + " AND PKG_GRP_CD   = '" & Trim(txtText11.Text) & "'  "

    End If
    
    If Trim(txt12_Cbo.Text) <> "" Then
        strSql = strSql + " AND UL_LISTED_FLAG   = '" & Trim(txt12_Cbo.Text) & "'  "

    End If
 
    strSql = strSql + " order by id"
    
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If

    If Cnn.State = 0 Then
    ConOracle
    End If
        
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
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
    txtText5.Text = ""
    txtText7.Text = ""
    txtText6.Text = ""
    txtText8.Text = ""
    txtText9.Text = ""
    txtText10.Text = ""
    txtText11.Text = ""
    txtText12.Text = ""
End Function
Private Function query2()
    Dim strSql       As String

    Dim rs           As New ADODB.Recordset

    strSql = "select ''as ""选择"", id, LOC,PART,MarkingCodeFirst,LEAD_FREE,ECAT,MSL,TEMP,HALIDE_FREE,PBF_DIE_ATTACH,MPQ_QTY," & _
    "PACKAGING_TYPE,PKG_GRP_CD,UL_LISTED_FLAG  from  CUSTOMERMPNAttributes where flag='Y' order by ID"
  
      If Cnn.State = 0 Then
    ConOracle
    End If
        
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
   
    If Not rs.EOF Then
        Call ListDataType(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
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

Private Sub CmdOutReport_Click()  '导出
  Dim TEMP As String
  
  TEMP = " select id, LOC,PART,MarkingCodeFirst,LEAD_FREE,ECAT,MSL,TEMP,HALIDE_FREE,PBF_DIE_ATTACH,MPQ_QTY," & _
  "PACKAGING_TYPE,PKG_GRP_CD,UL_LISTED_FLAG  from  CUSTOMERMPNAttributes where flag='Y' order by id "
   
 ExporToExcel (TEMP)
 
End Sub

Private Function report_1()

    Dim strSql As String
    Dim PART         As String
    Dim MARKINGCODEFIRST      As String
    Dim LEAD_FREE          As String
    Dim ECAT        As String
    Dim MSL      As String
    Dim TEMP       As String
    Dim HALIDE_FREE    As String
    Dim PBF_DIE_ATTACH   As String
    Dim MPQ_QTY         As String
    Dim PACKAGING_TYPE  As String
    Dim PKG_GRP_CD     As String
    Dim UL_LISTED_FLAG    As String
  PART = Trim(txtText1.Text)
    MARKINGCODEFIRST = Trim(txtText2.Text)    'FrontSideMarking
    LEAD_FREE = Trim(txt3_cbo.Text)
    ECAT = Trim(txtText4.Text)
    MSL = Trim(txtText5.Text)
    TEMP = Trim(txtText6.Text)
    HALIDE_FREE = Trim(txt7_Cbo.Text)
    PBF_DIE_ATTACH = Trim(txt8_Cbo.Text)
    MPQ_QTY = Trim(txtText9.Text)
    PACKAGING_TYPE = Trim(txtText10.Text)
    PKG_GRP_CD = Trim(txtText11.Text)
    UL_LISTED_FLAG = Trim(txt12_Cbo.Text)
    
    strSql = "select ''as ""选择"", id, LOC,PART,MarkingCodeFirst,LEAD_FREE,ECAT,MSL,TEMP,HALIDE_FREE,PBF_DIE_ATTACH,MPQ_QTY," & _
  "PACKAGING_TYPE,PKG_GRP_CD,UL_LISTED_FLAG  from  CUSTOMERMPNAttributes where flag='Y'"
  
     If Trim(txtText1.Text) <> "" Then
        strSql = strSql + " AND PART  = '" & Trim(txtText1.Text) & "'  "

    End If

    If Trim(txtText2.Text) <> "" Then
        strSql = strSql + " AND MarkingCodeFirst = '" & Trim(txtText2.Text) & "'  "

    End If

    If Trim(txt3_cbo.Text) <> "" Then
        strSql = strSql + " AND LEAD_FREE  = '" & Trim(txt3_cbo.Text) & "'  "

    End If
    
    If Trim(txtText4.Text) <> "" Then
       strSql = strSql + " AND ECAT  = '" & Trim(txtText4.Text) & "'  "
    
    End If
    
    If Trim(txtText5.Text) <> "" Then
        strSql = strSql + " AND MSL  = '" & Trim(txtText5.Text) & "'  "

    End If

    If Trim(txt7_Cbo.Text) <> "" Then
        strSql = strSql + " AND TEMP   = '" & Trim(txt7_Cbo.Text) & "'  "

    End If
    
    If Trim(txtText6.Text) <> "" Then
        strSql = strSql + " AND HALIDE_FREE   = '" & Trim(txtText6.Text) & "'  "

    End If
    
    If Trim(txt8_Cbo.Text) <> "" Then
        strSql = strSql + " AND PBF_DIE_ATTACH   = '" & Trim(txt8_Cbo.Text) & "'  "

    End If
    
    If Trim(txtText9.Text) <> "" Then
        strSql = strSql + " AND MPQ_QTY   = '" & Trim(txtText9.Text) & "'  "

    End If

    If Trim(txtText10.Text) <> "" Then
        strSql = strSql + " AND PACKAGING_TYPE   = '" & Trim(txtText10.Text) & "'  "

    End If

    If Trim(txtText11.Text) <> "" Then
        strSql = strSql + " AND PKG_GRP_CD   = '" & Trim(txtText11.Text) & "'  "

    End If
    
    If Trim(txt12_Cbo.Text) <> "" Then
        strSql = strSql + " AND UL_LISTED_FLAG   = '" & Trim(txt12_Cbo.Text) & "'  "

    End If
 
 
    strSql = strSql + " order by id"
  
    ExporToExcel (strSql)

End Function
