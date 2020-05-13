VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_Gcversion_shipreport 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GC出货二级代码维护(出货资料)"
   ClientHeight    =   10290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16815
   LinkTopic       =   "Form2"
   ScaleHeight     =   10290
   ScaleWidth      =   16815
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Txt_Id 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Txt_rule_Bin 
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox Txt_htdevice 
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Txt_GrossDie 
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Txt_rule2 
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Cmd_Output 
      Caption         =   "导出"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox TxtCustpn 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Cmd_Delete 
      Caption         =   "删除"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton CMD_Modify 
      BackColor       =   &H0000FFFF&
      Caption         =   "修改"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Cmd_Query 
      Caption         =   "查询"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Cmd_Insert 
      BackColor       =   &H0000FF00&
      Caption         =   "新增"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   855
   End
   Begin FPSpreadADO.fpSpread fpS 
      Height          =   10095
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   12135
      _Version        =   524288
      _ExtentX        =   21405
      _ExtentY        =   17806
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
      MaxCols         =   6
      MaxRows         =   0
      SpreadDesigner  =   "Frm_Gcversion_Shipreport.frx":0000
      AppearanceStyle =   0
   End
   Begin VB.ComboBox CobType 
      Height          =   300
      ItemData        =   "Frm_Gcversion_Shipreport.frx":0422
      Left            =   1920
      List            =   "Frm_Gcversion_Shipreport.frx":0432
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox TxtRule 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox TxtPN 
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "ID"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "分bin二级代码第三位"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "厂内机种"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Gross Die"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "二级代码第二位"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "客户机种"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "形式"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "二级代码第三位"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "成品料号"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Frm_Gcversion_shipreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Enum E_GCREV

    E_CHOOSE = 1
    E_PN     '料号
    E_qtechPTNo    '厂内机种
    E_GCREV      '二级代码
    E_GCREV_B      '二级代码
    E_CUSTPN       '客户机种
    E_Type     '形式
    E_GrossDie     'GrossDie
    E_GCREV2     '二级代码第二码
    e_ID     'A/B
    E_END

End Enum






Private Sub Cmd_Delete_Click()
Dim DelPn As String
Dim i As Integer
Dim strPN As String
Dim strtype As String
Dim intID As Integer

  With fps
        For i = 1 To .MaxRows
            .Row = i
            .Col = E_GCREV.E_CHOOSE

            If .text <> "" Then
                If .text = 1 Then
            
                    .Col = E_GCREV.E_PN      '料号
                    
                    If DelMaterial = "" Then
                        DelPn = Trim(.text)
                    Else
                        DelPn = DelPn & "," & Trim(.text)
                    End If
                    DelCnt = DelCnt + 1
                End If

            End If
        Next i
        If MsgBox("你确认要删除" & DelPn & ",共" & DelCnt & "笔记录吗?", vbOKCancel, "提示") = vbCancel Then
            Exit Sub

        End If
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = E_GCREV.E_CHOOSE

            If .text <> "" Then
                If .text = 1 Then
            
                    .Col = E_GCREV.E_PN      '料号
                    strPN = Trim(.text)
                
                    .Col = E_GCREV.E_Type   '形式
                    strtype = Trim$(.text)
                     
                    .Col = E_GCREV.e_ID
                    intID = .text
                    
                    strsql = "insert into  Erptemp..GcCode_Reference_history  select '删除','" & gUserName & "', sysdatetime() , 事业部, 尺寸, 客户机种名, 厂内机种名, 成品料号, 制程, 二级代码, 分bin二级代码, GrossDie, 二级代码第二位 from erptemp..GcCode_Reference  where ID=" & intID
                
                    AddSql2 (strsql)
                  
                    strsql = "delete from  erptemp..GcCode_Reference  where id=" & intID

                    AddSql2 (strsql)

                    
                End If

            End If

        Next i

    End With
    updatetogcrev
    cmd_query_Click '查询
        
        
End Sub

Private Sub Cmd_Insert_Click()
   

    Dim SMR        As New ADODB.Recordset
    Dim strsql     As String
    Dim strVerson     As String
    Dim strqtechPTNo     As String
    Dim intID As Integer
    Dim strGcrev_B     As String

    
    If Trim(TxtPN.text) = "" Then
        MsgBox "请输入成品料号", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(TxtCustpn.text) = "" Then
        MsgBox "请输入客户机种", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(TxtRule.text) = "" Then
        MsgBox "请输入二级代码第三位", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(CobType.text) = "" Then
        MsgBox "请选择形式", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(Txt_GrossDie.text) = "" Then
        MsgBox "GrossDie不可为空", vbInformation, "提示"
        Exit Sub
    End If
    If IsNumeric(Txt_GrossDie.text) = False Then
        MsgBox "GrossDie必须为数字", vbInformation, "提示"
        Exit Sub
    End If
    
    If UCase(Trim(CobType.text)) = "WLT" Then
        strVerson = "B"
    Else
        strVerson = "A"
    End If
    If Txt_htdevice.text = "" Then
        strsql = "select  QTECHPTNO FROM erptemp..TBLTSVNPIPRODUCT  WHERE QTECHPTNO2='" & Trim(TxtPN.text) & "'"
        strqtechPTNo = GetSqlServerStr(strsql)
    Else
        strqtechPTNo = Trim(Txt_htdevice.text)
    End If
    strGcrev_B = Trim$(Txt_rule_Bin.text)
    
    'wla,normal同时传,二级代码，A,B
    'wlt传分bin二级代码，B
    '转normal传二级代码,A
    If UCase(Trim(CobType.text)) = "WLA" Or UCase(Trim(CobType.text)) = "NORMAL" Then
     '   If Get_SqlserverCnt(" SELECT * from  erptemp..GcCode_Reference  where  客户机种名='" & Trim(TxtCustpn.text) & "' and 成品料号='" & Trim(TxtPN.text) & "'") > 0 Then
    '        MsgBox Trim(TxtPN.text) & "已存在", vbinfomation, "提示"
    '        Exit Sub
    '    End If
    ElseIf UCase(Trim(CobType.text)) = "WLT" Then
    '    If Get_SqlserverCnt(" SELECT * from  erptemp..GcCode_Reference  where  客户机种名='" & Trim(TxtCustpn.text) & "' and 成品料号='" & Trim(TxtPN.text) & "' and 制程<>'转NORMAL'") > 0 Then
    '        MsgBox strPN & "已存在", vbinfomation, "提示"
   '         Exit Sub
   '     End If
    ElseIf UCase(Trim(CobType.text)) = "转NORMAL" Then
    '    If Get_SqlserverCnt(" SELECT * from  erptemp..GcCode_Reference  where  客户机种名='" & Trim(TxtCustpn.text) & "' and 成品料号='" & Trim(TxtPN.text) & "' and 制程<>'WLT'") > 0 Then
    '        MsgBox strPN & "已存在", vbinfomation, "提示"
   '        Exit Sub
    '    End If
        
    End If
    
    intID = GetSqlServerStr("select max(id)+1 from  erptemp..GcCode_Reference")
    strsql = "Insert into erptemp..GcCode_Reference (客户机种名,厂内机种名,成品料号,二级代码,分bin二级代码,制程,creat_by,creat_date, GrossDie, 二级代码第二位,ID ) values('" & Trim(TxtCustpn.text) & "','" & strqtechPTNo & "','" & Trim(TxtPN.text) & "','" & Trim(TxtRule.text) & "','" & strGcrev_B & "','" & Trim(CobType.text) & "','" & gUserName & "', sysdatetime()," & Trim(Txt_GrossDie.text) & ",'" & Trim(Txt_rule2.text) & "'," & intID & " )"

    AddSql2 (strsql)
    
    updatetogcrev
    cmd_query_Click


End Sub

Private Sub cmd_Modify_Click()
Dim DelPn As String
Dim i As Integer
Dim strPN As String
Dim strqtechPTNo As String
Dim intID As Integer
Dim strtype As String
Dim strCustPN As String '形式
Dim strGcrev As String '二级代码第三位不分bin
Dim strGcrev_B As String '二级代码第三位分bin
Dim strGcrev2 As String '二级代码第二位
Dim strGrossdie As String 'grossdie

    If Txt_Id.text = "" Then
        MsgBox "ID栏位为空，请正确操作", vbInformation, "提示"
        Exit Sub
        
    End If
    intID = Txt_Id.text
    strPN = Trim(TxtPN.text)
    strqtechPTNo = Trim(Txt_htdevice.text)
    strtype = Trim$(CobType.text)
    strCustPN = Trim$(TxtCustpn.text) '客户机种
    strGcrev = Trim$(TxtRule.text)
    strGcrev_B = Trim$(Txt_rule_Bin.text)
    strGcrev2 = Trim$(Txt_rule2.text)
    strGrossdie = Trim$(Txt_GrossDie.text)
    If strGrossdie = "" Then
        strGrossdie = 0
    End If
    
    If strtype = "WLA" Or strtype = "NORMAL" Then
        If Get_SqlserverCnt(" SELECT DISTINCT 二级代码,分bin二级代码, GrossDie from  erptemp..GcCode_Reference  where  客户机种名='" & strCustPN & "' and  成品料号='" & strPN & "'") > 1 Then
            MsgBox "修改失败," & strPN & "已存在", vbinfomation, "提示"
            Exit Sub
        End If
    End If
    strsql = "insert into  Erptemp..GcCode_Reference_history  select '更改前','" & gUserName & "', sysdatetime() , 事业部, 尺寸, 客户机种名, 厂内机种名, 成品料号, 制程, 二级代码, 分bin二级代码, GrossDie, 二级代码第二位 from erptemp..GcCode_Reference  where ID=" & intID

    AddSql2 (strsql)
    
    strsql = "update    erptemp..GcCode_Reference   set 厂内机种名='" & strqtechPTNo & "', 成品料号='" & strPN & "', 制程='" & strtype & " ', 客户机种名 ='" & strCustPN & "', 二级代码='" & strGcrev & "', 分bin二级代码='" & strGcrev_B & "',update_by='" & gUserName & "', 二级代码第二位 = '" & strGcrev2 & "', GrossDie=" & strGrossdie & ", update_date=sysdatetime() where ID=" & intID
 
    AddSql2 (strsql)
    updatetogcrev
    cmd_query_Click '查询
End Sub

Private Sub Cmd_Output_Click()

    Dim strsql     As String
    
    strsql = "SELECT 客户机种名, 厂内机种名, 成品料号, 制程, 二级代码, 分bin二级代码, creat_by,creat_date, update_by , update_date, GrossDie, 二级代码第二位, ID FROM erptemp..GcCode_Reference order by 成品料号 "
    
    SqlServerExporToExcel (strsql)
    
End Sub

Private Sub cmd_query_Click()
    Dim SMR        As New ADODB.Recordset
    Dim strsql     As String


    
    If SMR.State = adStateOpen Then SMR.Close

    
    strsql = "SELECT 0 AS 选择,成品料号  AS 料号 ,厂内机种名 AS 厂内机种 ,二级代码 AS 二级代码 ,分bin二级代码 AS 分bin二级代码  , 客户机种名 AS 客户机种 ,制程 as 形式, GrossDie, 二级代码第二位,ID from erptemp..GcCode_Reference where 1=1"
    If Trim(TxtPN.text) <> "" Then
        strsql = strsql & " and  成品料号='" & Trim(TxtPN.text) & "'"
    End If
    If Trim(TxtCustpn.text) <> "" Then
        strsql = strsql & " and 客户机种名='" & Trim(TxtCustpn.text) & "'"
    End If
    strsql = strsql & "  order by 成品料号 "
    SMR.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        With fps
           .MaxRows = 0
           Set .DataSource = SMR
          
        End With
        
    Else
    
        With fps
           .MaxRows = 0

          
        End With

    End If
End Sub

Private Sub Form_Load()
    inictrl
End Sub

Private Sub inictrl()

  
    
    'Fps初始化
    With fps
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .ColsFrozen = 2
        .ButtonDrawMode = 1
        .Row = -1
        .Col = 1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        '设定列类型
        .Col = E_GCREV.E_CHOOSE   '选择
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '设定列宽
        .ColWidth(-1) = 10
        .ColWidth(E_GCREV.E_CHOOSE) = 4
        .RowHeight(-1) = 10
        '设定是否排序
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ZOrder
        .ReDraw = True
    End With
    If gUserName <> "16642" And gUserName <> "07885" Then
        Cmd_Insert.Enabled = False
        CMD_Modify.Enabled = False
        Cmd_Delete.Enabled = False
    End If
    

End Sub

    
    
    

Sub updatetogcrev()
    Dim SMR        As New ADODB.Recordset
    Dim strsql     As String
    Dim i As Integer
    Dim strtype As String
    Dim strrev_a As String
    Dim strrev_b As String
    Dim strqtechPTNo As String
    Dim strPN As String


    AddSql2 ("DELETE FROM erpdata..gcrev ")
    If SMR.State = adStateOpen Then SMR.Close
    strsql = "select  DISTINCT 成品料号,isnull(厂内机种名,'') as 厂内机种名 , 二级代码,isnull(分bin二级代码,'') as 分bin二级代码,制程 from erptemp..GcCode_Reference"
    SMR.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            strtype = UCase(Trim(SMR("制程")))
            strrev_a = Trim(SMR("二级代码"))
            strrev_b = Trim(SMR("分bin二级代码"))
            strqtechPTNo = Trim(SMR("厂内机种名"))
            strPN = Trim(SMR("成品料号"))
            If strrev_b = "" Then
                strrev_b = strrev_a
            End If
                
            If strtype = "WLT" Then
                strsql = "insert into erpdata..gcrev(product , DEVICE, [rule],[Version]) values('" & strPN & "','" & strqtechPTNo & "','" & strrev_b & "','B')"
                AddSql2 (strsql)
                
            ElseIf strtype = "转NORMAL" Then
                strsql = "insert into erpdata..gcrev(product , DEVICE, [rule], [Version]) values('" & strPN & "','" & strqtechPTNo & "','" & strrev_a & "','A')"
                AddSql2 (strsql)
          '  ElseIf strtype = "WLA" Or strtype = "NORMAL" Then
           Else
                strsql = "insert into erpdata..gcrev(product , DEVICE, [rule], [Version]) values('" & strPN & "','" & strqtechPTNo & "','" & strrev_a & "','A')"
                strsql = strsql & ";" & "insert into erpdata..gcrev(product , DEVICE, [rule], [Version]) values('" & strPN & "','" & strqtechPTNo & "','" & strrev_b & "','B')"
                AddSql2 (strsql)

            End If
            
            
            SMR.MoveNext
            
        Next
    End If

    
End Sub






Private Sub fps_Click(ByVal Col As Long, ByVal Row As Long)
Dim J As Integer


If Col <> 1 Then Exit Sub
With fps
    .Col = 1
    .Row = Row
    .Value = Abs(Val(.Value) - 1)

    If Val(.Value) = 1 Then
    
        For J = 1 To .MaxRows
            If J <> Row Then
                .Row = J
                .Col = 1
                .Value = 0

                .Col = -1
                .BackColor = &H8000000F
            End If
            
        Next
        .Row = Row
        .Col = -1
        .BackColor = &HC0C0FF
       
         .Col = E_GCREV.E_PN      '料号
         TxtPN.text = Trim(.text)
         .Col = E_GCREV.E_qtechPTNo      '厂内机种
         Txt_htdevice.text = Trim(.text)
         .Col = E_GCREV.E_GCREV      '二级代码
         TxtRule.text = Trim(.text)
         .Col = E_GCREV.E_GCREV_B      '分bin二级代码
         Txt_rule_Bin.text = Trim(.text)
         .Col = E_GCREV.E_CUSTPN       '客户机种
         TxtCustpn.text = Trim(.text)
         .Col = E_GCREV.E_Type     '形式
         CobType.text = Trim(.text)
         .Col = E_GCREV.E_GrossDie     'GrossDie
         Txt_GrossDie.text = Trim(.text)
         .Col = E_GCREV.E_GCREV2     '二级代码第二码
         Txt_rule2.text = Trim(.text)
          .Col = E_GCREV.e_ID
          Txt_Id.text = .text 'ID
          
    Else
        TxtPN.text = ""
        Txt_htdevice.text = ""
        TxtCustpn.text = ""
        CobType.text = ""
        Txt_GrossDie.text = ""
        TxtRule.text = ""
        Txt_rule2.text = ""
        Txt_rule_Bin.text = ""
        Txt_Id.text = ""
        .Row = Row
        .Col = -1
        .BackColor = &H8000000F
        
        
    End If
End With
End Sub





