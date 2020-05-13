VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_WOMOD 
   Caption         =   "PMC工单维护"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   405
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9855
   ScaleWidth      =   15675
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   1270
      ButtonWidth     =   1455
      ButtonHeight    =   1217
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "工单查询"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "工单删除"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "工单还原"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "工单重抛"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   9975
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   15735
      Begin VB.ComboBox cbTpe 
         Height          =   315
         ItemData        =   "Frm_WOMOD.frx":0000
         Left            =   1080
         List            =   "Frm_WOMOD.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1530
         Width           =   3975
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10560
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   22
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WOMOD.frx":00D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WOMOD.frx":091B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WOMOD.frx":12D5
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WOMOD.frx":1F27
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WOMOD.frx":2B79
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtID 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   1020
         Width           =   3975
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   7575
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   15255
         _Version        =   524288
         _ExtentX        =   26908
         _ExtentY        =   13361
         _StockProps     =   64
         DAutoCellTypes  =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   0
         MaxRows         =   0
         SpreadDesigner  =   "Frm_WOMOD.frx":3533
         AppearanceStyle =   0
      End
      Begin MSComDlg.CommonDialog com2 
         Left            =   6840
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单表"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   720
      End
      Begin MSForms.CheckBox chk 
         Height          =   495
         Left            =   5280
         TabIndex        =   3
         Top             =   930
         Width           =   1455
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "2566;873"
         Value           =   "0"
         Caption         =   "批量删除"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm_WOMOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

    Case 1  ' 查询
        OnQuery
    Case 3  ' 删除
        OnDel
    Case 5  ' 还原
    Case 7  ' 重抛
    Case 9  ' 退出
        Unload Me
End Select

End Sub

Private Sub OnQuery()

    Dim strID As String

    Dim rs    As ADODB.Recordset, rs1 As ADODB.Recordset

    Screen.MousePointer = 11

    strID = UCase(Trim(TxtID.Text))

    If Len(strID) = 0 Then
        MsgBox "请输入工单号", vbInformation, "提示"
        GoTo EXITTHIS

    End If

    If cbTpe.Text = "" Then
        MsgBox "请选择工单表", vbInformation, "提示"
        GoTo EXITTHIS

    End If
    
    Fps(0).MaxRows = 0
    
    Select Case cbTpe.ListIndex
    
        Case 0, 1, 2, 3, 4, 5, 6, 7
            Set rs = New ADODB.Recordset
            Set rs.ActiveConnection = OraConnect

        Case Else
            Set rs1 = New ADODB.Recordset
            Set rs1.ActiveConnection = SqlConnect
    
    End Select
    
    Select Case cbTpe.ListIndex
    
        Case 0
            rs.Source = "select * from shop_order where shop_order = '" & strID & "'"

        Case 1
            rs.Source = "select * from shop_order_detail where shop_order = '" & strID & "'"

        Case 2
            rs.Source = "select * from shop_order_property where shop_order = '" & strID & "'"
            
        Case 3
            rs.Source = "select * from ib_wohistory where ordername = '" & strID & "'"
            
        Case 4
            rs.Source = "select * from  ib_waferlist where ordername = '" & strID & "'"
            
        Case 5
            rs.Source = " select conn.* ,conn.rowid from container conn where conn.mfgorderid in  (select mfg.mfgorderid from mfgorder mfg where mfg.mfgordername  in ('" & strID & "'))"

        Case 6
            rs.Source = "select mfg.* , mfg.rowid from mfgorder mfg where mfg.mfgordername in ('" & strID & "')"
            
        Case 7
            rs.Source = " select * from A_Lotwafers al where al.workordername in ('" & strID & "')"

        Case 8
            rs1.Source = " select * from [erpbase].[dbo].[tblllplan] where 工单号 in ('" & strID & "')"
            
        Case 9
            rs1.Source = "select * from ERPBASE..TblERPFLToME where shop_order in ('" & strID & "')"
        
        Case 10
            rs1.Source = " select * from [erpdata].[dbo].[tblTSVworkorder] where ORDERNAME in  ('" & strID & "')"

        Case 11
            rs1.Source = " select * from [erpdata].[dbo].[tblTSVwaferlist] where ORDERNAME in ('" & strID & "') "

    End Select
    
    Select Case cbTpe.ListIndex
    
        Case 0, 1, 2, 3, 4, 5, 6, 7
            rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

            If rs.RecordCount = 0 Then
                MsgBox "查询不到数据", vbInformation, "提示"
                GoTo EXITTHIS

            End If
            
            With Fps(0)
                .MaxRows = 0
                Set .DataSource = rs

            End With

        Case Else
            rs1.Open , , adOpenStatic, adLockReadOnly, adCmdText
            
            If rs1.RecordCount = 0 Then
                MsgBox "查询不到数据", vbInformation, "提示"
                GoTo EXITTHIS

            End If
            
            With Fps(0)
                .MaxRows = 0
                
                Set .DataSource = rs1

            End With
    
    End Select

EXITTHIS:

    Screen.MousePointer = 0

End Sub

Private Sub OnDel()

    Dim strID       As String

    Dim i           As Integer, J As Integer

    Dim strFilePath As String

    Dim VBExcel     As Excel.Application

    Dim xlBook      As Excel.Workbook

    Dim xlSheet     As Excel.Worksheet

    If chk.Value = 0 Then
        strID = Trim$(UCase$(TxtID.Text))

        If Len(strID) = 0 Then
            MsgBox "请输入工单号", vbInformation, "提示"
            Exit Sub

        End If

        Call OnDelID(strID)
    Else
        com2.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
        com2.ShowOpen
        
        strFilePath = Replace(com2.filename, Chr(0), ",")
        
        If strFilePath = "" Then
            Exit Sub
        End If
        
        Set VBExcel = CreateObject("excel.application")
        VBExcel.Visible = False
        Set xlBook = VBExcel.Workbooks.Open(strFilePath)
        Set xlSheet = xlBook.Worksheets(1)

        For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
            For J = 1 To 1

                If J <= 26 Then
                    strChar = UCase(Chr(96 + J))
                Else
                    strChar = "A" & UCase(Chr(96 + J - 26))

                End If

                tempVal = Replace(Trim(xlSheet.Range(strChar & i).Value), Chr(13) + Chr(10), "")

                Select Case strChar
      
                    Case "A"
                        strID = tempVal

                End Select
                
            Next J
            
            Call OnDelID(strID)
            
        Next i
        
    End If
    
    MsgBox "全部删除成功", vbInformation, "提示"

End Sub

Private Sub OnDelID(OrderID As String)

  Dim Str_Sql       As String

    Dim STr_Sql1      As String

    Dim str_sql2      As String

    Dim Str_sql3      As String

    Dim STr_sql4      As String

    Dim STr_sql5      As String

    Dim str_sql6      As String

    Dim Str_sql7      As String

    Dim str_sql8      As String

    Dim str_sql9      As String

    Dim sty_sql10     As String

    Dim sty_sql11     As String

    Dim sty_sql12     As String

    Dim iRes          As Integer

    Dim Str_sql_Guard As String

    ' 加判断后再删除
    ' 0 是否退料
   Str_sql_Guard = "select SUM(实领数量) from [erpbase].[dbo].[tblllplan] where 工单号 =  '" + OrderID + "'"
    If Get_SqlserverNo(Str_sql_Guard) > 0 Then
    
        iRes = MsgBox("该工单还未全部退料,还要删除吗?", vbYesNoCancel, "提示:")
        If iRes <> vbYes Then

            Exit Sub

        End If
    End If
      
    ' 1 是否抛到金蝶
    Str_sql_Guard = "select * from erpdata..tblTSV_TLInfo a where a.工单号 = '" + OrderID + "'"

    If QuerySqlserver(Str_sql_Guard) Then
        iRes = MsgBox("工单已经抛到金蝶, 要继续删除吗?", vbYesNoCancel, "提示:")

        If iRes <> vbYes Then

            Exit Sub

        End If
    End If

    ' 2 产品是否在机台内
    Str_sql_Guard = "select a.RESOURCENAME from historymainline a,(select max(CONTAINERTXNSEQUENCE) mm, containername from historymainline " & "where containername in ( select conn.containername from container conn where conn.mfgorderid in  (select mfg.mfgorderid from mfgorder mfg where mfg.mfgordername  in ('" + OrderID + "')) ) " & "group by containername) b where a.containername = b.containername and a.CONTAINERTXNSEQUENCE = b.mm and a.RESOURCENAME is not null order by a.RESOURCENAME"

    If QueryStr(Str_sql_Guard) Then
        iRes = MsgBox("在机台内, 要继续删除吗?", vbYesNoCancel, "提示:")

        If iRes <> vbYes Then

            Exit Sub

        End If
    End If

    ' 3 产品是否在生产
    Str_sql_Guard = "select * from mfgorder a, a_lotwafers b, mappingdatatest c,customeroitbl_test d,ib_wohistory e,container f," & "currentstatus g,spec h,operation i,workcenter j, specbase k,container l,product m, productbase n " & "Where b.workordername = a.mfgordername and c.substrateid = b.waferscribenumber and to_char(d.id) = c.filename and e.ordername = b.workordername " & "and f.containerid = b.containerid and g.currentstatusid = f.currentstatusid and h.specid = g.specid and i.operationid = h.operationid " & "and j.workcenterid = i.workcenterid and k.specbaseid = h.specbaseid and a.mfgordername = '" + OrderID + "' and l.containerid = b.containerid " & "and l.status = 1 and m.productid = l.productid and n.productbaseid = m.productbaseid and k.specname <> '3010' "

    If QueryStr(Str_sql_Guard) Then
        iRes = MsgBox("在生产, 要继续删除吗?", vbYesNoCancel, "提示:")

        If iRes <> vbYes Then

            Exit Sub

        End If
    End If

    ' 备份数据
    STr_Sql1 = "insert into container_bak select * from container conn where conn.mfgorderid in  (select mfg.mfgorderid from mfgorder mfg where mfg.mfgordername  in ('" + OrderID + "'))"
    str_sql2 = "insert into mfgorder_bak select * from mfgorder mfg where mfg.mfgordername in ('" + OrderID + "') "
    Str_sql3 = "insert into A_Lotwafers_bak select * from A_Lotwafers al where al.workordername in ('" + OrderID + "')"
    STr_sql4 = "insert into ib_wohistory_bak select * from ib_wohistory where ordername in ('" + OrderID + "') "
    STr_sql5 = "insert into ib_waferlist_bak select * from ib_waferlist where ordername in ('" + OrderID + "') "
    str_sql6 = "insert into [erpdata].[dbo].[tblTSVworkorder_bak] select * from  [erpdata].[dbo].[tblTSVworkorder] where ORDERNAME in ('" + OrderID + "') "
    Str_sql7 = "insert into [erpdata].[dbo].[tblTSVwaferlist_bak] select * from  [erpdata].[dbo].[tblTSVwaferlist] where ORDERNAME in ('" + OrderID + "')"
    str_sql8 = "insert into [erpbase].[dbo].[tblllplan_bak] select * from [erpbase].[dbo].[tblllplan] where 工单号 in ('" + OrderID + "')"
    str_sql9 = "insert into PJ_WO_PRI_bak select * from PJ_WO_PRI where wo in ('" & OrderID & "')"

    AddSql (STr_Sql1)
    AddSql (str_sql2)
    AddSql (Str_sql3)
    AddSql (STr_sql4)
    AddSql (STr_sql5)
    AddSql (str_sql9)

    AddSql2 (str_sql6)
    AddSql2 (Str_sql7)
    AddSql2 (str_sql8)

  '  MsgBox "备份成功", vbInformation, "提示"

    ' 删除
    STr_Sql1 = "delete from container conn where conn.mfgorderid in  (select mfg.mfgorderid from mfgorder mfg where mfg.mfgordername  in ('" + OrderID + "')) "
    str_sql2 = "delete from mfgorder mfg where mfg.mfgordername in ('" + OrderID + "')"
    Str_sql3 = "delete from A_Lotwafers al where al.workordername in ('" + OrderID + "')"
    STr_sql4 = "delete from ib_wohistory where ordername in ('" + OrderID + "')"
    STr_sql5 = "delete from ib_waferlist where ordername in ('" + OrderID + "')"

    str_sql6 = "delete from  [erpdata].[dbo].[tblTSVworkorder] where ORDERNAME in ('" + OrderID + "') "
    Str_sql7 = "delete from  [erpdata].[dbo].[tblTSVwaferlist] where ORDERNAME in ('" + OrderID + "')"
    str_sql8 = "delete from  [erpbase].[dbo].[tblllplan] where 工单号 in ('" + OrderID + "')"
    str_sql9 = "delete from PJ_WO_PRI where wo in ('" & OrderID & "')"
    AddSql2 ("delete from erpdata..shop_order where shop_order = '" & OrderID & "' ")
    

    AddSql (STr_Sql1)
    AddSql (str_sql2)
    AddSql (Str_sql3)
    AddSql (STr_sql4)
    AddSql (STr_sql5)

    AddSql2 (str_sql6)
    AddSql2 (Str_sql7)
    AddSql2 (str_sql8)
    AddSql (str_sql9)
    
  '  MsgBox "删除成功", vbInformation, "提示"
End Sub
