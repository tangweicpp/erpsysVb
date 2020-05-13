VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_XBCWH 
   Caption         =   "Form1"
   ClientHeight    =   12825
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21960
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
   ScaleHeight     =   12825
   ScaleWidth      =   21960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame FrmCJ 
      Caption         =   "车间调拨主页"
      Height          =   735
      Left            =   7560
      TabIndex        =   23
      Top             =   840
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdCJ3 
         Caption         =   "盘点确认"
         Height          =   480
         Left            =   4080
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCJ2 
         Caption         =   "盘点"
         Height          =   480
         Left            =   2040
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCJ1 
         Caption         =   "车间用户进入"
         Height          =   480
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "执行查询"
         Height          =   495
         Left            =   6600
         TabIndex        =   25
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "关闭"
         Height          =   495
         Left            =   8880
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6960
         TabIndex        =   26
         Top             =   1800
         Width           =   360
      End
   End
   Begin VB.Frame FraJHB 
      Caption         =   "计划部调拨主页"
      Height          =   1575
      Left            =   720
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdJHBBB 
         Caption         =   "计划部报表"
         Height          =   480
         Left            =   3000
         TabIndex        =   22
         Top             =   600
         Width           =   1110
      End
      Begin VB.CommandButton cmdJHBZY 
         Caption         =   "计划部主页"
         Height          =   480
         Left            =   840
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "关闭"
         Height          =   495
         Left            =   8880
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "执行查询"
         Height          =   495
         Left            =   6600
         TabIndex        =   18
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lbl7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6960
         TabIndex        =   20
         Top             =   1800
         Width           =   360
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton commandexcel 
      BackColor       =   &H0080FFFF&
      Caption         =   "导出当前页面"
      Height          =   375
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdQery 
      BackColor       =   &H0080FFFF&
      Caption         =   "查询"
      Height          =   360
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtsup_sn 
      Height          =   285
      Left            =   11040
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtPrd_id 
      Height          =   285
      Left            =   7920
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmddel 
      BackColor       =   &H000000FF&
      Caption         =   "删除"
      Height          =   600
      Left            =   19200
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00C0C000&
      Caption         =   "提交"
      Height          =   600
      Left            =   16560
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   12
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   184745985
      CurrentDate     =   43738
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   375
      Index           =   0
      Left            =   11040
      TabIndex        =   15
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   184745985
      CurrentDate     =   43738
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   11655
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   21735
      _Version        =   524288
      _ExtentX        =   38338
      _ExtentY        =   20558
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
      MaxCols         =   20
      MaxRows         =   0
      SpreadDesigner  =   "FrmXBCWH.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束时间"
      Height          =   195
      Left            =   10200
      TabIndex        =   14
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间"
      Height          =   195
      Left            =   7080
      TabIndex        =   13
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "不代表最终品质"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   20640
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "当前版本:"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   17640
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "V20417"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   18600
      TabIndex        =   9
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblPH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "批号"
      Height          =   195
      Left            =   10560
      TabIndex        =   6
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblLH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "料号"
      Height          =   195
      Left            =   7440
      TabIndex        =   5
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Frm_XBCWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'通用全局变量
Dim FLAG As String
Dim strSQlQJ As String

Private Sub cmdCJ2_Click() '车间盘点2
    FLAG = "车间盘点2"
    cmddel.Visible = False
    query3
End Sub

Private Sub cmdCJ1_Click() '车间盘点1
    FLAG = "车间盘点1"
    cmddel.Visible = False
    Query2
End Sub

Private Sub cmdCJ3_Click() '车间盘点3
    FLAG = "车间盘点3"
    cmddel.Visible = False
    query5
End Sub

Private Sub cmdDel_Click()
Select Case FLAG
      
    Case "计划部报表"
        cmddel_JHBBB

    Case Else
        MsgBox "出错啦！"
End Select

End Sub

Private Sub cmdJHBZY_Click() '计划部主页
    FLAG = "计划部主页"
    cmddel.Visible = False
    Query1
End Sub

Private Sub cmdJHBBB_Click() '计划部报表
    FLAG = "计划部报表"
    cmddel.Visible = True
    query4
End Sub

Private Function Query1()
    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    

    strSql = "select '' as '选择' ,a.Prd_ID as 料号,a.SupSN as 批号,c.物料名称,a.stockName as 当前所在仓,a.qty as 目前总量,(a.qty-isnull(b.qty,0)) as 可调拨量,a.unit as 单位,'' as '调拨数量','' as '转调仓库',case when a.Flag = '1' then '费领' else '调拨' end as '类型','' as '备注',A.Sissuetime as '时间' " & _
    "from (select Prd_ID, supsn ,stockName , sum(qty) as qty,unit, Flag,max(CreateDate) as 'Sissuetime' from erptemp..tblErp_ShopOrderIssue group by Prd_ID,supsn,stockName,unit,Flag) A " & _
    "left join (select Prd_ID, SupSN, fromstoragename,sum(IssueQty) as qty,unit from erptemp..tblErp_ShopOrderIssue_STORAGE group by Prd_ID,supsn,fromstoragename,unit) B " & _
    "on a.prd_id=b.prd_id and a.supsn=b.supsn and a.StockName = b.fromstoragename and a.unit=b.unit left join (select 物料名称,料号 from erpdata..tblSmainM2) C on a.Prd_ID = c.料号 " & _
    "where (a.qty-isnull(b.qty,0)) > 0 "

    
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2
    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType1(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType1(rs)
        rs.Close
        Exit Function
    End If
End Function

Private Function Query2()
    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    strSql = "select '' as '选择',a.Prd_ID as '料号',a.SupSN as '批号',d.物料名称,a.issuestoragename as '仓名',(isnull(a.IssueQty,0)-isnull(b.IssueQty,0)-isnull(c.InventoryQty,0)) as '总量',a.unit as '单位','' as '调拨数量','' as '转调仓库',a.Remark as '备注',a.Sissuetime as '时间' from " & _
    "(select Prd_ID,SupSN,issuestoragename,unit,SUM(IssueQty) as IssueQty,max(issuetime) as 'Sissuetime', Remark from erptemp..tblErp_ShopOrderIssue_PLANT group by Prd_ID,SupSN,issuestoragename,unit,Remark) A " & _
    "left join(select Prd_ID,SupSN,fromstoragename,SUM(IssueQty) as IssueQty,Remark from erptemp..tblErp_ShopOrderIssue_PLANT where fromstoragename='CIS一楼黄光区线边仓' or fromstoragename='CIS二楼黄光区线边仓' or fromstoragename='CIS二楼电镀区线边仓' or fromstoragename='Bumping线边仓' or fromstoragename='WLP线边仓' or fromstoragename='12寸TSV制造部线边仓' group by Prd_ID,SupSN,fromstoragename,Remark) B " & _
    "on a.Prd_ID=b.Prd_ID and a.SupSN=b.SupSN and a.issuestoragename=b.fromstoragename and a.Remark=b.Remark " & _
    "left join (select prd_id,supsn,storagename,sum(dailyusageqty) as InventoryQty from erptemp..tblErp_ShopOrderIssue_Inventory group by prd_id,supsn,storagename) C " & _
    "on a.prd_id=c.prd_id and a.SupSN=c.SupSN and a.issuestoragename=c.storagename " & _
    "left join (select 物料名称,料号 from erpdata..tblSmainM2) D on a.Prd_ID = d.料号  " & _
    "where (isnull(a.IssueQty,0)-isnull(b.IssueQty,0)-isnull(c.InventoryQty,0)) > 0 "
    
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
    strSQlQJ = strSql

    If Not rs.EOF Then
        Call ListDataType2(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType2(rs)
        rs.Close
        Exit Function

    End If
End Function

Private Function query3()
    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    strSql = "select '' as '选择',a.Prd_ID as '料号',a.SupSN as '批号',a.Prd_Name as '物料名称',a.issuestoragename as '仓名',(isnull(a.IssueQty,0)-isnull(b.IssueQty,0)-isnull(c.InventoryQty,0)) as '总量',a.unit as '单位','' as '盘点量','' as '耗用量','' as '备注',A.Sissuetime as '时间'" & _
    "from (select Prd_ID,SupSN,Prd_Name,issuestoragename,unit,SUM(IssueQty) as IssueQty,max(issuetime) as 'Sissuetime' from erptemp..tblErp_ShopOrderIssue_PLANT group by Prd_ID,SupSN,issuestoragename,unit,Prd_Name) A " & _
    "left join (select Prd_ID,SupSN,Prd_Name,fromstoragename,SUM(IssueQty) as IssueQty from erptemp..tblErp_ShopOrderIssue_PLANT where fromstoragename='CIS一楼黄光区线边仓' or fromstoragename='CIS二楼黄光区线边仓' or fromstoragename='CIS二楼电镀区线边仓' or fromstoragename='Bumping线边仓' or fromstoragename='WLP线边仓' or fromstoragename ='Bumping线边仓' or fromstoragename = '12寸TSV制造部线边仓' group by Prd_ID,SupSN,fromstoragename,Prd_Name) B " & _
    "on a.Prd_ID=b.Prd_ID and a.SupSN=b.SupSN and a.issuestoragename=b.fromstoragename and a.Prd_Name = b.Prd_Name " & _
    "left join  (select prd_id,supsn,Prd_Name,storagename,sum(dailyusageqty) as InventoryQty from erptemp..tblErp_ShopOrderIssue_Inventory group by prd_id,Prd_Name,supsn,storagename) C " & _
    "on a.prd_id=c.prd_id and a.SupSN=c.SupSN and a.issuestoragename=c.storagename and a.Prd_Name=c.Prd_Name " & _
    "where (isnull(a.IssueQty,0)-isnull(b.IssueQty,0)-isnull(c.InventoryQty,0)) > 0   order by a.Prd_ID asc"

    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2
    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    strSQlQJ = strSql
    If Not rs.EOF Then
        Call ListDataType3(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType3(rs)
        rs.Close
        Exit Function
    End If
End Function

Private Function query4()
    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    strSql = "select '' as '选择',Prd_ID as '料号',SupSN as '批号',Prd_Name as '物料名称', " & _
    "IssueQty as '单次调拨量', Unit as '单位',FromStorageName as '来源仓',IssueStorageName as '目的仓',IssueUser as '调拨人',IssueTime as '调拨时间', Remark as '备注' " & _
    "from erptemp..tblErp_ShopOrderIssue_STORAGE where IssueTime >= CONVERT(VARCHAR(10),GETDATE(),120) " & _
    "Order by IssueTime DESC "
    
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType4(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType4(rs)
        rs.Close
        Exit Function

    End If
End Function

Private Function query5()
'    Dim strSql       As String
'    Dim rs           As New ADODB.Recordset
'    strSql = "select a.Prd_ID,a.SupSN,a.Prd_Name,a.issuestoragename,(isnull(a.IssueQty,0)-isnull(b.IssueQty,0)-isnull(c.DailyUsageQty,0)) as cur_qty,a.unit,row_number() over(order by a.Prd_ID) as ID " & _
'    "from (select Prd_ID,SupSN,Prd_Name,issuestoragename,unit,SUM(IssueQty) as IssueQty from erptemp..tblErp_ShopOrderIssue_PLANT group by Prd_ID,SupSN,Prd_Name,issuestoragename,unit) A " & _
'    "left join (select Prd_ID,SupSN,Prd_Name,fromstoragename,SUM(IssueQty) as IssueQty from erptemp..tblErp_ShopOrderIssue_PLANT where fromstoragename='CIS一楼黄光区线边仓' or fromstoragename='CIS二楼黄光区线边仓' or fromstoragename='CIS二楼电镀区线边仓' or fromstoragename='Bumping线边仓' or fromstoragename='WLP线边仓' group by Prd_ID,SupSN,Prd_Name,fromstoragename) B " & _
'    "on a.Prd_ID=b.Prd_ID and a.SupSN=b.SupSN and a.issuestoragename=b.fromstoragename and a.Prd_Name=b.Prd_Name " & _
'    "left join (select Prd_ID,SupSN,Prd_Name,StorageName,SUM(DailyUsageQty) as DailyUsageQty from erptemp..tblErp_ShopOrderIssue_Inventory where StorageName='CIS一楼黄光区线边仓' or StorageName = 'CIS二楼黄光区线边仓'  or StorageName= 'CIS二楼电镀区线边仓' or StorageName='Bumping线边仓' or StorageName='WLP线边仓'  group by Prd_ID,SupSN,Prd_Name,StorageName) C " & _
'    "on a.Prd_ID=c.Prd_ID and a.SupSN=c.SupSN and a.Prd_Name=c.Prd_Name and a.issuestoragename = c.StorageName " & _
'    "where (isnull(a.IssueQty,0)-isnull(b.IssueQty,0)-isnull(c.DailyUsageQty,0)) > 0 "
'
'    If INIadoCon.State <> adStateOpen Then
'        INIConnectSTART2
'
'    End If
'    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not rs.EOF Then
'        Call ListDataType(rs)
'    Else
'        MsgBox "无数据", vbInformation, "提示"
'        Call ListDataType(rs)
'        rs.Close
'        Exit Function
'
'    End If
End Function

Private Sub ListDataType1(rs As ADODB.Recordset) '计划部主页
  Dim i As Long
   
    Dim j As Long
    
    With fps
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
        With fps

            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                .BackColor = &HFFFF&
                .CellType = CellTypeCheckBox
                .text = 0
                .Col = 2
                .Lock = True
                .Col = 3
                .Lock = True
                .Col = 4
                .Lock = True
                .Col = 5
                .Lock = True
                .Col = 6
                .Lock = True
                .Col = 7
                .Lock = True
                .Col = 8
                .Lock = True
                .Col = 9
                .Lock = False
                .Col = 10
                .CellType = CellTypeComboBox
                .ColWidth(-1) = 15
                .RowHeight(-1) = 12
                .TypeComboBoxList = .TypeComboBoxList & "CIS一楼黄光区线边仓"
    
                .TypeComboBoxList = .TypeComboBoxList & "CIS二楼黄光区线边仓"

                .TypeComboBoxList = .TypeComboBoxList & "CIS二楼电镀区线边仓"

                .TypeComboBoxList = .TypeComboBoxList & "Bumping线边仓"
    
                .TypeComboBoxList = .TypeComboBoxList & "WLP线边仓"

                .TypeComboBoxList = .TypeComboBoxList & "12寸TSV制造部线边仓"

                .LockBackColor = vbYellow
                .Col = 11
                .Lock = True
                .Col = 12
                .Lock = False
                
            Next

    End With
End Sub
Private Sub ListDataType2(rs As ADODB.Recordset) '车间盘点1
  Dim i As Long
   
    Dim j As Long
    
    With fps
        .ColWidth(-1) = 15
        .RowHeight(-1) = 12
        .MaxRows = 0

        Set .DataSource = rs
    End With
    With fps
            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                .BackColor = &HFFFF&
                .CellType = CellTypeCheckBox
                .text = 0
                .Col = 2
                .Lock = True
                .Col = 3
                .Lock = True
                .Col = 4
                .Lock = True
                .Col = 5
                .Lock = True
                .Col = 6
                .Lock = True
                .Col = 7
                .Lock = True
                .Col = 8
                .Lock = False
                .Col = 9
                .CellType = CellTypeComboBox
                .TypeComboBoxList = .TypeComboBoxList & "CIS一楼黄光区线边仓"
    
                .TypeComboBoxList = .TypeComboBoxList & "CIS二楼黄光区线边仓"

                .TypeComboBoxList = .TypeComboBoxList & "CIS二楼电镀区线边仓"

                .TypeComboBoxList = .TypeComboBoxList & "Bumping线边仓"
    
                .TypeComboBoxList = .TypeComboBoxList & "WLP线边仓"

                .TypeComboBoxList = .TypeComboBoxList & "12寸TSV制造部线边仓"

                .LockBackColor = vbYellow
                .Col = 10
                .Lock = True
                
            Next

    End With
End Sub
Private Sub ListDataType3(rs As ADODB.Recordset) '车间盘点3
    
   Dim i As Long
   
    With fps
        .ColWidth(-1) = 15
        .RowHeight(-1) = 12
        .MaxRows = 0

        Set .DataSource = rs
    End With
    With fps
            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                .BackColor = &HFFFF&
                .CellType = CellTypeCheckBox
                .text = 0
                .Col = 2
                .Lock = True
                .Col = 3
                .Lock = True
                .Col = 4
                .Lock = True
                .Col = 5
                .Lock = True
                .Col = 6
                .Lock = True
                .Col = 7
                .Lock = True
                .Col = 8
                .Lock = False
                .Col = 9
                .Lock = True
                .LockBackColor = vbYellow
                .Col = 10
                .Lock = True
                
            Next

    End With

End Sub

Private Sub ListDataType4(rs As ADODB.Recordset) '计划部报表
  Dim i As Long
   
    Dim j As Long
    
    With fps
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
        With fps

            For i = 1 To .MaxRows
                .Row = i
                 .Col = 1
                .BackColor = &HFFFF&
                .CellType = CellTypeCheckBox
                .text = 0
                .Col = 2
                .Lock = True
                .Col = 3
                .Lock = True
                .Col = 4
                .Lock = True
                .Col = 5
                .Lock = True
                .Col = 6
                .Lock = True
                .Col = 7
                .Lock = True
                .Col = 8
                .Lock = True
                .Col = 9
                .Lock = True
                .Col = 10
                .Lock = True
                .Col = 11
                .Lock = False
            Next

    End With
End Sub

Private Sub cmdQery_Click()
Select Case FLAG

    Case "车间盘点2"
        Query_CJPD2
        FLAG = "车间盘点2"
    Case "车间盘点1"
        Query_CJPD1
        FLAG = "车间盘点1"
    Case "车间盘点3"
        'Query_CJPD3
        'FLAG = "车间盘点3"
    Case "计划部主页"
        Query_JHBZY
        FLAG = "计划部主页"
    Case "计划部报表"
        Query_JHBBB
        FLAG = "计划部报表"
    Case Else
        MsgBox "出错啦！"
End Select
End Sub

Private Sub cmdSubmit_Click()
Select Case FLAG

    Case "车间盘点2"
        submit_CJPD2

    Case "车间盘点1"
       submit_CJPD1

    Case "车间盘点3"
        MsgBox "该功能未实现"
        
    Case "计划部主页"
        submit_JHBZY
        
    Case "计划部报表"
        MsgBox "该功能未实现"

    Case Else
        MsgBox "出错啦！"
End Select

End Sub

Private Sub commandexcel_Click()
    If strSQlQJ <> "" Then
        SqlServerExporToExcel (strSQlQJ)
    Else
        MsgBox "当前页面无数据，无需导出！"
    End If
End Sub

Private Sub Form_Load()
    Dim strsql1 As String
    Dim rs           As New ADODB.Recordset
    
    strsql1 = ""
    FLAG = ""
    strSQlQJ = ""
    FraJHB.Visible = False
    FrmCJ.Visible = False
    strsql1 = "SELECT * from erptemp..tblErp_ShopOrderIssue_Grant where UserNo = '" & gUserName & "'"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql1, INIadoCon, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
         For i = 0 To rs.RecordCount - 1
            If rs.Fields("UserNo") = gUserName And rs.Fields("MenuName") = "车间" Then
                FraJHB.Visible = False
                FrmCJ.Visible = True
            End If
            If rs.Fields("UserNo") = gUserName And rs.Fields("MenuName") <> "车间" Then
                FraJHB.Visible = True
                FrmCJ.Visible = False
            End If
            rs.MoveNext
        Next
    Else
        MsgBox "你没有权限使用该界面"
        cmdSubmit.Visible = False
        cmddel.Visible = False
        cmdQery.Visible = False
        commandexcel.Visible = False
    End If
        
    cmddel.Visible = False
    DTP1(1).Value = DATE
    DTP2(0).Value = DATE
    If gUserName = "07885" Or gUserName = "08652" Then
        FraJHB.Visible = True
        FrmCJ.Visible = True
    End If
    
End Sub

Private Function submit_JHBZY()
     
    Dim rs        As New ADODB.Recordset
    '修改
    Dim i         As Integer

    Dim strSql    As String
    Dim strsql1    As String
    Dim strSql2    As String
    
    Dim Prd_ID As String
    Dim SupSN As String
    Dim WLMC As String
    Dim stockName As String
    Dim QTY As String
    Dim KDBL As String
    Dim unit As String
    Dim DB_num As String
    Dim ZDCK As String
    Dim LX As String
    Dim REMARK As String
    Dim SYL As String
    
    Dim count As Integer
    
    count = 0
    
    With fps

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 2
                If Trim(.text) <> "" Then
                    Prd_ID = Trim(.text)
                End If
                
                .Col = 3
                If Trim(.text) <> "" Then
                    SupSN = Trim(.text)
                End If

                .Col = 4
                If Trim(.text) <> "" Then
                    WLMC = Trim(.text)
                End If

                .Col = 5
                If Trim(.text) <> "" Then   '当前所在仓
                    stockName = Trim(.text)
                End If
                .Col = 6
                If Trim(.text) <> "" Then
                    QTY = Trim(.text)       '当前总量
                End If
                
                .Col = 7
                If Trim(.text) <> "" Then   '可调拨量
                    KDBL = Trim(.text)
                End If
                
                .Col = 8
                If Trim(.text) <> "" Then
                    unit = Trim(.text)
                End If
                
                .Col = 9
                If Trim(.text) <> "" Then   '调拨量
                    DB_num = Trim(.text)
                End If
                
                .Col = 10
                If Trim(.text) <> "" Then
                    ZDCK = Trim(.text)
                End If
                
                .Col = 12
                If Trim(.text) <> "" Then
                    REMARK = Trim(.text)
                End If
                
                If Check5_input(DB_num) = 1 Then
                    MsgBox "lot参数输入不合法"
                    Exit Function
                End If
                    
                If DB_num = "" Then
                    MsgBox "调拨数量为必填且为数字！"
                ElseIf ZDCK = "" Then
                    MsgBox "转调仓库为必填！"
                ElseIf Val(KDBL) < Val(DB_num) Then
                    MsgBox "调拨量超过可调拨量！"
                Else
                SYL = QTY - DB_num
                
                strSql = "insert into  erptemp..tblErp_ShopOrderIssue_PLANT (Prd_ID,SupSN,Prd_Name,IssueQty,Unit,FromStorageName, " & _
                "IssueStorageName,IssueUser,Remark) values ('" & Prd_ID & "','" & SupSN & "','" & WLMC & "','" & DB_num & "','" & unit & "','" & stockName & "','" & ZDCK & "','" & gUserName & "','" & REMARK & "') "
                strsql1 = "insert into  erptemp..tblErp_ShopOrderIssue_STORAGE (Prd_ID,SupSN,Prd_Name,IssueQty,Unit,FromStorageName, " & _
                "IssueStorageName,IssueUser,Remark) values ('" & Prd_ID & "','" & SupSN & "','" & WLMC & "','" & DB_num & "','" & unit & "','" & stockName & "','" & ZDCK & "','" & gUserName & "','" & REMARK & "') "
'               strSql2 = "insert into erptemp..tblErp_ShopOrderIssue_Inventory (Prd_ID,SupSN,Prd_Name,StorageName,SystemQty, " & _
'                "Unit,ActureQty,DailyUsageQty,IssueUser,Remark) values ( '" & Prd_ID & "','" & SupSN & "','" & WLMC & "','" & stockName & "','" & QTY & "','" & unit & "','" & SYL & "','" & DB_num & "','" & gUserName & "','" & REMARK & "') "

                AddSql2 (strSql)
                AddSql2 (strsql1)
'               AddSql2 (strSql2)
                count = count + 1
                End If
            End If
        Next i

    End With

    If count = 0 Then
        MsgBox "调拨失败"
    Else
        MsgBox "调拨成功" & "调拨记录数" & count & "! "
    
    End If

Query1

End Function

Private Function Check5_input(input_String As String) As Integer
    If InStr(input_String, "'") > 0 Or InStr(input_String, "’") > 0 Then
       ' MsgBox "输入字符串不合法"
        Check5_input = 1
    Else
        Check5_input = 0
    End If
End Function

Private Function cmddel_JHBBB()
   
    Dim rs        As New ADODB.Recordset
    '修改
    Dim i         As Integer

    Dim strSql    As String
    Dim strsql1    As String
    Dim strSql2 As String
    
    Dim Prd_ID As String
    Dim IssueQty As String
    Dim issuestoragename As String
    Dim IssueTime As String
    
    Dim count As Integer
    
    count = 0
    
    With fps

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 2
                If Trim(.text) <> "" Then
                    Prd_ID = Trim(.text)
                End If
                
                .Col = 3
                If Trim(.text) <> "" Then
                    SupSN = Trim(.text)
                End If

                .Col = 4
                If Trim(.text) <> "" Then
                    prd_name = Trim(.text)
                End If
                
                .Col = 5
                If Trim(.text) <> "" Then
                    DailyUsageQty = Trim(.text)
                End If
                
                .Col = 6
                If Trim(.text) <> "" Then
                    unit = Trim(.text)
                End If

                .Col = 7
                If Trim(.text) <> "" Then
                    StorageName = Trim(.text)
                End If
                
                .Col = 8
                If Trim(.text) <> "" Then
                    issuestoragename = Trim(.text)
                End If

                .Col = 10
                If Trim(.text) <> "" Then   '当前所在仓
                    IssueTime = Trim(.text)
                End If
       
                strSql = "delete from erptemp..tblErp_ShopOrderIssue_STORAGE where Prd_ID = '" & Prd_ID & "' and IssueQty = '" & IssueQty & "' and IssueStorageName = '" & issuestoragename & "' and CONVERT(VARCHAR(10),IssueTime,120) = '" & Format(IssueTime, "yyyy-MM-dd") & "'"
                strsql1 = "delete from erptemp..tblErp_ShopOrderIssue_PLANT where Prd_ID = '" & Prd_ID & "' and IssueQty = '" & IssueQty & "' and IssueStorageName = '" & issuestoragename & "' and CONVERT(VARCHAR(10),IssueTime,120) = '" & Format(IssueTime, "yyyy-MM-dd") & "'"
                strSql2 = "delete from erptemp..tblErp_ShopOrderIssue_Inventory where Prd_ID = '" & Prd_ID & "' and SupSN = '" & SupSN & "' and Prd_Name = '" & _
                    prd_name & "' and DailyUsageQty = '" & DailyUsageQty & "' and unit = '" & unit & "' and StorageName = " & StorageName & "', and IssueStorageName = " & issuestoragename & "'and  CONVERT(VARCHAR(10),IssueTime,120) = '" & Format(IssueTime, "yyyy/MM/dd") & "'"

                AddSql2 (strSql)
                AddSql2 (strsql1)
                AddSql2 (strSql2)
                count = count + 1
                
            End If
        Next i

    End With

    If count = 0 Then
        MsgBox "调拨失败"
    Else
        MsgBox "调拨成功" & "调拨记录数" & count & "! "
    
    End If

query4
End Function

Private Function submit_CJPD1()
    
    Dim rs        As New ADODB.Recordset
    '修改
    Dim i         As Integer

    Dim strSql    As String
    Dim strsql1    As String
    
    Dim Prd_ID As String
    Dim SupSN As String
    Dim prd_name As String
    Dim issuestoragename As String
    Dim QTY As String
    Dim unit As String
    Dim DB_num As String
    Dim ZDCK As String
    Dim REMARK As String
    
    Dim count As Integer
    
    count = 0
    
    With fps

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 2
                If Trim(.text) <> "" Then
                    Prd_ID = Trim(.text)
                End If
                
                .Col = 3
                If Trim(.text) <> "" Then
                    SupSN = Trim(.text)
                End If

                .Col = 4
                If Trim(.text) <> "" Then
                    prd_name = Trim(.text)
                End If

                .Col = 5
                If Trim(.text) <> "" Then   '当前所在仓
                    issuestoragename = Trim(.text)
                End If
                .Col = 6
                If Trim(.text) <> "" Then
                    QTY = Trim(.text)       '当前总量
                End If
                
                .Col = 7
                If Trim(.text) <> "" Then   '可调拨量
                    unit = Trim(.text)
                End If
                
                .Col = 8
                If Trim(.text) <> "" Then
                    DB_num = Trim(.text)
                End If
                
                .Col = 9
                If Trim(.text) <> "" Then   '转调仓库
                    ZDCK = Trim(.text)
                End If
                
                .Col = 10
                If Trim(.text) <> "" Then
                    REMARK = Trim(.text)
                End If

                
                If DB_num = "" Then
                    MsgBox "调拨数量为必填且为数字！"
                ElseIf ZDCK = "" Then
                    MsgBox "转调仓库为必填！"
                ElseIf Val(QTY) < Val(DB_num) Then
                    MsgBox "调拨量超过总量！"
                Else
                SYL = QTY - DB_num
                
                strSql = "insert into  erptemp..tblErp_ShopOrderIssue_PLANT (Prd_ID,SupSN,Prd_Name,IssueQty,Unit,FromStorageName, " & _
                "IssueStorageName,IssueUser,Remark) values ('" & Prd_ID & "','" & SupSN & "','" & prd_name & "','" & DB_num & "','" & unit & "','" & issuestoragename & "','" & ZDCK & "','" & gUserName & "','" & REMARK & "') "
'                strsql1 = "insert into erptemp..tblErp_ShopOrderIssue_Inventory (Prd_ID,SupSN,Prd_Name,StorageName,SystemQty, " & _
'                "Unit,ActureQty,DailyUsageQty,IssueUser,Remark) values ( '" & Prd_ID & "','" & SupSN & "','" & prd_name & "','" & issuestoragename & "','" & QTY & "','" & unit & "','" & SYL & "','" & DB_num & "','" & gUserName & "','" & REMARK & "') "
'
'                strSql2 = "insert into erptemp..tblErp_ShopOrderIssue_Inventory (Prd_ID,SupSN,Prd_Name,StorageName,SystemQty, " & _
'                "Unit,ActureQty,DailyUsageQty,IssueUser,Remark) values ( '" & Prd_ID & "','" & SupSN & "','" & prd_name & "','" & ZDCK & "','" & QTY & "','" & unit & "','" & DB_num & "','" & SYL & "','" & gUserName & "','" & REMARK & "') "
                AddSql2 (strSql)
'                AddSql2 (strSql2)
'                AddSql2 (strsql1)
                
                count = count + 1
                End If
            End If
        Next i

    End With

    If count = 0 Then
        MsgBox "调拨失败"
    Else
        MsgBox "调拨成功" & "调拨记录数" & count & "! "
    
    End If

Query2

End Function

Private Function submit_CJPD2()
    Dim rs        As New ADODB.Recordset
    '修改
    Dim i         As Integer

    Dim strSql    As String
    Dim strsql1    As String
    
    Dim Prd_ID As String
    Dim SupSN As String
    Dim prd_name As String
    Dim issuestoragename As String
    Dim QTY As String
    Dim unit As String
    Dim DB_num As String
    Dim ZDCK As String
    Dim REMARK As String
    
    Dim count As Integer
    
    count = 0
    
    With fps

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 2     '料号
                If Trim(.text) <> "" Then
                    Prd_ID = Trim(.text)
                End If
                 
                .Col = 3     '批号
                If Trim(.text) <> "" Then
                    SupSN = Trim(.text)
                End If

                .Col = 4     '物料名称
                If Trim(.text) <> "" Then
                    prd_name = Trim(.text)
                End If

                .Col = 5     '仓名
                If Trim(.text) <> "" Then   '当前所在仓
                    issuestoragename = Trim(.text)
                End If
                .Col = 6     '总量
                If Trim(.text) <> "" Then
                    QTY = Trim(.text)       '当前总量
                End If
                
                .Col = 7     '单位
                If Trim(.text) <> "" Then   '可调拨量
                    unit = Trim(.text)
                End If
                
                .Col = 8    '盘点量
                If Trim(.text) <> "" Then
                    DB_num = Trim(.text)
                End If
                
                SYL = QTY - DB_num
                
                .Col = 9
                If Trim(.text) <> "" Then
                   Trim(.text) = SYL
                End If
                
                If DB_num = "" Then
                    MsgBox "盘点数量为必填且为数字！"
                Else
'
'                strSql = "insert into  erptemp..tblErp_ShopOrderIssue_PLANT (Prd_ID,SupSN,Prd_Name,IssueQty,Unit,FromStorageName, " & _
'                "IssueStorageName,IssueUser,Remark) values ('" & Prd_ID & "','" & SupSN & "','" & prd_name & "','" & DB_num & "','" & unit & "','" & issuestoragename & "','" & ZDCK & "','" & gUserName & "','" & REMARK & "') "
                strsql1 = "insert into erptemp..tblErp_ShopOrderIssue_Inventory (Prd_ID,SupSN,Prd_Name,StorageName,SystemQty, " & _
                "Unit,ActureQty,DailyUsageQty,IssueUser,Remark) values ( '" & Prd_ID & "','" & SupSN & "','" & prd_name & "','" & issuestoragename & "','" & QTY & "','" & unit & "','" & DB_num & "','" & SYL & "','" & gUserName & "','" & REMARK & "') "
        
'                AddSql2 (strSql)
                AddSql2 (strsql1)
 
                count = count + 1
                End If
            End If
        Next i

    End With

    If count = 0 Then
        MsgBox "调拨失败"
    Else
        MsgBox "调拨成功" & "调拨记录数" & count & "! "
    
    End If

query3

End Function

Private Function Query_CJPD2()
    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    Dim start1 As String
    Dim end1 As String
    
    start1 = DTP1(1).Value
    end1 = DTP2(0).Value
    
    start1 = Format(start1, "YYYY-MM-DD")
    end1 = Format(end1, "YYYY-MM-DD")
    
    strSql = "select * from (select '' as '选择',a.Prd_ID as '料号',a.SupSN as '批号',a.Prd_Name as '物料名称',a.issuestoragename as '仓名',(isnull(a.IssueQty,0)-isnull(b.IssueQty,0)-isnull(c.InventoryQty,0)) as '总量',a.unit as '单位','' as '盘点量','' as '耗用量','' as '备注',Sissuetime as '时间'" & _
    "from (select Prd_ID,SupSN,Prd_Name,issuestoragename,unit,SUM(IssueQty) as IssueQty,max(issuetime) as 'Sissuetime' from erptemp..tblErp_ShopOrderIssue_PLANT group by Prd_ID,SupSN,issuestoragename,unit,Prd_Name) A " & _
    "left join (select Prd_ID,SupSN,Prd_Name,fromstoragename,SUM(IssueQty) as IssueQty from erptemp..tblErp_ShopOrderIssue_PLANT where fromstoragename='CIS一楼黄光区线边仓' or fromstoragename='CIS二楼黄光区线边仓' or fromstoragename='CIS二楼电镀区线边仓' or fromstoragename='Bumping线边仓' or fromstoragename='WLP线边仓'or fromstoragename='12寸TSV制造部线边仓' group by Prd_ID,SupSN,fromstoragename,Prd_Name) B " & _
    "on a.Prd_ID=b.Prd_ID and a.SupSN=b.SupSN and a.issuestoragename=b.fromstoragename and a.Prd_Name = b.Prd_Name " & _
    "left join  (select prd_id,supsn,Prd_Name,storagename,sum(dailyusageqty) as InventoryQty from erptemp..tblErp_ShopOrderIssue_Inventory group by prd_id,Prd_Name,supsn,storagename) C " & _
    "on a.prd_id=c.prd_id and a.SupSN=c.SupSN and a.issuestoragename=c.storagename and a.Prd_Name=c.Prd_Name " & _
    "where (isnull(a.IssueQty,0)-isnull(b.IssueQty,0)-isnull(c.InventoryQty,0)) > 0) ss where 1=1"

   If Trim(txtPrd_id.text) <> "" Then
       strSql = strSql + "and ss.料号 = '" & Trim(txtPrd_id.text) & "' "
    End If

    If Trim(txtsup_sn.text) <> "" Then
       strSql = strSql + "and ss.批号 = '" & Trim(txtsup_sn.text) & "' "
    End If
    
    If Check1.Value = 1 Then
        strSql = strSql + " AND convert(varchar(10),ss.时间,23) >= '" & start1 & "'AND convert(varchar(10),ss.时间,23) <= '" & end1 & "'"
    End If

    strSQlQJ = strSql
    
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType3(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType3(rs)
        rs.Close
        Exit Function

    End If

End Function
Private Function Query_CJPD1()
    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    Dim start1 As String
    Dim end1 As String
    
    start1 = DTP1(1).Value
    end1 = DTP2(0).Value
    
    start1 = Format(start1, "YYYY-MM-DD")
    end1 = Format(end1, "YYYY-MM-DD")
    
    strSql = "select * from (select '' as '选择',a.Prd_ID as '料号',a.SupSN as '批号',d.物料名称,a.issuestoragename as '仓名',(isnull(a.IssueQty,0)-isnull(b.IssueQty,0)-isnull(c.InventoryQty,0)) as '总量',a.unit as '单位','' as '调拨数量','' as '转调仓库',a.Remark as '备注',a.Sissuetime as '时间' from " & _
    "(select Prd_ID,SupSN,issuestoragename,unit,SUM(IssueQty) as IssueQty,max(issuetime) as 'Sissuetime', Remark from erptemp..tblErp_ShopOrderIssue_PLANT group by Prd_ID,SupSN,issuestoragename,unit,Remark) A " & _
    "left join(select Prd_ID,SupSN,fromstoragename,SUM(IssueQty) as IssueQty,Remark from erptemp..tblErp_ShopOrderIssue_PLANT where fromstoragename='CIS一楼黄光区线边仓' or fromstoragename='CIS二楼黄光区线边仓' or fromstoragename='CIS二楼电镀区线边仓' or fromstoragename='Bumping线边仓' or fromstoragename='WLP线边仓' or fromstoragename='12寸TSV制造部线边仓' group by Prd_ID,SupSN,fromstoragename,Remark) B " & _
    "on a.Prd_ID=b.Prd_ID and a.SupSN=b.SupSN and a.issuestoragename=b.fromstoragename and a.Remark=b.Remark " & _
    "left join (select prd_id,supsn,storagename,sum(dailyusageqty) as InventoryQty from erptemp..tblErp_ShopOrderIssue_Inventory group by prd_id,supsn,storagename) C " & _
    "on a.prd_id=c.prd_id and a.SupSN=c.SupSN and a.issuestoragename=c.storagename " & _
    "left join (select 物料名称,料号 from erpdata..tblSmainM2) D on a.Prd_ID = d.料号  " & _
    "where (isnull(a.IssueQty,0)-isnull(b.IssueQty,0)-isnull(c.InventoryQty,0)) > 0 ) ss where 1=1"

    If Trim(txtPrd_id.text) <> "" Then
       strSql = strSql + "and ss.料号 = '" & Trim(txtPrd_id.text) & "'"
    End If

    If Trim(txtsup_sn.text) <> "" Then
       strSql = strSql + "and ss.批号 = '" & Trim(txtsup_sn.text) & "'"
    End If
    
    If Check1.Value = 1 Then
        strSql = strSql + " AND convert(varchar(10),ss.时间,23) >= '" & start1 & "'AND convert(varchar(10),ss.时间,23) <= '" & end1 & "'"
    End If
    
    strSQlQJ = strSql
    
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2
    End If

    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType2(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType2(rs)
        rs.Close
        Exit Function
    End If

End Function
Private Function Query_CJPD3()
   If Trim(txtPrd_id.text) <> "" Then
       strSql = strSql + "and ss.料号 = '" & Trim(txtPrd_id.text) & "'"
    End If

    If Trim(txtsup_sn.text) <> "" Then
       strSql = strSql + "and ss.批号 = '" & Trim(txtsup_sn.text) & "'"
    End If
End Function
        
Private Function Query_JHBZY()
  Dim strSql       As String
    Dim start1 As String
    Dim end1 As String
    
    start1 = DTP1(1).Value
    end1 = DTP2(0).Value
    
    start1 = Format(start1, "YYYY-MM-DD")
    end1 = Format(end1, "YYYY-MM-DD")
    
    Dim rs           As New ADODB.Recordset
    strSql = "select * from (select '' as '选择' ,a.Prd_ID as 料号,a.SupSN as 批号,c.物料名称,a.stockName as 当前所在仓,a.qty as 目前总量,(a.qty-isnull(b.qty,0)) as 可调拨量,a.unit as 单位,'' as '调拨数量','' as '转调仓库',case when a.Flag = '1' then '费领' else '调拨' end as '类型','' as '备注',Sissuetime as '时间' " & _
    "from (select Prd_ID, supsn ,stockName , sum(qty) as qty,unit,Flag,max(CreateDate) as 'Sissuetime' from erptemp..tblErp_ShopOrderIssue group by Prd_ID,supsn,stockName,unit,Flag) A " & _
    "left join (select Prd_ID, SupSN, fromstoragename,sum(IssueQty) as qty,unit from erptemp..tblErp_ShopOrderIssue_STORAGE group by Prd_ID,supsn,fromstoragename,unit) B " & _
    "on a.prd_id=b.prd_id and a.supsn=b.supsn and a.StockName = b.fromstoragename and a.unit=b.unit left join (select 物料名称,料号 from erpdata..tblSmainM2) C on a.Prd_ID = c.料号 " & _
    "where (a.qty-isnull(b.qty,0)) > 0 ) ss  where 1=1"

    If Trim(txtPrd_id.text) <> "" Then
       strSql = strSql + "and ss.料号 = '" & Trim(txtPrd_id.text) & "'"
    End If

    If Trim(txtsup_sn.text) <> "" Then
       strSql = strSql + "and ss.批号 = '" & Trim(txtsup_sn.text) & "'"
    End If

    If Check1.Value = 1 Then
        strSql = strSql + " AND convert(varchar(10),ss.时间,23) >= '" & start1 & "'AND convert(varchar(10),ss.时间,23) <= '" & end1 & "'"
    End If
    
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If
    
    strSQlQJ = strSql
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType1(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType1(rs)
        rs.Close
        Exit Function
    End If
End Function
        
Private Function Query_JHBBB()
    Dim strSql       As String
    Dim rs           As New ADODB.Recordset
    
    Dim start1 As String
    Dim end1 As String
    
    start1 = DTP1(1).Value
    end1 = DTP2(0).Value
    
    start1 = Format(start1, "YYYY-MM-DD")
    end1 = Format(end1, "YYYY-MM-DD")
    
    strSql = "select * from (select '' as '选择',Prd_ID as '料号',SupSN as '批号',Prd_Name as '物料名称', " & _
    "IssueQty as '单次调拨量', Unit as '单位',FromStorageName as '来源仓',IssueStorageName as '目的仓',IssueUser as '调拨人',IssueTime as '调拨时间', Remark as '备注' " & _
    "from erptemp..tblErp_ShopOrderIssue_STORAGE where IssueTime >= CONVERT(VARCHAR(10),GETDATE(),120) ) ss where 1=1 "

    If Trim(txtPrd_id.text) <> "" Then
       strSql = strSql + "and ss.料号 = '" & Trim(txtPrd_id.text) & "'"
    End If

    If Trim(txtsup_sn.text) <> "" Then
       strSql = strSql + "and ss.批号 = '" & Trim(txtsup_sn.text) & "' Order by 调拨时间 DESC"
    End If
    
   If Check1.Value = 1 Then
        strSql = strSql + " AND convert(varchar(10),ss.调拨时间,23) >= '" & start1 & "'AND convert(varchar(10),ss.调拨时间,23) <= '" & end1 & "'"
    End If
    
    If INIadoCon.State <> adStateOpen Then
        INIConnectSTART2

    End If
    
    strSQlQJ = strSql
    
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        Call ListDataType4(rs)
    Else
        MsgBox "无数据", vbInformation, "提示"
        Call ListDataType4(rs)
        rs.Close
        Exit Function

    End If
 
End Function
