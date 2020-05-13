VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Form_Prod_Control 
   Caption         =   "工单-排程"
   ClientHeight    =   13050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21750
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
   MinButton       =   0   'False
   ScaleHeight     =   13050
   ScaleMode       =   0  'User
   ScaleWidth      =   21750
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "排程"
      Height          =   13095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   21615
      Begin VB.CommandButton cmdflag01 
         Caption         =   "抛送中查询"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   11880
         TabIndex        =   25
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtQTY 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   16320
         TabIndex        =   24
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CheckBox chk02 
         Caption         =   "全选"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   3840
         Width           =   735
      End
      Begin VB.CheckBox chk01 
         Caption         =   "未抛送"
         Height          =   255
         Left            =   14280
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdMES 
         Caption         =   "抛MES"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   11880
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdPC 
         Caption         =   "排程"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   9720
         TabIndex        =   19
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtorder_qty 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13800
         TabIndex        =   14
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtLot 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   11
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtShop_Order 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtdept 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtdevice 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtcust 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton cmd 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   9720
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   8295
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   4080
         Width           =   21135
         _Version        =   524288
         _ExtentX        =   37280
         _ExtentY        =   14631
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
         MaxCols         =   21
         MaxRows         =   0
         SpreadDesigner  =   "Form_Prod_Control.frx":0000
         Appearance      =   1
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   0
         Left            =   8040
         TabIndex        =   15
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   65280
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16777215
         Format          =   161546241
         CurrentDate     =   43271
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   1
         Left            =   8040
         TabIndex        =   16
         Top             =   3120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   65280
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16777215
         Format          =   161546241
         CurrentDate     =   43271
      End
      Begin VB.Label lbl09 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "片数:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   15360
         TabIndex        =   23
         Top             =   3240
         Width           =   840
      End
      Begin VB.Label lbl08 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "计划完成时间:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5760
         TabIndex        =   18
         Top             =   3120
         Width           =   2160
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "计划开始时间:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5760
         TabIndex        =   17
         Top             =   2520
         Width           =   2160
      End
      Begin VB.Label lbl06 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单数:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12480
         TabIndex        =   13
         Top             =   3240
         Width           =   1170
      End
      Begin VB.Label lbl05 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOT:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   6
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lbl04 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lbl03 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "事业部:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label lbl02 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "机种:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lbl01 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   780
      End
   End
End
Attribute VB_Name = "Form_Prod_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum fpSOrder
    e_ID
    e_choice
    E_dept
    e_order
    e_create
    E_cust
    e_device
    e_PKG
    E_LOT
    e_wqty
    e_wlist
    e_start
    E_END
    E_FLAG
    E_REMARK
    e_MCol
End Enum


Private Sub chk02_Click()

Dim FLAG As Integer
Dim i As Integer
Dim WAFER_QTY As Integer




With Fps(0)

 For i = 1 To .MaxRows
 .Row = i
  .Col = 9
 WAFER_QTY = Val(.text)
 .Col = 1
 
   If Abs(Val(.Value) - 1) = 1 Then
      
    
        If Trim(txtorder_qty.text) = "" Then
          txtorder_qty.text = 1
          txtQTY.text = WAFER_QTY
          
        Else
          txtorder_qty.text = Val(txtorder_qty.text) + 1
          txtQTY.text = Val(txtQTY.text) + WAFER_QTY
          
        End If
    Else
          txtorder_qty.text = Val(txtorder_qty.text) - 1
           txtQTY.text = Val(txtQTY.text) - WAFER_QTY
    End If
        
     .Value = Abs(Val(.Value) - 1)
 
 Next
    End With

End Sub

Private Sub cmdflag01_Click()

Query2

End Sub

Private Sub Form_Load()

 With Fps(0)
 
    .Col = -1
    .Row = -1
    .Lock = True
    .SetText 1, 0, "选择"
    .ColWidth(e_choice) = 2
    .SetText 2, 0, "部门"
      .ColWidth(E_dept) = 8
    .SetText 3, 0, "工单号"
      .ColWidth(e_order) = 15
    .SetText 4, 0, "工单开立时间"
    .ColWidth(e_create) = 15
    .SetText 5, 0, "客户代码"
      .ColWidth(E_cust) = 8
    .SetText 6, 0, "机种"
      .ColWidth(e_device) = 15
    .SetText 7, 0, "结构"
      .ColWidth(e_PKG) = 8
     .SetText 8, 0, "LOT"
      .ColWidth(E_LOT) = 15
    .SetText 9, 0, "片数"
      .ColWidth(e_wqty) = 5
    .SetText 10, 0, "WAFER_LIST"
    .ColWidth(e_wlist) = 15
      .SetText 11, 0, "计划开始时间"
    .ColWidth(e_start) = 15
        .SetText 12, 0, "计划完成时间"
    .ColWidth(E_END) = 15
        .SetText 13, 0, "状态"
    .ColWidth(E_FLAG) = 5
        .SetText 14, 0, "备注"
    .ColWidth(E_REMARK) = 20
    
    
 End With


 With Fps(0)
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .ColsFrozen = 5
        .MaxCols = fpSOrder.e_MCol - 1
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        .ZOrder
        .ReDraw = True
    End With
    
  chk01.Value = 1
  
  DTPicker1(1).Value = Format(Year(Now()) & "-" & Month(Now()) & "-" & "28", "yyyy-MM-dd")
  DTPicker1(0).Value = Format(Now(), "yyyy-MM-dd")
  

End Sub


Private Sub cmd_Click()


 If chk01.Value = 1 Then
    
    Query
    
 Else

    Query1
 End If
 


End Sub


Private Sub Query()

 Dim rs         As New ADODB.Recordset

 Dim strSql     As String
 
 txtorder_qty.text = ""

strSql = " select '' as 选择, x.updateprice2 AS 部门, x.shop_order AS 工单号,to_char( x.erp_create_date,'YYYY-MM-DD') AS 工单开立时间, x.cust_id AS 客户代码, x.ht_device AS 机种, x.struckstr2 AS 结构, x.cust_lot_id AS LOT,count(distinct wafer_id) as 片数, max(wafer_id) as WAFER_LIST " & _
             " , '' as 计划开始时间  ,to_char( x.plan_end_date,'YYYY-MM-DD') AS 计划完成时间 , x.状态, x.memo AS 备注 from (select a.shop_order, a.cust_id,  a.ht_device, b.cust_lot_id, a.plan_star_date " & _
             "  ,a.erp_create_date  , a.plan_end_date  , decode(a.flag, 8, '待抛', '失败') as 状态,a.memo, c.struckstr2, c.updateprice2, to_char(wm_concat(substr(replace(b.wafer_id, '+', ''), length(replace(b.wafer_id, '+', '')) - 1, 2)) " & _
             "  over(partition by b.cust_lot_id order by b.wafer_id)) as wafer_id  from shop_order a, ib_wohistory aa, shop_order_detail b, tbltsvnpiproduct  c where a.flag in ('8', '3') " & _
             " and b.shop_order = a.shop_order and aa.ordername = a.shop_order  and c.customershortname = aa.customer and c.qtechptno2 = aa.product "
           


      If Trim(txtcust.text) <> "" Then
        
        strSql = strSql + " AND aa.customer = '" & Trim(txtcust.text) & "'"
        
      End If
      
      If Trim(txtShop_Order.text) <> "" Then
      
       strSql = strSql + " AND b.shop_order = '" & Trim(txtShop_Order.text) & "'"
        
      End If
      
     If Trim(txtdevice.text) <> "" Then
      
       strSql = strSql + " AND  A.ht_device = '" & Trim(txtdevice.text) & "'"
        
      End If
      
      If Trim(txtLot.text) <> "" Then
      
       strSql = strSql + " AND  b.cust_lot_id = '" & Trim(txtLot.text) & "'"
        
      End If
      
    If Trim(txtdept.text) <> "" Then
      
       strSql = strSql + " AND   c.updateprice2 = '" & Trim(txtdept.text) & "'"
        
      End If
      
      
     strSql = strSql + " ) x  group by x.cust_lot_id, x.shop_order, x.cust_id, x.ht_device,x.plan_star_date,x.plan_end_date, x.状态,  x.Memo , x.StruckStr2, x.UpdatePrice2,x.erp_create_date "
   
    
    Fps(0).MaxRows = 0

    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType(rs)
    Else
        MsgBox "没有待抛工单", vbInformation, "提示"
        Exit Sub

    End If

End Sub



Private Sub Query1()

 Dim rs         As New ADODB.Recordset

 Dim strSql     As String
 
 txtorder_qty.text = ""
 
    If Trim(txtcust.text) = "" And Trim(txtdept.text) = "" Then
        
        MsgBox " 已下线工单请输入客户代码或者事业部查询近10天开立的工单", vbInformation, "提示"
        Exit Sub
        
      End If
 

strSql = " select '' as 选择, x.updateprice2, x.shop_order,to_char( x.erp_create_date,'YYYY-MM-DD') AS 工单开立时间, x.cust_id, x.ht_device,x.struckstr2, x.cust_lot_id,  count(distinct wafer_id) as 片数, max(wafer_id) as WAFER_LIST " & _
         " , to_char(x.plan_star_date, 'YYYY-MM-DD') as 计划开始时间, to_char(x.plan_end_date, 'YYYY-MM-DD') AS 计划完成时间, x.jobid as 抛送日期, replace(nvl(x.状态, '未打印'), x.状态, '') as 状态 " & _
         " from (select a.shop_order, a.cust_id,a.ht_device, b.cust_lot_id,a.erp_create_date,  a.plan_star_date, a.plan_end_date,a.bonded as 状态,a.jobid,  a.memo,c.struckstr2,c.updateprice2, " & _
         " to_char(wm_concat(substr(replace(b.wafer_id, '+', ''), length(replace(b.wafer_id, '+','')) - 1,  2))  over(partition by b.cust_lot_id order by b.wafer_id)) as wafer_id " & _
         " from shop_order  a, ib_wohistory  aa,  shop_order_detail b, tbltsvnpiproduct  c  where a.flag in ('2') and b.shop_order = a.shop_order  and to_char(a.erp_create_date, 'YYYY-MM-DD') > " & _
         "   to_char(sysdate - 5, 'YYYY-MM-DD') and aa.ordername = a.shop_order and c.customershortname = aa.customer  and c.qtechptno2 = aa.product"
           


      If Trim(txtcust.text) <> "" Then
        
        strSql = strSql + " AND aa.customer = '" & Trim(txtcust.text) & "'"
        
      End If
      
      If Trim(txtShop_Order.text) <> "" Then
      
       strSql = strSql + " AND b.shop_order = '" & Trim(txtShop_Order.text) & "'"
        
      End If
      
     If Trim(txtdevice.text) <> "" Then
      
       strSql = strSql + " AND  A.ht_device = '" & Trim(txtdevice.text) & "'"
        
      End If
      
      If Trim(txtLot.text) <> "" Then
      
       strSql = strSql + " AND  b.cust_lot_id = '" & Trim(txtLot.text) & "'"
        
      End If
      
    If Trim(txtdept.text) <> "" Then
      
       strSql = strSql + " AND   c.updateprice2 = '" & Trim(txtdept.text) & "'"
        
      End If
      
      
     strSql = strSql + " ) x  group by x.cust_lot_id,  x.shop_order, x.cust_id, x.ht_device, x.plan_star_date, x.plan_end_date, x.状态,x.memo, x.struckstr2, x.updateprice2,x.erp_create_date,x.jobid  "
   
    
    Fps(0).MaxRows = 0

    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType(rs)
    Else
        MsgBox "没有待抛工单", vbInformation, "提示"
        Exit Sub

    End If

End Sub


Private Sub Query2()

 Dim rs         As New ADODB.Recordset

 Dim strSql     As String
 
 txtorder_qty.text = ""
 
   
 

strSql = " select '' as 选择, x.updateprice2, x.shop_order,to_char( x.erp_create_date,'YYYY-MM-DD') AS 工单开立时间, x.cust_id, x.ht_device,x.struckstr2, x.cust_lot_id,  count(distinct wafer_id) as 片数, max(wafer_id) as WAFER_LIST " & _
         " , to_char(x.plan_star_date, 'YYYY-MM-DD') as 计划开始时间, to_char(x.plan_end_date, 'YYYY-MM-DD') AS 计划完成时间, x.状态, replace(nvl(x.状态, '未打印'), x.状态, '') as 状态 " & _
         " from (select a.shop_order, a.cust_id,a.ht_device, b.cust_lot_id,a.erp_create_date,  a.plan_star_date, a.plan_end_date,a.bonded as 状态, a.memo,c.struckstr2,c.updateprice2, " & _
         " to_char(wm_concat(substr(replace(b.wafer_id, '+', ''), length(replace(b.wafer_id, '+','')) - 1,  2))  over(partition by b.cust_lot_id order by b.wafer_id)) as wafer_id " & _
         " from shop_order  a, ib_wohistory  aa,  shop_order_detail b, tbltsvnpiproduct  c  where a.flag in ('1','0') and b.shop_order = a.shop_order  and to_char(a.erp_create_date, 'YYYY-MM-DD') > " & _
         "   to_char(sysdate - 10, 'YYYY-MM-DD') and aa.ordername = a.shop_order and c.customershortname = aa.customer  and c.qtechptno2 = aa.product ) x  group by x.cust_lot_id,  x.shop_order, x.cust_id, x.ht_device, x.plan_star_date, x.plan_end_date, x.状态,x.memo, x.struckstr2, x.updateprice2 ,x.erp_create_date"
           

    
    Fps(0).MaxRows = 0

    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType(rs)
    Else
        MsgBox "没有待抛工单", vbInformation, "提示"
        Exit Sub

    End If

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
            .ColWidth(1) = 2
            .CellType = CellTypeCheckBox
            .text = 0
            
            
            .Col = 10
            .Lock = False
         '   .CellType = CellTypeDate
            
            .Col = 11
            .Lock = False
         '   .CellType = CellTypeDate
            
            
        Next
        
    End With

End Sub

Private Sub fps_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim WAFER_QTY As Integer

If Row < 1 Then Exit Sub

If Col = 10 Or Col = 11 Then
    Fps(0).Row = Row
    Fps(0).Col = Col
    Fps(0).CellType = CellTypeDate
End If


If Col <> 1 Then Exit Sub

With Fps(0)
    .Row = Row
    .Col = 9
    WAFER_QTY = Val(.text)
    .Col = 1
    If Abs(Val(.Value) - 1) = 1 Then
        If Trim(txtorder_qty.text) = "" Then
            txtorder_qty.text = 1
            txtQTY.text = WAFER_QTY
        Else
            txtorder_qty.text = Val(txtorder_qty.text) + 1
            txtQTY.text = Val(txtQTY.text) + WAFER_QTY

        End If

    Else
        txtorder_qty.text = Val(txtorder_qty.text) - 1
        txtQTY.text = Val(txtQTY.text) - WAFER_QTY

    End If

    .Value = Abs(Val(.Value) - 1)

End With

End Sub

Private Sub cmdPC_Click()
Dim i        As Integer
Dim strstart As String
Dim strend   As String

strstart = Format(DTPicker1(0).Value, "YYYY-MM-DD")
strend = Format(DTPicker1(1).Value, "YYYY-MM-DD")

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .text = 1 Then
            .Col = 11
            .CellType = CellTypeEdit
            .text = strstart
            .Col = 12
            .CellType = CellTypeEdit
            .text = strend

        End If

    Next

End With

End Sub

Private Sub cmdMes_Click()
Dim strup As String
Dim strstart As String
Dim strend As String
Dim SHOP_ORDER As String
Dim i As Integer
Dim j As Integer
j = 0

 With Fps(0)
 For i = 1 To .MaxRows
 .Row = i
 .Col = 1

If .text = 1 Then
 .Col = 3
SHOP_ORDER = .text
.Col = 11
strstart = .text
If Trim(strstart) = "" Then

 MsgBox " 工单计划开始时间不能为空!"
  
Exit Sub
End If

.Col = 12
strend = .text

strup = "update shop_order set plan_star_date = '" & strstart & "',plan_end_date = '" & strend & "' ,flag = 0 ,jobid = sysdate where  shop_order = '" & SHOP_ORDER & "'"
AddSql (strup)
    
j = j + 1
End If

 Next
 End With


 MsgBox "已更新" & j & "笔工单抛送状态,每笔工单需要1分钟时间抛送MES，请知悉!"
 Query
 
 

End Sub


















