VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_WAFER_SEND 
   Caption         =   "WAFER发料"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17835
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
   ScaleHeight     =   11010
   ScaleMode       =   0  'User
   ScaleWidth      =   17835
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame WAFER 
      Height          =   13215
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   17775
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   15000
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdsend 
         Caption         =   "领料单生成"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5040
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtOrder 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   2895
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   3615
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   17535
         _Version        =   524288
         _ExtentX        =   30930
         _ExtentY        =   6376
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
         SpreadDesigner  =   "Frm_WAFER_SEND.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   7095
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   6000
         Width           =   17535
         _Version        =   524288
         _ExtentX        =   30930
         _ExtentY        =   12515
         _StockProps     =   64
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         MaxRows         =   0
         SpreadDesigner  =   "Frm_WAFER_SEND.frx":050C
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lbl04 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 备注:近100天开立工单明细,超过100的请确认工单是否有效"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   4665
      End
      Begin VB.Label lbl03 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发料清单:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   5760
         Width           =   780
      End
      Begin VB.Label lbl02 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "待发明细:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号："
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1320
      End
   End
End
Attribute VB_Name = "Frm_WAFER_SEND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum FpsD

    e_ID
    e_shop_order
    E_LOT
    e_manum
    E_maid
    E_qty
    E_stock
    e_dept
    e_order_num
    e_MCol

End Enum


Private Sub cmdRefresh_Click()

 With Fps(0)
        .MaxRows = 0
    End With
 With Fps(1)
        .MaxRows = 0
    End With


txtOrder.text = ""
cmdsend.Enabled = True
ListData

End Sub



Private Sub Form_Load()

 With Fps(1)
 

    .Col = -1
    .Row = -1
    .Lock = True
    .SetText 1, 0, "工单号"
    .ColWidth(e_shop_order) = 15
    .SetText 2, 0, "LOT"
      .ColWidth(E_LOT) = 10
    .SetText 3, 0, "晶圆料号"
      .ColWidth(e_manum) = 15
    .SetText 4, 0, "物料编码"
      .ColWidth(E_maid) = 15
    .SetText 5, 0, "数量"
      .ColWidth(E_qty) = 8
    .SetText 6, 0, "仓别"
      .ColWidth(E_stock) = 5
    .SetText 7, 0, "部门"
      .ColWidth(e_dept) = 10
    .SetText 8, 0, "领料单"
      .ColWidth(e_order_num) = 18
    
    
 End With

ListData

End Sub



Private Sub ListData()
Dim strsql As String
Dim rs     As New ADODB.Recordset

strsql = " SELECT x.工单号,x.PARA8 as 部门,x.F_101 AS 晶圆料号,x.物料编号,CONVERT(INT, x.用量) AS 需求量,CONVERT(INT, ISNULL(x.实领数量,0)) AS 已领数,x.WAFERLOT AS 批号,x.工单数 AS 工单片数 " & _
         " ,ISNULL(y.物料编号,'') AS 库存物料编号 ,ISNULL(y.仓库编号,'') AS 仓别 ,CONVERT(INT, ISNULL(y.当前存量,0)) AS 库存数,'' AS  状态  FROM ( " & _
         " select b.工单号,d.PARA8,a.F_101,b.物料编号,b.用量,b.实领数量,c.WAFERLOT,COUNT(c.WAFERID) AS 工单数 ,CONVERT(VARCHAR(100), d.ERPCREATEDATE ,23) AS 开立时间 FROM " & _
         " AIS20141114094336..t_ICItem a,ERPBASE .. tblllplan b,erpdata .. tblTSVwaferlist c,erpdata .. tblTSVworkorder d WHERE SUBSTRING(a.F_101,1,2) = '60' " & _
         "  AND b.物料编号 = a.FNumber AND b.用量 <> ISNULL(b.实领数量,0) AND c.ORDERNAME = b.工单号 AND d.ORDERNAME = c.ORDERNAME AND b.工单号 NOT LIKE '%M-%' " & _
         "  AND CONVERT(VARCHAR(100),d.ERPCREATEDATE,23) >= CONVERT(VARCHAR(100),GETDATE()-100,23 ) GROUP BY b.工单号,a.F_101, b.物料编号,b.用量,b.实领数量 ,c.WAFERLOT " & _
         "  ,CONVERT(VARCHAR(100), d.ERPCREATEDATE ,23) ,d.PARA8) x  LEFT JOIN ERPBASE..tblStockNum y  ON y.批号 = x.WAFERLOT AND y.当前存量 >0 AND y.仓库编号  <> '54'" & _
         "   LEFT JOIN ERPBASE..tblStockSQ2 z ON z.物料编号 = x.物料编号  AND z.批号 = x.WAFERLOT AND z.工单号 = x.工单号 WHERE z.单据编号 IS null    ORDER BY x.开立时间 "


If rs.State = adStateOpen Then rs.Close

rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

If Not rs.EOF Then
   Call ListDataType(rs)
   
   Else
   
     MsgBox "无待领晶圆信息", vbInformation, "提示"
   
End If


End Sub





Private Sub ListDataType(rs As ADODB.Recordset)

    Dim i As Long

    Dim strmater As String
    Dim strmater1 As String
    
    Dim qty1 As Integer
    Dim qty2 As Integer
    Dim orderqty As Integer
    Dim stock As Integer
    

    With Fps(0)
        
        .MaxRows = 0
        .Lock = True

        Set .DataSource = rs

    End With
    
    With Fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4
            strmater = Trim(.text)
            .Col = 5
            qty1 = Val(.text)
             .Col = 6
             qty2 = Val(.text)
             .Col = 8
             orderqty = Val(.text)
              .Col = 9
             strmater1 = Trim(.text)
              .Col = 11
             stock = Val(.text)
             .Col = 12
             
             If strmater <> strmater1 Or qty1 <> orderqty Or qty2 > qty1 Then
                
                .text = "ERROR"
                .BackColor = &HFF&
             Else
                             
             If Val(stock) <> Val(qty1 - qty2) Then
             
                  .text = "WARN"
                   .BackColor = &HFFFF&
              Else
                  
                  .text = "READY"
                  .BackColor = &HFF00&
                  
              End If
              
             End If
             
     
        Next

    End With

End Sub




Private Sub txtOrder_KeyPress(KeyAscii As Integer)

If KeyAscii <> vbKeyReturn Then
    Exit Sub

End If

Dim dept As String
Dim Strma As String
Dim strmaid As String
Dim Strqty1 As String
Dim Strqty2 As String
Dim strlot As String
Dim strstock As String
Dim i As Integer
Dim j As Integer
Dim k As Integer


k = 0

 With Fps(0)
    For i = 1 To .MaxRows
    .Row = i
    .Col = 1
    If Trim(txtOrder.text) = Trim(.text) Then
    .Col = 12
    If .text = "READY" Then
      .Col = 2
      dept = Trim(.text)
      .Col = 3
      Strma = Trim(.text)
      .Col = 4
      strmaid = Trim(.text)
      .Col = 5
      Strqty1 = Trim(.text)
      .Col = 6
      Strqty2 = Trim(.text)
      .Col = 7
      strlot = Trim(.text)
       .Col = 10
      strstock = Trim(.text)
      
     k = k + 1
      
    ElseIf .text = "WARN" Then
    
      .Col = 2
      dept = Trim(.text)
      .Col = 3
      Strma = Trim(.text)
      .Col = 4
      strmaid = Trim(.text)
      .Col = 5
      Strqty1 = Trim(.text)
      .Col = 6
      Strqty2 = Trim(.text)
      .Col = 7
      strlot = Trim(.text)
       .Col = 10
      strstock = Trim(.text)
      MsgBox "库存数大于需求数,实物请根据工单ID发放", vbCritical, "警告"
      
      k = k + 1
    
    Else
    
        
      MsgBox "请确认晶圆库存是否满足工单需求", vbCritical, "警告"
      Exit Sub
        
    End If
    
    With Fps(1)
    For j = 1 To .MaxRows
    .Row = j
    .Col = 1
    If .text = Trim(txtOrder.text) Then
    
    MsgBox "请不要选择重复工单", vbCritical, "警告"
      Exit Sub
    
    End If

     Next
    End With
    
     
    With Fps(1)
      
      
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .text = Trim(txtOrder.text)
     .Col = 2
     .text = strlot
      .Col = 3
     .text = Strma
      .Col = 4
     .text = strmaid
      .Col = 5
     .text = Val(Strqty1) - Val(Strqty2)
      .Col = 6
     .text = strstock
      .Col = 7
     .text = dept

    End With

    End If

  Next
    End With

txtOrder.text = ""

End Sub


Private Sub CmdSend_Click()

Dim Strma As String
Dim Strma1 As String
Dim strqty As Integer
Dim strlot As String
Dim strstock As String
Dim strorder As String
Dim strdept As String
Dim i As Integer
Dim id As String
Dim strid As String
Dim stridsave As String
Dim strsend1 As String
Dim strsend2 As String
Dim User         As String

User = gUserName

 INIadoCon.BeginTrans

 With Fps(1)
    For i = 1 To .MaxRows
    .Row = i
    .Col = 1
    strorder = .text
    .Col = 2
    strlot = .text
    .Col = 3
    Strma1 = .text
    .Col = 4
    Strma = .text
    .Col = 5
    strqty = Val(.text)
    .Col = 6
    strstock = Trim(.text)
    .Col = 7
    strdept = Trim(.text)
    
    strid = "select  'LW'|| TO_CHAR(SYSDATE,'YYYYMMDD')||  lpad(send_num.nextval,4,'0') from dual "
    
    id = getStr2(strid)
    
    strsend1 = "  insert  into tblstockSQ2  " & _
               "  (单据编号,序号,单据日期,物料编号,数量,仓库编号,领料部门,工序,备注,领料员,接收标记,标记,主管,审核数,审核部门,工单号,批号) " & _
               "  Values " & _
               " ('" & id & "',1,getdate(),'" & Strma & "','" & strqty & "','" & strstock & "','" & strdept & "', 1 ,'工单号:'+rtrim('" & strorder & "') " & _
               " ,'E17363 秦智E',0,1,'E13323 聂运斌E' ,'" & strqty & "','" & strdept & "','" & strorder & "','" & strlot & "') "
    
    If AddSql2(strsend1) = 0 Then
        
        GoTo DealError
       
    End If
    
    strsend2 = "  INSERT INTO erptemp..wafer_send  " & _
               "  (shop_order,LOT,material_id ,material_num,qty,storehouse,order_num,dept,create_date,create_by,flag ) " & _
               "  Values " & _
               "   ('" & strorder & "','" & strlot & "','" & Strma & "','" & Strma1 & "','" & strqty & "','" & strstock & "','" & id & "','" & strdept & "',getdate(),'" & User & "',0) "
    
       
      If AddSql2(strsend2) = 0 Then
        
        GoTo DealError
       
    End If
       
    .Col = 8
    .text = id
    
  Next
    End With
INIadoCon.CommitTrans
    
cmdsend.Enabled = False

Exit Sub
DealError:
INIadoCon.RollbackTrans
MsgBox "保存失败！" + Chr(13) + "原因:" + Err.DESCRIPTION, vbInformation, Me.Caption
    
End Sub




















