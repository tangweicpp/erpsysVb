VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_TSV_BILLQuery 
   Caption         =   "工单查询"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13005
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
   ScaleHeight     =   9465
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CmbLine 
      Height          =   315
      ItemData        =   "Frm_TSV_BILLQuery.frx":0000
      Left            =   960
      List            =   "Frm_TSV_BILLQuery.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   14843
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "工单Head信息查询"
      TabPicture(0)   =   "Frm_TSV_BILLQuery.frx":0018
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fps(0)"
      Tab(0).Control(1)=   "ComQueryHead"
      Tab(0).Control(2)=   "ComOutLine"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "工单Line信息查询"
      TabPicture(1)   =   "Frm_TSV_BILLQuery.frx":0034
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(2)=   "ComQueryLine"
      Tab(1).Control(3)=   "fps(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "回货产品工单增加+"
      TabPicture(2)   =   "Frm_TSV_BILLQuery.frx":0050
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lbl"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblWAFERID"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtText1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdCommand3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdCommand3 
         Caption         =   "添加+号"
         Enabled         =   0   'False
         Height          =   360
         Left            =   2760
         TabIndex        =   22
         Top             =   2520
         Width           =   990
      End
      Begin VB.TextBox txtText1 
         Height          =   405
         Left            =   2400
         TabIndex        =   21
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         Caption         =   "查询"
         Height          =   615
         Left            =   -68040
         TabIndex        =   17
         Top             =   480
         Width           =   4695
         Begin VB.CommandButton Command1 
            Caption         =   "查询"
            Height          =   375
            Left            =   3720
            TabIndex        =   20
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox TxtWaferID 
            Height          =   375
            Left            =   840
            TabIndex        =   19
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label Label6 
            Caption         =   "WaferID:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "导出Excel"
         Height          =   360
         Left            =   -71040
         TabIndex        =   12
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton ComQueryLine 
         Caption         =   "查询"
         Height          =   360
         Left            =   -73440
         TabIndex        =   11
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton ComOutLine 
         Caption         =   "导出Excel"
         Height          =   360
         Left            =   -71040
         TabIndex        =   10
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton ComQueryHead 
         Caption         =   "查询"
         Height          =   360
         Left            =   -73440
         TabIndex        =   9
         Top             =   600
         Width           =   990
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   7095
         Index           =   0
         Left            =   -74880
         TabIndex        =   15
         Top             =   1200
         Width           =   15135
         _Version        =   524288
         _ExtentX        =   26696
         _ExtentY        =   12515
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
         SpreadDesigner  =   "Frm_TSV_BILLQuery.frx":006C
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   7095
         Index           =   1
         Left            =   -74880
         TabIndex        =   16
         Top             =   1200
         Width           =   15015
         _Version        =   524288
         _ExtentX        =   26485
         _ExtentY        =   12515
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
         SpreadDesigner  =   "Frm_TSV_BILLQuery.frx":04DC
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lblWAFERID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "只适合最后两位是WAFERID的产品。如果某些产品后三位是wafer号则不可用"
         Height          =   195
         Left            =   5760
         TabIndex        =   24
         Top             =   1560
         Width           =   5985
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输入工单号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         TabIndex        =   23
         Top             =   1680
         Width           =   1200
      End
   End
   Begin VB.TextBox TxtProduct 
      Height          =   375
      Left            =   12840
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox TxtWo 
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   240
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   191889409
      CurrentDate     =   41424
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   220856321
      CurrentDate     =   41424
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "线别："
      Height          =   195
      Left            =   480
      TabIndex        =   13
      Top             =   360
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间："
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "产品料号："
      Height          =   195
      Left            =   11880
      TabIndex        =   5
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工单号："
      Height          =   195
      Left            =   7800
      TabIndex        =   3
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束时间："
      Height          =   195
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "Frm_TSV_BILLQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
    E_id = 1                 'id
    E_CustomerNo             '客户代码
    E_Wo                     '工单号
    E_WOType                 '工单类型
    E_ProductName            '成品料号
    E_QTY                     '数量
    E_PieceQty                     '数量
    E_PMCCreatedate            '开工单日期
    E_PMCBegindate            '预计开始日
    E_PMCEnddate              '预计完工日
    E_PO                      '订单号
    E_POItem                   '订单号Item
    E_CustProductName          '客户料号
    E_Fab                      'FAB设备
    E_ImageCustomerRev         'ImageCustomerRev
    E_DesignID                  'DesignID
    E_Leval235                  'Level235
    E_Level260                  'Level260
    E_NGFlag                    'NG标志
    E_MarkingCode               'MarkingCode
    E_Rate                      '比率
    E_CountryFab                'CountryFab
    E_MicromMate                'MicromMate
    E_LotStatus                 'LotStatus
    e_MPN                      'MPN
    E_ProtectiveFilmApld      'ProtectiveFilmApld
    E_CustNeedDate             'CustNeedDate
    E_ShipSite                  'ShipSite
    E_WOCustomerNo             '接口中的客户代码
    E_DateCode                 'DateCode
    E_End
    
    
End Enum

Private Enum E_FPS1          'Detail
   E_id = 1                 'id
    E_Wo                     '工单号
    E_WaferId                'Waferid
    E_CompleteFlag           '完成标志
    E_TotalDie               '总数量
    E_gooddie                'good数量
    E_WaferLot               'wafer
    E_MarkingCode            'markingcode

    E_End
    
End Enum

Dim reportRS As New ADODB.Recordset
Dim mainItemRS As New ADODB.Recordset
Dim bomRS2        As New ADODB.Recordset



Private Sub CmdOut_Click()
Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(TxtWo.Text))
productTemp = UCase(Trim(TxtProduct.Text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
          
sql3 = " order by a.ordername,b.waferid  "

  
  If productTemp <> "" Then
  
  sql2 = " and a.product='" + productTemp + "'"
  
  End If
  
  
 If woTemp <> "" Then
  
  sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
  End If
  
  If Trim(sql2) <> "" Then
  
  sqlTemp = sql1 & sql2 & sql3
  
  Else
  
  sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<=to_date('" + endTime + "','YYYY/MM/DD')"
  
   sqlTemp = sql1 & sql2 & sql3
  
  End If
  
  
  
  
  
  ExporToExcel (sqlTemp)



End Sub

Private Sub cmdCommand3_Click()
Dim workOrderName As String

workOrderName = Trim(txtText1.Text)
If workOrderName = "" Then
MsgBox "请输入工单号 ！"
Exit Sub
End If

If Cnn.State = 0 Then
  ConOracle
End If

         Set adoCmd = New ADODB.Command
         Set adoCmd.ActiveConnection = Cnn
             adoCmd.CommandText = "WORKNAME_PLUS"
             'adoCmd.Parameters.Refresh
             adoCmd.CommandType = adCmdStoredProc
        
          Set adoprm1 = New ADODB.Parameter   '参数为工单号
          adoprm1.Type = adChar
          adoprm1.Size = 20
          adoprm1.Direction = adParamInput
          adoprm1.Value = workOrderName
          adoCmd.Parameters.Append adoprm1
          adoCmd.Execute

MsgBox "修改成功打印流程卡后确认下对错！"

End Sub

Private Sub Command1_Click()


Dim waferIdTemp As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""
waferIdTemp = UCase(Trim(txtWaferID.Text))


'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
'  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
'"  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & _
'" WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & _
'"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
'" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
'" from erpintegration2.ib_wohistory a where  modifyflag='1' "


  sql1 = " select b.wafersequence, b.ordername,b.waferid,b.completeflag,b.dieqty, b.FGDieQty , b.WaferLot, b.MarkingCode " & _
"from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername and b.waferid='" + waferIdTemp + "' "

          
          
sql3 = " order by a.ordername ,b.WaferLot,b.waferid "

  
 
  
  If Trim(sql2) <> "" Then
  
  sqlTemp = sql1 & sql2 & sql3
  
  Else
  
  sql2 = ""
  
   sqlTemp = sql1 & sql2 & sql3
  
  End If
  
  
  
  



Set mainItemRS = GetPMCWOLine(sqlTemp)

With fps(1)
        .MaxRows = 0
        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
       
        End If
End With





End Sub

Private Sub Command2_Click()


Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(TxtWo.Text))
productTemp = UCase(Trim(TxtProduct.Text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
'  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
'"  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & _
'" WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & _
'"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
'" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
'" from erpintegration2.ib_wohistory a where  modifyflag='1' "


  sql1 = " select b.wafersequence, b.ordername,b.waferid,b.completeflag,b.dieqty, b.FGDieQty , b.WaferLot, b.MarkingCode " & _
"from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername"

          
          
sql3 = " order by a.ordername ,b.WaferLot,b.waferid "

  
  If productTemp <> "" Then
  
  sql2 = " and a.product='" + productTemp + "'"
  
  End If
  
  
 If woTemp <> "" Then
  
  sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
  End If
  
  If Trim(sql2) <> "" Then
  
  sqlTemp = sql1 & sql2 & sql3
  
  Else
  
  sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
   sqlTemp = sql1 & sql2 & sql3
  
  End If
  
  
  
     ExporToExcel (sqlTemp)






End Sub

Private Sub ComOutLine_Click()

Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(TxtWo.Text))
productTemp = UCase(Trim(TxtProduct.Text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
"  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & _
" WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & _
"  PRODUCT ,QTY,Get_WoRPT_Piece(ORDERNAME) as 片数,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
" from erpintegration2.ib_wohistory a where  modifyflag='1' "
          
          
sql3 = " order by a.ordername  "

  
  If productTemp <> "" Then
  
  sql2 = " and a.product='" + productTemp + "'"
  
  End If
  
  
 If woTemp <> "" Then
  
  sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
  End If
  
  If Trim(sql2) <> "" Then
  
  sqlTemp = sql1 & sql2 & sql3
  
  Else
  
  sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
   sqlTemp = sql1 & sql2 & sql3
  
  End If
  
  
  
   ExporToExcel (sqlTemp)










End Sub

Private Sub ComQueryHead_Click()
'HEAD

Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(TxtWo.Text))
productTemp = UCase(Trim(TxtProduct.Text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
"  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & _
" WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & _
"  PRODUCT ,QTY, Get_WoRPT_Piece(ORDERNAME), ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
" from erpintegration2.ib_wohistory a where  modifyflag='1' "
          
          
sql3 = " order by a.ordername  "

  
  If productTemp <> "" Then
  
  sql2 = " and a.product='" + productTemp + "'"
  
  End If
  
  
 If woTemp <> "" Then
  
  sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
  End If
  
  If Trim(sql2) <> "" Then
  
  sqlTemp = sql1 & sql2 & sql3
  
  Else
  
  sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
   sqlTemp = sql1 & sql2 & sql3
  
  End If
  
  
  
  



Set reportRS = GetPMCWOHeader(sqlTemp)

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With





End Sub

Private Sub ComQueryLine_Click()
'Line


Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(TxtWo.Text))
productTemp = UCase(Trim(TxtProduct.Text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as 片数 , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
'  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
'"  CASE ORDERTYPE  WHEN '1' THEN '一般工单'  WHEN '5' THEN '再加工工单'   WHEN '7' THEN '委外工单'   WHEN '8' THEN '重工委外工单' " & _
'" WHEN '11' THEN '拆件式工单'    WHEN '13' THEN '预测工单'   WHEN '15' THEN '试产工单' Else '其他' END as ORDERTYPE ," & _
'"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
'" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
'" from erpintegration2.ib_wohistory a where  modifyflag='1' "


  sql1 = " select b.wafersequence, b.ordername,b.waferid,b.completeflag,b.dieqty, b.FGDieQty , b.WaferLot, b.MarkingCode " & _
"from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername"

          
          
sql3 = " order by a.ordername ,b.WaferLot,b.waferid "

  
  If productTemp <> "" Then
  
  sql2 = " and a.product='" + productTemp + "'"
  
  End If
  
  
 If woTemp <> "" Then
  
  sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
  End If
  
  If Trim(sql2) <> "" Then
  
  sqlTemp = sql1 & sql2 & sql3
  
  Else
  
  sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
   sqlTemp = sql1 & sql2 & sql3
  
  End If
  
  
  
  



Set mainItemRS = GetPMCWOLine(sqlTemp)

With fps(1)
        .MaxRows = 0
        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
       
        End If
End With


End Sub

Private Sub Form_Activate()
CmbLine.Text = "TSV"

DTP1.Value = Now - 1

DTP2.Value = Now

IniFpsHeader1
IniFpsHeader2

End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub IniFpsHeader1()
    With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
    
          
        .SetText E_FPS0.E_id, 0, "序号"
        .SetText E_FPS0.E_CustomerNo, 0, "客户代码"
        .SetText E_FPS0.E_Wo, 0, "工单号"
        .SetText E_FPS0.E_WOType, 0, "工单类型"
        .SetText E_FPS0.E_ProductName, 0, "成品料号"
        .SetText E_FPS0.E_QTY, 0, "Die数量"
        .SetText E_FPS0.E_PieceQty, 0, "片数"
        

        .SetText E_FPS0.E_PMCCreatedate, 0, "开单日期"
        .SetText E_FPS0.E_PMCBegindate, 0, "预计开始日"
        .SetText E_FPS0.E_PMCEnddate, 0, "预计完工日"
        .SetText E_FPS0.E_PO, 0, "订单号"
        .SetText E_FPS0.E_POItem, 0, "订单Item"
        .SetText E_FPS0.E_CustProductName, 0, "客户料号"
        .SetText E_FPS0.E_Fab, 0, "FAB设备"
        .SetText E_FPS0.E_ImageCustomerRev, 0, "ImageCustomerRev"
        .SetText E_FPS0.E_DesignID, 0, "DesignID"
        .SetText E_FPS0.E_Leval235, 0, "Level235"
        .SetText E_FPS0.E_Level260, 0, "Level260"
        .SetText E_FPS0.E_NGFlag, 0, "NG标志"
        .SetText E_FPS0.E_MarkingCode, 0, "MarkingCode"
        .SetText E_FPS0.E_Rate, 0, "比率"
        .SetText E_FPS0.E_CountryFab, 0, "CountryFab"
        .SetText E_FPS0.E_MicromMate, 0, "MicromMate"
        .SetText E_FPS0.E_LotStatus, 0, "LotStatus"
        .SetText E_FPS0.e_MPN, 0, "MPN"
        .SetText E_FPS0.E_ProtectiveFilmApld, 0, "ProtectiveFilmApld"
        .SetText E_FPS0.E_CustNeedDate, 0, "客户需求日"
        .SetText E_FPS0.E_ShipSite, 0, "ShipSite"
        .SetText E_FPS0.E_WOCustomerNo, 0, "接口中的客户代码"
        .SetText E_FPS0.E_DateCode, 0, "DateCode"



        .ColWidth(E_FPS0.E_id) = 10
        .ColWidth(E_FPS0.E_CustomerNo) = 10
        .ColWidth(E_FPS0.E_Wo) = 10
        .ColWidth(E_FPS0.E_WOType) = 10
        .ColWidth(E_FPS0.E_ProductName) = 10
        .ColWidth(E_FPS0.E_QTY) = 10
        .ColWidth(E_FPS0.E_PieceQty) = 5
        
        .ColWidth(E_FPS0.E_PMCCreatedate) = 10
        .ColWidth(E_FPS0.E_PMCBegindate) = 10
        .ColWidth(E_FPS0.E_PMCEnddate) = 10
        .ColWidth(E_FPS0.E_PO) = 10
        .ColWidth(E_FPS0.E_POItem) = 10
        .ColWidth(E_FPS0.E_CustProductName) = 10
        
        .ColWidth(E_FPS0.E_Fab) = 10
        .ColWidth(E_FPS0.E_ImageCustomerRev) = 10
        .ColWidth(E_FPS0.E_DesignID) = 10
        .ColWidth(E_FPS0.E_Leval235) = 10
        .ColWidth(E_FPS0.E_Level260) = 10
        .ColWidth(E_FPS0.E_NGFlag) = 10
        
        .ColWidth(E_FPS0.E_MarkingCode) = 10
        .ColWidth(E_FPS0.E_Rate) = 10
        .ColWidth(E_FPS0.E_CountryFab) = 10
        .ColWidth(E_FPS0.E_MicromMate) = 10
        .ColWidth(E_FPS0.E_LotStatus) = 10
        .ColWidth(E_FPS0.e_MPN) = 10
        
         .ColWidth(E_FPS0.E_ProtectiveFilmApld) = 10
        .ColWidth(E_FPS0.E_CustNeedDate) = 10
        .ColWidth(E_FPS0.E_ShipSite) = 10
        .ColWidth(E_FPS0.E_WOCustomerNo) = 10
        .ColWidth(E_FPS0.E_DateCode) = 10
 
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub

Private Sub IniFpsHeader2()
    With fps(1)
        .ReDraw = False
        .MaxCols = E_FPS1.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
          
        .SetText E_FPS1.E_id, 0, "序号"
        .SetText E_FPS1.E_Wo, 0, "工单号"
        .SetText E_FPS1.E_WaferId, 0, "WaferId"
        .SetText E_FPS1.E_CompleteFlag, 0, "完成标志"
        .SetText E_FPS1.E_TotalDie, 0, "TotalDie数量"
        .SetText E_FPS1.E_gooddie, 0, "GoodDie数量"
        .SetText E_FPS1.E_WaferLot, 0, "WaferLot"
        .SetText E_FPS1.E_MarkingCode, 0, "MarkingCode"
        
        
        .ColWidth(E_FPS1.E_id) = 10
        .ColWidth(E_FPS1.E_Wo) = 10
        .ColWidth(E_FPS1.E_WaferId) = 10
        .ColWidth(E_FPS1.E_CompleteFlag) = 10
        .ColWidth(E_FPS1.E_TotalDie) = 10
        .ColWidth(E_FPS1.E_gooddie) = 10
        .ColWidth(E_FPS1.E_WaferLot) = 10
        .ColWidth(E_FPS1.E_MarkingCode) = 10
        
     

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub



