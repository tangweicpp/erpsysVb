VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form Frm_PackList_Label 
   Caption         =   "PackList WaferID标签"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17055
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
   ScaleHeight     =   8940
   ScaleWidth      =   17055
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "报表"
      Height          =   7335
      Left            =   9600
      TabIndex        =   10
      Top             =   600
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "导出Excel"
         Height          =   480
         Left            =   1560
         TabIndex        =   15
         Top             =   1920
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   180224001
         CurrentDate     =   41424
      End
      Begin MSComCtl2.DTPicker DTP2 
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   180224001
         CurrentDate     =   41424
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期："
         Height          =   195
         Left            =   840
         TabIndex        =   14
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期： "
         Height          =   195
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查询"
      Height          =   360
      Left            =   6600
      TabIndex        =   8
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox TxtLotId 
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.CheckBox ChkAll 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox TxtDir 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   8160
      Width           =   6855
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "取消"
      Height          =   360
      Left            =   9480
      TabIndex        =   2
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "生成标签"
      Height          =   360
      Left            =   8040
      TabIndex        =   1
      Top             =   120
      Width           =   990
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   7335
      Index           =   0
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   8055
      _Version        =   196613
      _ExtentX        =   14208
      _ExtentY        =   12938
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
      SpreadDesigner  =   "Frm_PackList_Label.frx":0000
      TextTip         =   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LotID"
      Height          =   195
      Left            =   2640
      TabIndex        =   9
      Top             =   240
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Txt路径："
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   8280
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "系统中的主批号:"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1320
   End
End
Attribute VB_Name = "Frm_PackList_Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
    E_ID = 0                 'Id
    E_LotId                  'LotId
    E_Qty                    '料号
     E_OK
    E_End
    
End Enum

Dim listRS        As New ADODB.Recordset


Private Sub ChkAll_Click()
Dim i As Integer
    If ChkAll.Value = 1 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .Text = 1
            End With
        Next i
        
    ElseIf ChkAll.Value = 0 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .Text = 0
            End With
        Next i
        
    End If

End Sub

Private Sub CmdExit_Click()
TxtWaferID.Text = ""
TxtWaferID.SetFocus
End Sub



Private Sub IniFpsHeader()
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
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
  
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        
             .Col = E_FPS0.E_OK
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
          
        .SetText E_FPS0.E_ID, 0, "序号"
        .SetText E_FPS0.E_LotId, 0, "WaferID"
        .SetText E_FPS0.E_Qty, 0, "数量"
        .SetText E_FPS0.E_OK, 0, ""
        
        
        .ColWidth(E_FPS0.E_ID) = 5
        .ColWidth(E_FPS0.E_LotId) = 15
        .ColWidth(E_FPS0.E_Qty) = 15
        
        .ColWidth(E_FPS0.E_OK) = 10
     

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        

        .Col = E_FPS0.E_OK
        .Lock = False
        
        

        
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub CmdOK_Click()
''把资料生成一个txt
'Dim txtStr As String
'Dim dirTemp As String
'
'
'Dim fileNameTemp As String
'Dim msgTxtTemp As String
'fileNameTemp = ""
'msgTxtTemp = ""
'
'txtStr = TxtWaferID.Text
'
'msgTxtTemp = Replace(txtStr, vbCrLf, ",")
'
'fileNameTemp = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1)
'
'dirTemp = TxtDir.Text
'
'
'Call addLabelTxt(fileNameTemp, msgTxtTemp, dirTemp)
'TxtWaferID.Text = ""
'TxtWaferID.SetFocus


Dim i As Integer
Dim waferid As String
Dim qtyTemp As Long
Dim strSql As String
Dim waferIdOI As String
Dim productTemp As String
Dim packNoTemp As String
Dim dirtemp As String
Dim cmdStr2 As String
Dim flagTemp As Boolean



dirtemp = TxtDir.Text
flagTemp = False



   For i = 1 To fps(0).MaxRows
        With fps(0)
            .Row = i
            .Col = E_FPS0.E_OK
            If .Text = "1" Then
                .Row = i
                .Col = E_FPS0.E_LotId
                waferid = Mid(.Text, 1, InStr(.Text, "-") - 1)
                waferIdOI = waferid
                
                If flagTemp = False Then
                productTemp = GetGTTxtProduct(waferIdOI)
                packNoTemp = GetGTTxtPackNo(waferIdOI)
                
                flagTemp = True
                
                End If
                
                
                
                
                .Row = i
                .Col = E_FPS0.E_Qty
                qtyTemp = CLng(.Text)
                
                
                cmdStr2 = " insert into  tsv_qboxnumber_GT (qboxnumber,waferscribenumber,qty) values('" & packNoTemp & "', '" & waferid & "'," & qtyTemp & " )"

                AddSql (cmdStr2)
                
                strSql = strSql & "," & waferid & "," & qtyTemp
            End If
        End With
    Next i
    
    strSql = productTemp & "," & packNoTemp & strSql


  Call addLabelTxt(waferIdOI, strSql, dirtemp)


 MsgBox "此笔Txt已生成，请确认标签是否正常 ！"




End Sub

Private Sub GetFpsData()
Dim i As Integer
Set listRS = GetGT5271()
If (listRS.RecordCount > 0) Then

fps(0).MaxRows = listRS.RecordCount


For i = 0 To listRS.RecordCount - 1

  With fps(0)
         .Row = i + 1
         .Col = E_FPS0.E_ID
         .Text = i + 1
         
        .Row = i + 1
         .Col = E_FPS0.E_LotId
        .Text = listRS.fields(0).Value
        
        
         .Row = i + 1
         .Col = E_FPS0.E_Qty
        .Text = CStr(listRS.fields(1).Value)
        
         .Row = i + 1
         .Col = E_FPS0.E_OK
        .Text = CStr(listRS.fields(2).Value)
        
        
   End With
    
listRS.MoveNext

Next
End If
End Sub


Private Sub GetFpsDataWhere(lotIDtemp As String)
Dim i As Integer
Set listRS = GetGT5271Where(lotIDtemp)
If (listRS.RecordCount > 0) Then

fps(0).MaxRows = listRS.RecordCount


For i = 0 To listRS.RecordCount - 1

  With fps(0)
         .Row = i + 1
         .Col = E_FPS0.E_ID
         .Text = i + 1
         
        .Row = i + 1
         .Col = E_FPS0.E_LotId
        .Text = listRS.fields(0).Value
        
        
         .Row = i + 1
         .Col = E_FPS0.E_Qty
        .Text = CStr(listRS.fields(1).Value)
        
         .Row = i + 1
         .Col = E_FPS0.E_OK
        .Text = CStr(listRS.fields(2).Value)
        
        
   End With
    
listRS.MoveNext

Next
End If
End Sub


Private Sub CmdQuery_Click()
Dim lotIDtemp As String

lotIDtemp = UCase(Trim(TxtLotid.Text))



Dim i As Integer
Set listRS = GetGT5271Where(lotIDtemp)
If (listRS.RecordCount > 0) Then

fps(0).MaxRows = listRS.RecordCount


For i = 0 To listRS.RecordCount - 1

  With fps(0)
         .Row = i + 1
         .Col = E_FPS0.E_ID
         .Text = i + 1
         
        .Row = i + 1
         .Col = E_FPS0.E_LotId
        .Text = listRS.fields(0).Value
        
        
         .Row = i + 1
         .Col = E_FPS0.E_Qty
        .Text = CStr(listRS.fields(1).Value)
        
         .Row = i + 1
         .Col = E_FPS0.E_OK
        .Text = CStr(listRS.fields(2).Value)
        
        
   End With
    
listRS.MoveNext

Next
End If



End Sub

Private Sub Command1_Click()
'导出报表



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


'  sql1 = " select b.wafersequence, b.ordername,b.waferid,b.completeflag,b.dieqty, b.FGDieQty , b.WaferLot, b.MarkingCode " & _
'"from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername"
'
'
          
          
 sql1 = "select distinct to_char(a.create_date,'YYYY-MM-DD') outdate ,b.ordername  ,p.alternatename,pb.productname,d.firstname,b.waferlot,a.waferscribenumber," & _
        " b.waferlot||'-'||substr(a.waferscribenumber,length(a.waferscribenumber)-1) waferid2,e.passbincount+e.failbincount designqty,a.qty" & _
        " from  tsv_qboxnumber_GT a , ib_waferlist b  ,   ib_wohistory c, mappingdatatest e, container d, product p , productbase pb " & _
       " Where b.waferid = a.waferscribenumber And c.OrderName = b.OrderName And p.productbaseid = pb.productbaseid " & _
        " and pb.productname=c.product and d.containername=a.waferscribenumber||'-A' and e.substrateid=a.waferscribenumber"
 
 
          
sql3 = "  order by b.ordername,d.firstname,b.waferlot,a.waferscribenumber  "

  

  

  
  sql2 = " and  a.create_date>=to_date('" + beginTime + "','YYYY/MM/DD') and  a.create_date<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
   sqlTemp = sql1 & sql2 & sql3
  

  
  
  
     ExporToExcel (sqlTemp)









End Sub

Private Sub Form_Load()

DTP1.Value = Now - 1

DTP2.Value = Now

TxtDir.Text = "\\10.160.1.14\BarCode\SI\SIOUT\"




IniFpsHeader

GetFpsData



End Sub
