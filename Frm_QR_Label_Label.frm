VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_QR_Label 
   Caption         =   "QR标签"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14175
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
   ScaleHeight     =   8430
   ScaleWidth      =   14175
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "报表(QR的模板)"
      Height          =   6855
      Left            =   10440
      TabIndex        =   6
      Top             =   480
      Width           =   4575
      Begin VB.Frame Frame2 
         Caption         =   "每一次出去的"
         Height          =   2655
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4335
         Begin VB.CommandButton Command1 
            Caption         =   "导出Excel"
            Height          =   480
            Left            =   1560
            TabIndex        =   9
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox CusPT 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DTP1 
            Height          =   375
            Left            =   2040
            TabIndex        =   10
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   176029697
            CurrentDate     =   41424
         End
         Begin MSComCtl2.DTPicker DTP2 
            Height          =   375
            Left            =   2040
            TabIndex        =   11
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   176029697
            CurrentDate     =   41424
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结束日期："
            Height          =   195
            Left            =   840
            TabIndex        =   14
            Top             =   1560
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开始日期： "
            Height          =   195
            Left            =   840
            TabIndex        =   13
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "客户机种： "
            Height          =   195
            Left            =   840
            TabIndex        =   12
            Top             =   360
            Visible         =   0   'False
            Width           =   945
         End
      End
   End
   Begin VB.TextBox TxtDir 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   6855
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "取消"
      Height          =   480
      Left            =   5400
      TabIndex        =   3
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定"
      Height          =   480
      Left            =   3000
      TabIndex        =   2
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox TxtWaferID 
      Height          =   6615
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Frm_QR_Label_Label.frx":0000
      Top             =   600
      Width           =   9255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Txt路径："
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "扫入的WaferID:"
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Frm_QR_Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
TxtWaferID.Text = ""
TxtWaferID.SetFocus
End Sub

Private Sub CmdOK_Click()
'把资料生成一个txt
Dim txtStr As String
Dim dirtemp As String
Dim cmdStr2 As String

Dim fileNameTemp As String
Dim msgTxtTemp As String
Dim msgTxtTemp2 As String
Dim txtStrTemp3 As String


Dim sqlDB As String

fileNameTemp = ""
msgTxtTemp = ""
txtStrTemp3 = ""

txtStr = TxtWaferID.Text

msgTxtTemp = Replace(txtStr, vbCrLf, "','")

''1234,'456,'789'
msgTxtTemp2 = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & "," & Right(msgTxtTemp, Len(msgTxtTemp) - InStr(msgTxtTemp, ","))


Dim bid
bid = Split(Replace(msgTxtTemp2, "'", "") & ",", ",")

Dim lotStr As String

For i = 0 To UBound(bid) - 1
    lotStr = bid(i)
    
    txtStrTemp3 = txtStrTemp3 & lotStr & "-A" & "','"
    
    cmdStr2 = " insert into tsv_qboxnumber_WaiBao_QR select containerid,containername,qty,sysdate from container where containername= '" & lotStr & "-A' "
      
    AddSql (cmdStr2)
     
Next i

Dim testtemp01 As String

testtemp01 = msgTxtTemp2

Dim testtemp02 As String

testtemp02 = txtStrTemp3

txtStrTemp3 = Left(txtStrTemp3, Len(txtStrTemp3) - 8)


sqlDB = GetWaiBaoQRString(txtStrTemp3)


fileNameTemp = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1)

dirtemp = TxtDir.Text


Call addLabelTxt(fileNameTemp, sqlDB, dirtemp)
TxtWaferID.Text = ""
TxtWaferID.SetFocus

End Sub

Private Sub Command1_Click()

Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqltemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""

beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

   
 sql1 = "select distinct to_char(a.create_date,'YYYY-MM-DD') outdate ,b.ordername  ,p.alternatename,pb.productname,d.firstname,b.waferlot, replace(a.waferscribenumber,'-A','') as waferid," & _
        "   b.waferlot||'-'|| substr( replace(a.waferscribenumber,'-A',''),-2) as waferid2 ,e.passbincount+e.failbincount designqty,a.qty" & _
        " from  tsv_qboxnumber_WaiBao_QR a , ib_waferlist b  ,   ib_wohistory c, mappingdatatest e, container d, product p , productbase pb " & _
       " Where b.waferid = replace(a.waferscribenumber,'-A','') And c.OrderName = b.OrderName And p.productbaseid = pb.productbaseid " & _
        " and pb.productname=c.product and d.containername=a.waferscribenumber and e.substrateid=replace(a.waferscribenumber,'-A','') "
 
 
sql3 = "  order by b.ordername,d.firstname,b.waferlot,waferid "

  
  sql2 = " and  a.create_date>=to_date('" + beginTime + "','YYYY/MM/DD') and  a.create_date<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
   sqltemp = sql1 & sql2 & sql3
   
   

  
     ExporToExcel (sqltemp)


End Sub

Private Sub Command2_Click()
'ERP的导出


Dim billNoTemp As String

 billNoTemp = Trim(TxtBillNoGC.Text)
  
      sqltemp = "  SELECT row_number() OVER(ORDER BY a.工单号,a.流程卡编号) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
" replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID]," & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST QTY]," & _
" 'SH' as [Bond Pro.],a.工单号 as [FAB Lot ID],right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID]," & _
" a.数量 as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.数量 as [Sampling Qty]," & _
" ''as [Pass Dies],''as [Yield],''as [Remark] " & _
" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c WHERE a.单据编号='" + billNoTemp + "'" & _
" and b.SOURCE_BATCH_ID=a.工单号 and c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 "
        
        
        
     SqlServerExporToExcel (sqltemp)


End Sub

Private Sub Form_Activate()
TxtWaferID.SetFocus
End Sub

Private Sub Form_Load()
TxtWaferID.Text = ""
TxtDir.Text = "\\10.160.1.14\BarCode\QR\QRCP\"


'TxtDir.Text = "C:\WAIBAOSAW\"

DTP1.Value = Now - 1

DTP2.Value = Now

CusPT.AddItem ("GC0310-3")
CusPT.AddItem ("GC0312-3")
CusPT.AddItem ("GC6123-3")


End Sub
