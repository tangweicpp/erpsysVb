VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_HY_WaferId_Label 
   Caption         =   "PackList WaferID标签 GC_G004N"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
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
   ScaleHeight     =   11010
   ScaleWidth      =   20280
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "报表(WLA的模板)"
      Height          =   6855
      Left            =   10440
      TabIndex        =   6
      Top             =   480
      Width           =   4575
      Begin VB.Frame Frame3 
         Caption         =   "走ERP发货的"
         Height          =   3015
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   4335
         Begin VB.CommandButton Command2 
            Caption         =   "导出Excel"
            Height          =   480
            Left            =   1560
            TabIndex        =   18
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox TxtBillNoGC 
            Height          =   375
            Left            =   720
            TabIndex        =   16
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单据编号："
            Height          =   195
            Left            =   720
            TabIndex        =   17
            Top             =   360
            Width           =   900
         End
      End
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
      Text            =   "Frm_HY_WaferId_Label.frx":0000
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
Attribute VB_Name = "Frm_HY_WaferId_Label"
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
Dim qboxNoTemp As String
Dim qboxNoContainerTemp As String



Dim sqlDB As String

fileNameTemp = ""
msgTxtTemp = ""

txtStr = TxtWaferID.Text

msgTxtTemp = Replace(txtStr, vbCrLf, "','")

''1234,'456,'789'
msgTxtTemp2 = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & "," & Right(msgTxtTemp, Len(msgTxtTemp) - InStr(msgTxtTemp, ","))


Dim bid
bid = Split(Replace(msgTxtTemp2, "'", "") & ",", ",")

Dim lotStr As String

For i = 0 To UBound(bid) - 1
    lotStr = bid(i)

    
    If i = 0 Then
    qboxNoContainerTemp = "WLC" & lotStr & "-A"
    
    qboxNoTemp = GetWLAQbox(qboxNoContainerTemp)
    
    
    End If
    
     
    cmdStr2 = " insert into tsv_qboxnumber_GC select containerid,containername,qty,sysdate,'" & qboxNoTemp & "'  from container where containername= '" & lotStr & "' "
      
    AddSql (cmdStr2)
     
Next i


sqlDB = GetGC400NString(msgTxtTemp2, qboxNoTemp)


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
Dim sqlTemp As String
Dim cusPTTemp As String




beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")
cusPTTemp = CusPT.Text

 
  sqlTemp = " select  row_number() over(order by 1) as ""No."" , X.SubName as ""Sub Name"",X.ShipTo as ""Ship To"",X.CustomerDevice as ""Customer Device"",X.GCVersion as ""GC Version"",X.CSTID as ""CST ID"",X.CSTQTY as ""CST QTY"",X.BondPro as ""Bond Pro."",X.FABLotID as ""FAB Lot ID"",X.WaferID as ""Wafer ID"",X.GrossDies as ""Gross Dies"",X.PONO as ""PO NO"",X.WO as ""WO"",X.InvoiceNO as ""Invoice NO"",X.FABDevice as ""FAB Device"",X.PacklotID as ""Pack lot ID"",X.FABOutDate as ""FAB-Out Date"", " & _
 " X.SamplingQty as ""Sampling Qty"",X.PassDies as ""Pass Dies"",X.Yield as ""Yield"",X.Remark as ""Remark""  from ( " & _
 " select distinct 'HTKS' as SubName, 'GC_LG' as ShipTo, replace(a.mpn_desc,'-3','-2.5') as CustomerDevice, a.imager_customer_rev as GCVersion, " & _
        "   Get_GCWLA_LotID(b.lotid,b.substrateid,to_date('" + beginTime + "','YYYY/MM/DD'),to_date('" + endTime + "','YYYY/MM/DD'),'" + cusPTTemp + "') as CSTID,   Get_GCWLA_Qty(b.lotid,b.substrateid,to_date('" + beginTime + "','YYYY/MM/DD'),to_date('" + endTime + "','YYYY/MM/DD'),'" + cusPTTemp + "') as CSTQTY, 'SH' as BondPro, b.lotid as FABLotID,  substr(b.substrateid,length(b.substrateid)-1,2) as WaferID, b.passbincount as GrossDies, " & _
        " a.po_num as PONO,a.mtrl_num as WO,  '' InvoiceNO, a.fab_conv_id as FABDevice, c.firstname as PacklotID,to_char(sysdate, 'YYYY-MM-DD') as FABOutDate, " & _
        " b.passbincount as SamplingQty,  '' as PassDies, '' as Yield, '' as Remark " & _
        " from  tsv_qboxnumber_GC d, mappingdatatest b, customeroitbl_test a, container c " & _
        " Where d.create_date >=to_date('" + beginTime + "','YYYY/MM/DD') and  d.create_date <to_date('" + endTime + "','YYYY/MM/DD') and b.customershortname = 'GC' and b.substrateid =d.waferscribenumber and b.filename = a.id " & _
        " and a.customershortname = 'GC' and c.containername = b.substrateid and a.mpn_desc='" + cusPTTemp + "'  " & _
        " order by   b.lotid,  9 ) X"

 
     ExporToExcel (sqlTemp)



End Sub

Private Sub Command2_Click()
'ERP的导出


Dim billNoTemp As String

 billNoTemp = Trim(TxtBillNoGC.Text)
  
      sqlTemp = "  SELECT row_number() OVER(ORDER BY a.工单号,a.流程卡编号) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
" replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID]," & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST QTY]," & _
" 'SH' as [Bond Pro.],a.工单号 as [FAB Lot ID],right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID]," & _
" a.数量 as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.数量 as [Sampling Qty]," & _
" ''as [Pass Dies],''as [Yield],''as [Remark] " & _
" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c WHERE a.单据编号='" + billNoTemp + "'" & _
" and b.SOURCE_BATCH_ID=a.工单号 and c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 "
        
        
        
     SqlServerExporToExcel (sqlTemp)


End Sub

Private Sub Form_Activate()
TxtWaferID.SetFocus
End Sub

Private Sub Form_Load()
TxtWaferID.Text = ""
TxtDir.Text = "\\10.160.1.14\BarCode\GCWLA\"

DTP1.Value = Now - 1

DTP2.Value = Now

CusPT.AddItem ("GC0310-3")
CusPT.AddItem ("GC0312-3")
CusPT.AddItem ("GC6123-3")


End Sub
