VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmCustomosReport 
   Caption         =   "关务报表"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16245
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
   ScaleHeight     =   10770
   ScaleWidth      =   16245
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "报表明细"
      ForeColor       =   &H00800000&
      Height          =   8535
      Left            =   0
      TabIndex        =   8
      Top             =   2160
      Width           =   16215
      Begin FPSpreadADO.fpSpread fpS 
         Height          =   8175
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   15615
         _Version        =   524288
         _ExtentX        =   27543
         _ExtentY        =   14420
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
         SpreadDesigner  =   "FrmCustomosReport.frx":0000
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "菜单选项"
      ForeColor       =   &H00800000&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16215
      Begin VB.TextBox txtDNList 
         BackColor       =   &H00FFC0FF&
         Height          =   1935
         Left            =   6840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   120
         Width           =   1575
      End
      Begin VB.ComboBox cboExportFileFormat 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "FrmCustomosReport.frx":044C
         Left            =   1200
         List            =   "FrmCustomosReport.frx":044E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1035
         Width           =   1335
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00808080&
         Caption         =   "退出"
         Height          =   360
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   990
      End
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00C0FFFF&
         Caption         =   "导出"
         Height          =   360
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1560
         Width           =   990
      End
      Begin VB.CommandButton cmdRead 
         BackColor       =   &H00FFC0C0&
         Caption         =   "查询"
         Height          =   360
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   990
      End
      Begin VB.ComboBox cboReportName 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "FrmCustomosReport.frx":0450
         Left            =   1200
         List            =   "FrmCustomosReport.frx":0452
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   690
         Width           =   3255
      End
      Begin VB.ComboBox cboCustCode 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "FrmCustomosReport.frx":0454
         Left            =   1200
         List            =   "FrmCustomosReport.frx":0456
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblDNList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN列表"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6120
         TabIndex        =   13
         Top             =   840
         Width           =   660
      End
      Begin VB.Label lblExportFileFormat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "导出格式"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblReportName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报表名称"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   735
         Width           =   900
      End
      Begin VB.Label lblCustCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   405
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmCustomosReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum E_37_PACKINGLIST

    E_CHOOSE = 1
    E_DOCUMENT_NO
    E_DN
    E_SHIP_DATE
    E_SHIPTONAME
    E_SHIPTOSTREET1
    E_SHIPTOSTREET2
    E_SHIPTOSTREET3
    E_CITY
    E_STATE
    E_POSTALCODE
    E_COUNTRYKEY
    E_CONTACTNAME
    E_PHONE
    E_SALESDOCUMENT
    E_PURCHASINGDOCNO
    E_BOXID
    E_PRODUCTID
    E_MPN
    E_QTY
    E_JOBID
    E_DATECODE
    E_HTLOTID
    E_CPN
    E_NET_WEIGHT
    E_GROSS_WEIGHT
    E_SIZE
    E_END
    
End Enum

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call InitData
Call InitCtrls
End Sub
'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitData
' Description:       初始化数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/26-11:35:59
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitData()

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitCtrls
' Description:       初始化控件
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/26-11:36:06
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitCtrls()
Call InitCustCode

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitCustCode
' Description:       初始化客户代码
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/26-11:37:26
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitCustCode()

cboCustCode.AddItem ("37")
cboCustCode.AddItem ("GC")
cboCustCode.AddItem ("68")
cboCustCode.Text = "37"
End Sub


'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cboCustCode_Click
' Description:       客户代码变更切换资料模板
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/26-11:42:03
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cboCustCode_Click()
If cboCustCode.Text = "" Then Exit Sub
cboReportName.Clear

Select Case cboCustCode.Text

    Case "37"
        cboReportName.AddItem ("PACKINGLIST")
        cboReportName.AddItem ("INVOICE")
    
End Select

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cboReportName_Click
' Description:       资料模板变更切换资料文件格式
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/26-11:49:50
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cboReportName_Click()
If cboReportName.Text = "" Then Exit Sub
cboExportFileFormat.Clear
fpS(0).MaxRows = 0
Select Case cboReportName.Text

    Case "PACKINGLIST", "INVOICE"
        cboExportFileFormat.AddItem ("xlsx")
        cboExportFileFormat.Text = "xlsx"
    
End Select

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cmdRead_Click
' Description:       查询资料
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/26-13:18:02
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdRead_Click()
If cboCustCode.Text = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub
End If
If cboReportName.Text = "" Then
    MsgBox "请选择报表名称", vbInformation, "提示"
    Exit Sub

End If

Select Case cboCustCode.Text

    Case "37"
        Select Case cboReportName.Text
        
            Case "PACKINGLIST"
                Call InitFps_37_PACKINGLIST
                Call LstFps_37_PACKINGLIST
            
            Case "INVOICE"
                Call InitFps_37_INVOICE
                Call LstFps_37_INVOICE
            
        End Select
    
    Case "GC"
    
    Case "68"

End Select

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitFps_37_PACKINGLIST
' Description:       初始化37PACKINGLIST表头
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/26-13:22:26
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitFps_37_PACKINGLIST()

With fpS(0)
    .ReDraw = False
    .MaxCols = E_37_PACKINGLIST.E_END - 1
    .MaxRows = 0
    .FontBold = False
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    '.TypeVAlign = TypeVAlignCenter
    '.TypeHAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    
    .Col = E_37_PACKINGLIST.E_CHOOSE
    .CellType = CellTypeCheckBox
    
    .SetText E_37_PACKINGLIST.E_CHOOSE, 0, "√"
    .SetText E_37_PACKINGLIST.E_DOCUMENT_NO, 0, "单据编号"
    .SetText E_37_PACKINGLIST.E_DN, 0, "DN"
    .SetText E_37_PACKINGLIST.E_SHIP_DATE, 0, "发货日期"
    .SetText E_37_PACKINGLIST.E_SHIPTONAME, 0, "SHIPTO"
    .SetText E_37_PACKINGLIST.E_SHIPTOSTREET1, 0, "SHIPTOSTREET1"
    .SetText E_37_PACKINGLIST.E_SHIPTOSTREET2, 0, "SHIPTOSTREET2"
    .SetText E_37_PACKINGLIST.E_SHIPTOSTREET3, 0, "SHIPTOSTREET3"
    .SetText E_37_PACKINGLIST.E_CITY, 0, "CITY"
    .SetText E_37_PACKINGLIST.E_STATE, 0, "STATE"
    .SetText E_37_PACKINGLIST.E_POSTALCODE, 0, "POSTALCODE"
    .SetText E_37_PACKINGLIST.E_COUNTRYKEY, 0, "COUNTRYKEY"
    .SetText E_37_PACKINGLIST.E_CONTACTNAME, 0, "CONTACTNAME"
    .SetText E_37_PACKINGLIST.E_PHONE, 0, "PHONE"
    .SetText E_37_PACKINGLIST.E_SALESDOCUMENT, 0, "SALESDOCUMENT"
    .SetText E_37_PACKINGLIST.E_PURCHASINGDOCNO, 0, "PURCHASINGDOCNO"
    .SetText E_37_PACKINGLIST.E_BOXID, 0, "箱号"
    .SetText E_37_PACKINGLIST.E_PRODUCTID, 0, "料号"
    .SetText E_37_PACKINGLIST.E_MPN, 0, "MPN"
    .SetText E_37_PACKINGLIST.E_QTY, 0, "数量"
    .SetText E_37_PACKINGLIST.E_JOBID, 0, "JOBID"
    .SetText E_37_PACKINGLIST.E_DATECODE, 0, "DATECODE"
    .SetText E_37_PACKINGLIST.E_HTLOTID, 0, "HTLOTID"
    .SetText E_37_PACKINGLIST.E_CPN, 0, "CPN"
    .SetText E_37_PACKINGLIST.E_NET_WEIGHT, 0, "净重"
    .SetText E_37_PACKINGLIST.E_GROSS_WEIGHT, 0, "毛重"
    .SetText E_37_PACKINGLIST.E_SIZE, 0, "尺寸"
    
    .ColWidth(E_37_PACKINGLIST.E_CHOOSE) = 4
    .ColWidth(E_37_PACKINGLIST.E_DOCUMENT_NO) = 12
    .ColWidth(E_37_PACKINGLIST.E_DN) = 12
    .ColWidth(E_37_PACKINGLIST.E_SHIP_DATE) = 8
    .ColWidth(E_37_PACKINGLIST.E_SHIPTONAME) = 10

    .ReDraw = True

End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       LstFps_37_PACKINGLIST
' Description:       显示报表结果
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/26-13:42:57
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LstFps_37_PACKINGLIST()
Dim strDNList As String
Dim strsql As String
Dim rsPackList As New ADODB.Recordset

strDNList = GetDNList
If strDNList = "" Then
    MsgBox "请输入需要查询的正确DN", vbInformation, "提示"
    Exit Sub
End If

strsql = "SELECT 1 AS  选择,y.代码简称 + b.单据编号 AS 单据编号,d.Delivery, CONVERT(VARCHAR(100), c.单据日期,23) AS 发货日期, ISNULL(dn_address_new, d.shiptoname) AS shiptoname, ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1 " & _
 " ,  ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2, ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,  ISNULL(x.city_new, d.city) AS city,  ISNULL(x.dn_st_new, d.State) AS state,  ISNULL(x.postal_code_new, d.postalcode) AS postalcode, " & _
 "  ISNULL(x.country_new, d.countrykey) AS countrykey,  ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone, d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo " & _
 " ,f.箱号 ,b.料号 ,d.MarketingPN,SUM(b.数量),d.BatchNumber,d.DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no, d.customerPartNumber, ROUND(CAST(f.重量 AS FLOAT) * 0.4, 2) 净重, f.重量 毛重, " & _
 " f.尺寸 FROM erpdata..tblStockNumTree a INNER JOIN erpdata..tblStocksqfhsub b ON b.箱号 = a.箱号 INNER JOIN erpdata..tblStockSQfh c ON c.单据编号 = b.单据编号 AND c.序号 = b.单据项次 INNER JOIN erpdata..tblstock y " & _
 " ON y.库房代码 = c.仓库编号  INNER JOIN (SELECT a.BOX_ID,  SUBSTRING(a.KEY_VALUE,  1, CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job " & _
 " FROM erpdata .. tblErpInStockDetailInfo a  WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.箱号 = aa.qbox  INNER JOIN (SELECT dn.Delivery, dn.shiptoname, dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3 " & _
 " , dn.city, dn.State, dn.postalcode, dn.countrykey, dn.contactname,  dn.phone, dn.SalesDocument,  dn.PurchasingDocNo, dn.BatchNumber, SUM(dn.Quantity) AS Quantity, dn.customerPartNumber, dn.MarketingPN, dn.DATE_CODE " & _
 " FROM ERPBASE..tblCustomerShippingUp dn WHERE dn.Delivery IN ('" & strDNList & "') GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1, dn.shiptostreet2, dn.shiptostreet3,dn.city, dn.State, dn.postalcode, dn.countrykey " & _
 " , dn.contactname,dn.phone, dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber, dn.customerPartNumber, dn.MarketingPN,dn.DATE_CODE) d  ON d.Delivery = a.DN AND d.BatchNumber = aa.job " & _
 " INNER JOIN erpdata..tblStockNumTree e ON e.箱号 = b.箱号 INNER JOIN erpdata..tblstocknumtree f  ON f.序号 = e.上级序号 LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName " & _
"  WHERE a.DN IN ('" & strDNList & "')GROUP BY  b.单据编号,c.单据日期,ISNULL(dn_address_new, d.shiptoname) , ISNULL(x.ship_to_street1_new, d.shiptostreet1) ,ISNULL(x.ship_to_street2_new, d.shiptostreet2) " & _
 " ,ISNULL(x.ship_to_street3_new, d.shiptostreet3),ISNULL(x.city_new, d.city),ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode), ISNULL(x.country_new, d.countrykey) " & _
 " ,ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) , d.SalesDocument,d.PurchasingDocNo ,f.箱号,b.料号 ,d.MarketingPN,d.BatchNumber ,d.DATE_CODE " & _
 " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber, f.重量 , f.尺寸,y.代码简称,d.Delivery "

Set rsPackList = Get_SqlserveRs(strsql)
If rsPackList.RecordCount = 0 Then
    MsgBox "查询不到该DN的资料", vbInformation, "提示"
    rsPackList.Close
    Set rsPackList = Nothing
    Exit Sub
End If

With fpS(0)
    .MaxRows = 0
    If rsPackList.RecordCount > 0 Then
        Set .DataSource = rsPackList
    End If

End With

rsPackList.Close
Set rsPackList = Nothing

End Sub

Private Function GetDNList() As String
Dim strDNList As String
Dim strDNArray() As String
Dim strDN As String
Dim i As Integer

If Len(Trim(txtDNList.Text)) = 0 Then
    GetDNList = ""
    Exit Function
End If

strDNArray = Split(Trim(txtDNList.Text), vbCrLf)

For i = 0 To UBound(strDNArray)
    strDN = Trim("" & strDNArray(i))
    If strDN <> "" And InStr(strDNList, strDN) = 0 Then
        strDNList = strDNList & strDN & "','"
    End If
    
Next

strDNList = Mid(strDNList, 1, Len(strDNList) - 3)
GetDNList = strDNList

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitFps_37_INVOICE
' Description:       初始化37INVOICE表头
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/26-13:22:56
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitFps_37_INVOICE()


End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       LstFps_37_INVOICE
' Description:       显示报表结果
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/26-13:42:42
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LstFps_37_INVOICE()



End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cmdExport_Click
' Description:       导出报表
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/27-13:48:13
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdExport_Click()
Select Case cboCustCode.Text

    Case "37"
        Select Case cboReportName.Text
        
            Case "PACKINGLIST"
                Call ExportRep_37_PACKINGLIST
            
            Case "INVOICE"
       
            
        End Select
    
    Case "GC"
    
    Case "68"

End Select

End Sub

Private Sub ExportRep_37_PACKINGLIST()
Dim xlApp                 As Excel.Application
Dim xlBook                As Excel.Workbook
Dim xlSheet               As Excel.Worksheet
Dim strTemplateExFileName As String
Dim strTmp1               As String
Dim strTmp2               As String
Dim strTmp3               As String
Dim strTmp4               As String

On Error GoTo hErr

If fpS(0).MaxRows = 0 Then
    MsgBox "没有资料可以导出", vbInformation, "提示"
    Exit Sub

End If

strTemplateExFileName = App.Path & "\NewSemtechReport\shipping_packing_list.xlsx"
Set xlApp = CreateObject("excel.application")
xlApp.Visible = True
Set xlBook = xlApp.Workbooks.Open(strTemplateExFileName)
Set xlSheet = xlBook.Worksheets(1)

'赋值开始
With fpS(0)
    .Row = 1
    .Col = E_37_PACKINGLIST.E_SHIP_DATE
    xlSheet.Cells(8, 17) = Trim$("" & .Text)
    .Col = E_37_PACKINGLIST.E_SHIPTONAME
    xlSheet.Cells(17, 2) = Trim$("" & .Text)
    .Col = E_37_PACKINGLIST.E_SHIPTOSTREET1
    xlSheet.Cells(18, 2) = Trim$("" & .Text)
    .Col = E_37_PACKINGLIST.E_SHIPTOSTREET2
    strTmp1 = Trim$("" & .Text)
    .Col = E_37_PACKINGLIST.E_SHIPTOSTREET3
    strTmp2 = Trim$("" & .Text)
    xlSheet.Cells(19, 2) = strTmp1 & " " & strTmp2
    .Col = E_37_PACKINGLIST.E_CITY
    strTmp1 = Trim$("" & .Text)
    .Col = E_37_PACKINGLIST.E_STATE
    strTmp2 = Trim$("" & .Text)
    .Col = E_37_PACKINGLIST.E_POSTALCODE
    strTmp3 = Trim$("" & .Text)
    .Col = E_37_PACKINGLIST.E_COUNTRYKEY
    xlSheet.Cells(17, 17) = Trim$("" & .Text)
    strTmp4 = Trim$("" & .Text)
    xlSheet.Cells(20, 2) = strTmp1 & " " & strTmp2 & " " & strTmp3 & " " & strTmp4
    .Col = E_37_PACKINGLIST.E_CONTACTNAME
    strTmp1 = Trim$("" & .Text)
    .Col = E_37_PACKINGLIST.E_PHONE
    strTmp2 = Trim$("" & .Text)
    xlSheet.Cells(22, 2) = "Attn:" & strTmp1 & " ,Tel:" & strTmp2
    
     .Col = E_37_PACKINGLIST.E_SALESDOCUMENT
'    xlSheet.Cells(25, 3) = Trim$("" & rs.Fields(14).Value)
'    xlSheet.Cells(25, 6) = Trim$("'" & rs.Fields(15).Value)
    
    
    
    
    
    
    

End With

hErr:
If Not xlApp Is Nothing Then
    xlApp.DisplayAlerts = False '关闭时不提示保存
    xlBook.Close
    xlApp.Quit
    '释放资源
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing

End If

'Dim strSql            As String
'Dim lngRows           As Long
'Dim rsQuery           As EXCEL.QueryTable
'Dim xlApp             As EXCEL.Application
'Dim wkbk              As New Workbook
'Dim wkst              As New Worksheet
'Dim i                 As Long
'Dim J                 As Long
'Dim IntCols           As Integer
'Dim strCols           As String
'Dim StrFileName       As String
'Dim IntInertRow       As Integer, IntMaxDetailRow As Integer
'Dim DblNum            As Double
'Dim DblAmt            As Double  '总金额
'Dim intBoxNum         As Integer '箱数
'Dim strPBigBox        As String  '前箱号
'Dim strNBigBox        As String  '新箱号
'Dim IntBMegerRow      As Integer
'Dim IntEMegerRow      As Integer
'Dim DblJZ             As Double   '净重
'Dim DblMZ             As Double   '毛重
'Dim DblJZ1            As Double   '净重
'Dim DblMZ1            As Double   '毛重
'Dim DblJZ2            As Double   '净重
'Dim DblMZ2            As Double   '毛重
'Dim intBegin          As Integer
'Dim strdjTmp          As String
'Dim SD                As String
'Dim SD1               As String
'Dim strTmp()          As String
'Dim strExtsion        As String '后缀名
'Dim strNewFullPath    As String '新Excel文件
'Dim strNewFullPathNew As String
'Dim RsNew             As New adodb.Recordset '记录大箱的个数，方便后面计算体积重
'Dim rs                As New adodb.Recordset
'Dim dnnum             As String
'Dim dnnum1            As String
'
'strNewFullPathNew = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\excel"
'dnnum = ""
'dnnum1 = ""
'strPBigBox = ""
'strNBigBox = ""
'strdjTmp = ""
'intBoxNum = 1
'
'StrFileName = DirShare & "\shipping_packing_list.xlsx" '要打开的文件
'
'strSql = " select * from ( SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & ",箱号,料号,replace(mpn_desc,'.P2','') AS mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,净重,毛重,尺寸 " & " FROM Vw_InvShippedPLFor37_NEW a  where 单据编号 in ('" & Ordertemp & "')  " & " union all " & "SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & ",箱号,料号,replace(mpn_desc,'.P2','') AS mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,净重,毛重,尺寸 " & " FROM Vw_InvShippedPLFor37 a  where 单据编号 in ('" & Ordertemp & "') ) x order by x.箱号 "
'Dim strDNList As String
'
'strDNList = Mid(strDNList, 1, Len(strDNList) - 3)
'strSql = "SELECT 0 选择,h.代码简称 + a.单据编号 单据编号,c.delivery,dbo.usp_date(a.单据日期) 发货日期,ISNULL(dn_address_new, c.shiptoname) AS shiptoname,ISNULL(x.ship_to_street1_new, c.shiptostreet1) AS shiptostreet1, " & _
'   "ISNULL(x.ship_to_street2_new, c.shiptostreet2) AS shiptostreet2,ISNULL(x.ship_to_street3_new, c.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, c.city) AS city, ISNULL(x.dn_st_new, c.State) AS state, " & _
'   "ISNULL(x.postal_code_new, c.postalcode) AS postalcode,ISNULL(x.country_new, c.countrykey) AS countrykey,ISNULL(x.contact_new, c.contactname) AS contactname,ISNULL(x.phone_new, c.phone) AS phone, " & _
'   "c.SalesDocument,'''' + c.PurchasingDocNo AS PurchasingDocNo,RTRIM(erpdata.dbo.f_getparent(b.箱号)) 箱号,b.料号, " & _
'   "CASE WHEN RTRIM(gg.MPN_DESC) = 'UCLAMP0541Z.TFT' THEN RTRIM(gg.MPN_DESC) + '.P2' ELSE REPLACE(REPLACE(gg.MPN_DESC, '.P2', ''), '.P3', '') END AS mpn_desc,SUM(b.数量) 数量,c.batchnumber, " & _
'   "hh.CREATE_DATE DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no,c.customerPartNumber,ROUND(CAST(f.重量 AS FLOAT) * 0.4, 2) 净重,f.重量 毛重,f.尺寸 FROM erpdata .. tblStockSQfh a " & _
'   "INNER JOIN erpdata .. tblStocksqfhsub b ON a.单据编号 = b.单据编号 AND a.序号 = b.单据项次 INNER JOIN erpdata .. tblStockNumTree g ON g.箱号 = b.箱号 INNER JOIN erpdata .. tblStockNumTree f " & _
'   "ON f.序号 = g.上级序号 INNER JOIN (SELECT a.BOX_ID,SUBSTRING(a.KEY_VALUE,1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox,SUBSTRING(a.KEY_VALUE,CHARINDEX('|', a.KEY_VALUE) + 1,10) AS job " & _
'   "FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON g.箱号 = aa.qbox INNER JOIN (SELECT dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, " & _
'   "dn.shiptostreet3, dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber " & _
'   "FROM tblCustomerShippingUp dn GROUP BY dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname, " & _
'   "dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,dn.customerPartNumber) c ON c.Delivery = g.DN AND c.BatchNumber = aa.job INNER JOIN dbo.tblstock h " & _
'   "ON CONVERT(NVARCHAR(4), h.库房代码) = CONVERT(NVARCHAR(4), a.仓库编号) INNER JOIN ERPBASE .. tblmappingData ff ON ff.SUBSTRATEID = b.流程卡编号 " & _
'   "INNER JOIN ERPBASE .. tblCustomerOI gg ON CONVERT(VARCHAR(100), gg.ID) = ff.FILENAME AND gg.SOURCE_BATCH_ID = ff.LOTID AND gg.CUSTOMERSHORTNAME = '37' INNER JOIN erpbase .. weight37 hh " & _
'   "ON hh.WAFERID = REPLACE(b.流程卡编号, '+', '') INNER JOIN erpdata .. tblErpInStockMainData jj ON jj.BOX_ID = aa.BOX_ID LEFT JOIN erptemp .. dn_address x ON dn_address = c.ShipToName " & _
'   "WHERE a.客户代码 = '37' AND a.单据编号 LIKE 'F%' AND a.单据日期 >= CONVERT(VARCHAR(100), GETDATE() - 5, 23) AND c.Delivery = g.DN " & _
'   "and c.Delivery in ('" & strDNList & "') GROUP BY h.代码简称, a.单据编号, c.delivery,dbo.usp_date(a.单据日期), c.shiptoname,c.shiptostreet1,c.shiptostreet2,c.shiptostreet3,c.city,c.State,c.postalcode, " & _
'   "c.countrykey,c.contactname,c.phone,c.SalesDocument,c.PurchasingDocNo,erpdata.dbo.f_getparent(b.箱号),b.料号,gg.MPN_DESC,c.batchnumber,hh.CREATE_DATE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2), " & _
'   "c.customerPartNumber,f.重量,f.尺寸,dn_address_new,x.ship_to_street1_new,x.ship_to_street2_new,x.ship_to_street3_new,x.city_new,x.dn_st_new,x.postal_code_new,x.country_new,x.contact_new,x.phone_new order by Delivery"
'rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
'
'strExtsion = Mid$(StrFileName, InStrRev(StrFileName, "."))      '获取后缀名
'strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
'
'If rs.RecordCount > 0 Then
'
'    Set xlApp = New EXCEL.Application
'    xlApp.Visible = False   '是否显示
'    Set wkbk = xlApp.Workbooks.Open(StrFileName)
'    Set wkst = wkbk.Worksheets(1)
'
'    DblNum = 0
'    DblJZ = 0
'    DblMZ = 0
'    '
'    '赋值到Excel中，表头
'    wkst.Cells(22, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
'    wkst.Cells(23, 2) = ""
'    wkst.Cells(17, 17) = Trim$("" & rs.Fields(11).Value) 'To
'    wkst.Cells(25, 3) = Trim$("" & rs.Fields(14).Value)
'    wkst.Cells(25, 6) = Trim$("'" & rs.Fields(15).Value)
'    lngRows = 28
'    IntInertRow = rs.RecordCount * 2
'
'    For i = 1 To IntInertRow - 1
'        wkst.Rows(lngRows & ":" & lngRows).Select
'        xlApp.Selection.Copy
'        xlApp.Selection.Insert Shift:=xlDown
'        wkst.Rows(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '边框无
'    Next i
'
'    IntMaxDetailRow = rs.RecordCount
'
'    IntBMegerRow = 27
'    IntEMegerRow = 30
'    intBegin = 1
'    Dim QBX As String
'
'    For i = 0 To rs.RecordCount - 1
'        If dnnum1 <> Trim$("" & rs.Fields(2).Value) Then
'            dnnum = Trim$("" & rs.Fields(2).Value) + "/" + dnnum
'            dnnum1 = Trim$("" & rs.Fields(2).Value)
'
'        End If
'
'        strPBigBox = Trim$("" & rs.Fields(16).Value) '箱号
'
'        If strPBigBox <> strNBigBox Then
'            strNBigBox = Trim$("" & rs.Fields(16).Value) '箱号
'            '箱数
'            intBoxNum = intBoxNum + 1
'            wkst.Cells(lngRows, 1) = "K" & Trim(intBoxNum - 1) '箱号进行转换为客户所要资料
'            IntBMegerRow = IntBMegerRow + intBegin
'            intBegin = 1
'        Else
'
'            intBegin = intBegin + 1
'
'        End If
'
'        If SD <> Trim$("" & rs.Fields(14).Value) Then
'            SD = Trim$("" & rs.Fields(14).Value)
'            SD1 = SD1 & SD & " "
'
'        End If
'
'        wkst.Cells(25, 3) = SD1
'        wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(15).Value) 'PO
'        wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(17).Value)
'        wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(18).Value)
'        wkst.Cells(lngRows, 7) = Val(Trim$("" & rs.Fields(19).Value)) / 1000 '数量改为已千为单位
'        DblNum = DblNum + Val(Trim$("" & rs.Fields(19).Value))
'        wkst.Cells(lngRows, 9) = "KPCS"
'        wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(20).Value)
'        wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(21).Value) 'datacode
'        wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(22).Value) 'lotno
'        If strPBigBox <> QBX Then
'            wkst.Cells(lngRows, 14) = Trim$("" & rs.Fields(24).Value) '净重
'            wkst.Cells(lngRows, 15) = "KG"   '净重单位
'            wkst.Cells(lngRows, 18) = "KG"   '毛重单位
'            wkst.Cells(lngRows, 19) = Trim$("" & rs.Fields(26).Value)   '尺寸
'            wkst.Cells(lngRows, 17) = Trim$("" & rs.Fields(25).Value)   '毛重
'
'        End If
'
'        DblJZ1 = Val(Trim$("" & rs.Fields(24).Value))
'        If strPBigBox <> QBX Then
'            DblJZ = DblJZ1 + DblJZ
'
'        End If
'
'        DblMZ1 = Val(Trim$("" & rs.Fields(25).Value))
'        If strPBigBox <> QBX Then
'            DblMZ = DblMZ + DblMZ1
'
'        End If
'
'        '
'        lngRows = lngRows + 1
'        wkst.Cells(lngRows, 4) = "CPN:"
'        wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(23).Value)
'        QBX = strPBigBox
'        lngRows = lngRows + 1
'        IntEMegerRow = lngRows
'        rs.MoveNext
'    Next
'    '计算汇总
'    wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '总数量改为已千为单位
'    wkst.Cells(lngRows + 1, 9) = "KPCS" '单位
'    wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)    '箱数
'    wkst.Cells(lngRows + 1, 14) = Format(DblJZ, "0.00") '净重
'    wkst.Cells(lngRows + 1, 17) = DblMZ '毛重，记录它到后面进行对比
'Else
'    MsgBox "无需导出数据！", vbInformation, "提示！"
'    Exit Sub
'
'End If
'
''查询箱号尺寸，计算体积重
'Dim strXHCC As String       '箱数和尺寸
'Dim DblTJZ  As String       '体积重
'Dim order   As String
'
'order = Replace(Ordertemp, "A", "")
'order = Replace$(order, "B", "")
'strXHCC = ""
'DblTJZ = 0
'
'strSql = "SELECT  COUNT(DISTINCT d.箱号) 箱数,d.尺寸  " & " FROM erpdata..tblStockSQfh  a  " & "  INNER JOIN erpdata..tblStockSQfhsub b ON a.单据编号 = b.单据编号 AND a.序号=b.单据项次 " & "   INNER JOIN erpdata..tblStockNumTree c On c.箱号=b.箱号 AND c.基层标记 = 0 " & "   INNER JOIN erpdata..tblStockNumTree d On d.序号 = c.上级序号 AND d.基层标记 = 1 " & " WHERE a.单据编号 IN ('" & order & "','') GROUP BY d.尺寸 "
'If RsNew.State = adStateOpen Then RsNew.Close
'RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
'If Not RsNew.EOF Then
'
'    Do While Not RsNew.EOF
'        '循环得到箱号的箱数和尺寸，进行拼接
'        strXHCC = strXHCC & Trim$("" & RsNew!箱数) & "@" & Trim$("" & RsNew!尺寸) & "cm;"
'        '对尺寸进行分割，计算体积重
'        If Trim$("" & RsNew!尺寸) <> "" And InStr(Trim$("" & RsNew!尺寸), "*") > 0 Then
'            strTmp = Split(Trim$("" & RsNew!尺寸), "*") '分割字符
'            '计算体积重并汇总
'            DblTJZ = DblTJZ + Val(Trim$("" & RsNew!箱数)) * strTmp(0) * strTmp(1) * strTmp(2) / 5000
'
'        End If
'
'        RsNew.MoveNext
'    Loop
'
'End If
'
'RsNew.Close
'wkst.Cells(8, 2) = Mid(dnnum, 1, Len(dnnum) - 1)
'wkst.Cells(8, 5) = ""
'' wkst.Cells(7, 10).Select
'' 生成二维码
'Dim strQrCodePath As String
'
'strNewFullPathNew = strNewFullPathNew & "\" & strExName & strExtsion
'strNewFullPathNew = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\excel" & "\" & strExName & strExtsion
'strQrCodePath = DirQrShare & "\" & strExName & ".JPG"
'strQrCodePath = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\jpg" & "\" & strExName & ".JPG"
'test.Visible = False
'test.QRmaker1.InputData = wkst.Cells(8, 2)
'test.QRmaker1.Refresh
'test.QRmaker1.CreateQrMetaFile hDC, strQrCodePath, 2
'Unload test
'
'wkst.Shapes.AddPicture strQrCodePath, True, True, 1100, 200, 400, 400
'
'wkst.Cells(lngRows + 3, 4) = Format(DblTJZ, "0.00")
''比较体积重和毛重看哪个大,就取哪个
'If DblMZ > DblTJZ Then
'    wkst.Cells(lngRows + 3, 11) = Format(DblMZ, "0.00")
'Else
'    wkst.Cells(lngRows + 3, 11) = Format(DblTJZ, "0.00")
'
'End If
'
''赋值到EXCEL箱数和尺寸
'wkst.Cells(lngRows + 4, 3) = strXHCC
'
'If Len(Dir(strNewFullPath)) > 0 Then
'    If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
'        Exit Sub
'    Else
'
'        On Error Resume Next
'
'        Kill strNewFullPath
'        If Err.number <> 0 Then
'            MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
'            Exit Sub
'
'        End If
'
'    End If
'
'End If
'
'wkbk.SaveAs strNewFullPathNew, xlNormal, "", "", False, False
'wkbk.Saved = True
'
'xlApp.Visible = True
'
'If Not xlApp Is Nothing Then
'    Set wkst = Nothing
'    Set wkbk = Nothing
'    Set xlApp = Nothing
'
'End If
'
'Exit Sub
'ErrHandle:
'
'On Error Resume Next
'
'If Not xlApp Is Nothing Then
'    Set wkst = Nothing
'    Set wkbk = Nothing
'    Set xlApp = Nothing
'
'End If
'
'MsgBox Err.DESCRIPTION, vbInformation, "提示！"
'Exit Sub
End Sub
