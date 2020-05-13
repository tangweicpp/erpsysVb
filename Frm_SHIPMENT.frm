VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_SHIPMENT 
   Caption         =   "出货\入库资料[通用]"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18405
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
   ScaleHeight     =   10935
   ScaleWidth      =   18405
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9600
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_SHIPMENT.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_SHIPMENT.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_SHIPMENT.frx":18A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   18405
      _ExtentX        =   32464
      _ExtentY        =   1535
      ButtonWidth     =   1561
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   查询    "
            Key             =   "SEARCH"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   导出  "
            Key             =   "EXPORT"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   退出   "
            Key             =   "EXIT"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   15735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   28575
      Begin VB.Frame Frame2 
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
         Height          =   8055
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   18135
         Begin FPSpreadADO.fpSpread fpS1 
            Height          =   7455
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   17775
            _Version        =   524288
            _ExtentX        =   31353
            _ExtentY        =   13150
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
            SpreadDesigner  =   "Frm_SHIPMENT.frx":1BF6
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
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
         Height          =   3735
         Left            =   120
         TabIndex        =   18
         Top             =   7200
         Width           =   18015
         Begin FPSpreadADO.fpSpread fpS2 
            Height          =   3375
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   17775
            _Version        =   524288
            _ExtentX        =   31353
            _ExtentY        =   5953
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
            SpreadDesigner  =   "Frm_SHIPMENT.frx":1FE0
         End
      End
      Begin VB.ComboBox cbExcel 
         Height          =   315
         ItemData        =   "Frm_SHIPMENT.frx":23CA
         Left            =   5040
         List            =   "Frm_SHIPMENT.frx":23D7
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox cbTpe 
         Height          =   315
         ItemData        =   "Frm_SHIPMENT.frx":23EE
         Left            =   1200
         List            =   "Frm_SHIPMENT.frx":2416
         TabIndex        =   5
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtShipNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1200
         Width           =   5535
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   2400
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "YYYY-MM-DD"
         Format          =   108658689
         CurrentDate     =   41387
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   1
         Left            =   5040
         TabIndex        =   13
         Top             =   2400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "YYYY-MM-DD"
         Format          =   108658689
         CurrentDate     =   41387
      End
      Begin VB.Label lblLabel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
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
         Left            =   3840
         TabIndex        =   15
         Top             =   2480
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
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
         TabIndex        =   14
         Top             =   2480
         Width           =   900
      End
      Begin MSForms.ComboBox Cbshipto 
         Height          =   315
         Left            =   5040
         TabIndex        =   11
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2990;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblLabel1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GC出货地址"
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
         Left            =   3840
         TabIndex        =   10
         Top             =   1965
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "格式"
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
         Index           =   2
         Left            =   3840
         TabIndex        =   9
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "资料模板"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1590
         Width           =   1200
      End
      Begin MSForms.TextBox txtCusCode 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   1905
         Width           =   2415
         VariousPropertyBits=   746604567
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "4260;661"
         SpecialEffect   =   0
         FontName        =   "宋体"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   1965
         Width           =   1200
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   960
      End
   End
End
Attribute VB_Name = "Frm_SHIPMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gcrev_normal As String
Dim gcrev_wlt As String
Dim wflag As String
Private Enum E_GC_WLA
    e_NO = 1
    E_SubName
    E_ShipTo
    E_FABDevice
    E_CustomerDevice
    E_GCVersion
    E_CSTID
    E_WaferQty
    E_BondPro
    E_PONO
    E_InvoiceNO
    E_ShipOutDate
    E_FABLotID
    E_WAFERID
    E_GROSSDIES
    E_SamplingQty
    E_PassDies
    E_NGDie
    E_Yield
    E_PackLotID
    E_WaferMark
    E_Grade
    E_CartonNO
    E_WO
    E_REMARK
    E_END
End Enum

Private Enum E_GC_WLT
    e_NO = 1
    E_SubName
    E_ShipTo
    E_FABDevice
    E_CustomerDevice
    E_GCVersion
    E_PONO
    E_InvoiceNO
    E_ShipOutDate
    E_FABLotID
    E_WAFERID
    E_GROSSDIES
    E_SamplingQty
    E_PassDies1
    E_PassDies2
    E_PassDies3
    E_NGDie
    E_Yield
    E_PackLotID
    E_WaferMark
    E_Grade
    E_CartonNO
    E_WO
    E_REMARK
    E_END
End Enum

Private Enum E_GC_Normal
    e_NO = 1
    E_SubName
    E_ShipTo
    E_FABDevice
    E_CustomerDevice
    E_GCVersion
    E_PONO
    E_InvoiceNO
    E_ShipOutDate
    E_FABLotID
    E_WAFERID
    E_GROSSDIES
    E_SamplingQty
    E_PassDies
    E_NGDie
    E_Yield
    E_PackLotID
    E_WaferMark
    E_Grade
    E_CartonNO
    E_WO
    E_REMARK
    E_END
End Enum
Private Enum E_GC_Shipping
    e_NO = 1
    E_SubName
    E_ShipTo
    E_FABDevice
    E_CustomerDevice
    E_GCVersion
    e_PO_NO
    E_WO
    E_InvoiceNO
    E_FAB_OutDate
    E_FABLotID
    E_WAFERID
    E_GROSSDIES
    E_SamplingQty
    E_PassDies
    E_Yield
    E_REMARK
    E_TotalQty
    E_unitprice
    E_AmountPrice
    E_NetWeight
    E_TotalWeight
    E_BOX
    E_Extra1
    E_Extra2
    E_Extra3
    E_END
End Enum

Private Sub cbTpe_Click()
   Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    
    'If cbTpe.Text = "WLA出货" Or cbTpe.Text = "GC Shippinglist" Then
    If cbTpe.text = "WLA出货" Then
        Cbshipto.Visible = True
        lblLabel1.Visible = True
   
        If rs.State = adStateOpen Then rs.Close
        rs.Source = "SELECT DISTINCT SHIP_TO  FROM erptemp..customer_information a WHERE a.CUSTOMER = 'GC'"
        
        rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            For i = 1 To rs.RecordCount
                Cbshipto.AddItem (Trim(rs("SHIP_TO")))
                rs.MoveNext
            Next

        End If
    Else
        Cbshipto.text = ""
        Cbshipto.Visible = False
        lblLabel1.Visible = False
     
        
    End If

End Sub






Private Sub Form_Load()
    cbTpe.ListIndex = 0
    
    With fpS1
        .Col = -1
        .Row = -1
        .Lock = True
        
        .ColWidth(1) = 6
        
    End With
    DTP(0).Value = Format(Now(), "YYYY/MM/DD")
    DTP(1).Value = Format(Now(), "YYYY/MM/DD")

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 
    Select Case Button.Key

        Case "SEARCH"
            OnQuery

        Case "EXPORT"
            OnExport

        Case "EXIT"
            Unload Me

    End Select

End Sub

Private Sub OnQuery()

    Dim strShipNO As String

    Toolbar1.Buttons("SEARCH").Enabled = False
    Screen.MousePointer = 11
    
    If cbTpe.text = "GC Shippinglist" Or cbTpe.text = "GC WLA Shippinglist" Or cbTpe.text = "GC list汇总" Then
        txtCusCode.text = "GC"
        ListData
        GoTo EXITTHIS
    End If
    
    If cbTpe.text = "68 Shippinglist" Then
        txtCusCode.text = "68"
        ListData
        GoTo EXITTHIS
    End If
    If cbTpe.text = "HK006 Shippinglist" Then
        txtCusCode.text = "HK006"
        ListData
        GoTo EXITTHIS
    End If
    
    strShipNO = UCase(Trim(txtShipNo.text))

    If strShipNO = "" Then
        MsgBox "请输入单据编号", vbInformation, "提示"
        GoTo EXITTHIS

    End If

    If cbTpe.text = "" Then
        MsgBox "请选择资料模板", vbInformation, "提示"
        GoTo EXITTHIS

    End If


    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    
    If cbTpe.text = "WLA出货" Then
        If Left(strShipNO, 3) = "FWW" Then
            rs.Source = "SELECT  DISTINCT b.CUSTOMERSHORTNAME as 客户代码 FROM erptemp.dbo.tblStockdbsub_temp a INNER JOIN [ERPBASE].[dbo].[tblmappingData] b ON a.lot = b.LOTID AND a.ORDER_NUM = '" & strShipNO & "' "
        Else
            rs.Source = "SELECT  DISTINCT b.CUSTOMERSHORTNAME as 客户代码 FROM erpdata.dbo.tblStockdbsub a INNER JOIN [ERPBASE].[dbo].[tblmappingData] b ON a.工单号 = b.LOTID AND a.调拨编号 = '" & strShipNO & "' "
        End If
    ElseIf cbTpe.text = "BUMPING入库" Then
        rs.Source = "SELECT  DISTINCT b.CUSTOMERSHORTNAME as 客户代码 FROM erpdata .. tblPackToHouse a INNER JOIN [ERPBASE].[dbo].[tblmappingData] b ON a.工单号 = b.LOTID AND a.入库单编号 = '" & strShipNO & "' "
    Else

        rs.Source = "select distinct a.客户代码, b.客户名称 from erpdata..tblStockSQfh a inner join erpdata..tblxcustomer b on a.客户代码 = b.客户代码 where a.单据编号 = '" & strShipNO & "'"
    End If
    
    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        txtCusCode.text = "" & Trim(rs("客户代码"))

    Else
        MsgBox "查询不到该单号,请确认该单号是否有误", vbCritical, "提示"
        GoTo EXITTHIS

    End If
    
    If Trim(txtCusCode.text) = "GC" And Trim(cbTpe.text) = "WLA出货" Then
        If Cbshipto.text = "" Then
            MsgBox "GC客户WLA产品请选择出货地", vbCritical, "提示"
            GoTo EXITTHIS
        End If
    End If
    
    ListData

EXITTHIS:

    Toolbar1.Buttons("SEARCH").Enabled = True
    Screen.MousePointer = 0

End Sub

Private Sub ListData()
    Dim GcVer_Type As String
    
    FPS2hide
    With fpS1
        .MaxRows = 0
    End With
    With fpS2
        .MaxRows = 0
    End With
    Select Case cbTpe.ListIndex

        Case 0
            ListNormal

        Case 1
            ListWLT

        Case 2
            ListWLA
        Case 3
        Case 4
        Case 5
            ListBUMP
        Case 6
            Call ListGCShippinglist("", "")
        Case 7
            ListGCTotallist
        Case 8
           FPS2show
           
            GcVer_Type = GetGcVer_Type(UCase(Trim$(txtShipNo.text)))
        
            If GcVer_Type = "" Then
                Exit Sub
            End If
          
            gcrev_normal = Split(GcVer_Type, ",")(0)
            gcrev_wlt = Split(GcVer_Type, ",")(1)
            If gcrev_normal <> "" Then
                ListNormal
            End If
            If gcrev_wlt <> "" Then
                ListWLT
            End If
       Case 9
            Call ListGCShippinglist_WLA("")
        Case 10, 11
            ListNormal_68
            
    End Select

End Sub



Private Sub FPS2show()
Frame2.Height = Frame3.Top - Frame2.Top - 20
fpS1.Height = Frame2.Height - 30
Frame2.Caption = "Normal"
Frame3.Caption = "WLT"
End Sub

Private Sub FPS2hide()
Frame2.Height = 8055
fpS1.Height = Frame2.Height * 0.9
Frame2.Caption = ""
Frame3.Caption = ""
fpS2.MaxRows = 0
End Sub

Private Sub ListNormal_68()
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    Dim date1Temp   As String
    Dim date2Temp   As String
    
    date1Temp = Format(DTP(0).Value, "YYYY-MM-DD")
    date2Temp = Format(DTP(1).Value + 1, "YYYY-MM-DD")
    strSql = " SELECT DISTINCT d.客户代码, f.QTECHPTNO AS 厂内机种 ,e.MPN_DESC AS 客户机种,x.工单号 AS 'LOT ID',count(DISTINCT x.流程卡编号) AS wafer数量, " & _
   " WaferId = (STUFF((SELECT ',' + SUBSTRING(REPLACE(b1.流程卡编号, '+', ''), len(REPLACE(b1.流程卡编号, '+', '')) - 1, 2)  FROM erpdata .. tblStocksqfhsub b1  " & _
   " WHERE x.工单号 = b1.工单号 and  x.单据编号  = b1.单据编号  order by b1.流程卡编号 FOR XML PATH('')), 1,  1, ''))  " & _
   " , case LEFT(x.大工单,1) when 'A' then '保税' else '非保' end AS 保税非保,d.发货地址 AS 出货地址 ,e.po_num AS PO " & _
   " , x.数量 AS DIE,sum(x.数量) AS 总die数, x.单据编号 AS 单据号,'' AS 发票号,'' AS PRICE  ,'' AS 总价 " & _
   " FROM     erpdata..tblStockSQfh d   " & _
   " inner JOIN  erpdata..tblStocksqfhsub x ON x.单据编号 = d.单据编号  AND x.单据项次 = d.序号  " & _
   " LEFT JOIN ERPBASE..tblmappingData dd ON dd.SUBSTRATEID = x.流程卡编号  " & _
   " LEFT JOIN erpbase..tblCustomerOI e ON CONVERT(VARCHAR(20), convert(int,e.id))  = dd.FILENAME AND e.SOURCE_BATCH_ID = dd.LOTID " & _
   " LEFT JOIN erptemp..TBLTSVNPIPRODUCT f ON x.料号 = f.QTECHPTNO2  and f.customershortname=e.customershortname and f.customerptno1=e.MPN_DESC" & _
   " Where  d.单据编号 IN   (SELECT DISTINCT mm.单据编号 FROM erpdata..tblStockSQfh mm WHERE MM.客户代码 ='" & txtCusCode.text & "' AND mm.单据编号 LIKE 'F%' " & _
   " AND CONVERT(VARCHAR(20), mm.单据日期,23)  >= '" & date1Temp & "' AND CONVERT(VARCHAR(20), mm.单据日期,23) < '" & date2Temp & "')" & _
   " GROUP BY d.客户代码,f.QTECHPTNO,e.MPN_DESC,x.工单号,LEFT(x.大工单,1),d.发货地址 ,e.po_num, x.数量, x.单据编号 " & _
   " ORDER  BY d.客户代码,x.单据编号"
   

    Set rs = Get_SqlserveRs(strSql)

    With fpS1
        .MaxRows = 0
        If rs.RecordCount > 0 Then
                Set .DataSource = rs
        End If
    End With
End Sub

Private Sub ListGCTotallist()
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    Dim date1Temp   As String
    Dim date2Temp   As String
    
    date1Temp = Format(DTP(0).Value, "YYYY-MM-DD")
    date2Temp = Format(DTP(1).Value + 1, "YYYY-MM-DD")
    
    strSql = " SELECT row_number() over(order by a.create_date,d.LOTID , CASE LEN(d.WAFER_ID) WHEN 1 THEN '0'+d.WAFER_ID ELSE d.WAFER_ID END  ) AS 'NO' ,'HTKS' AS 'Sub Name',  " & _
 " case b.发货地址 when '嘉善' then ( case left(e.PO_NUM,2)  when 'HK' THEN 'GCZJ' WHEN 'SH' THEN 'GCZJ' Else 'GCSH'  end )else 'GCSH' end AS 'Ship To',  " & _
 " e.FAB_CONV_ID AS 'Fab Device1',e. MPN_DESC AS 'Customer Device', e.IMAGER_CUSTOMER_REV AS 'GC Version',  " & _
 " e.PO_NUM AS 'PO NO','' AS 'Invoice NO', a.create_date AS 'Ship Out Date',d.LOTID AS 'FAB Lot ID',  " & _
 " CASE LEN(d.WAFER_ID) WHEN 1 THEN '0'+d.WAFER_ID ELSE d.WAFER_ID END  AS 'Wafer ID' , d.PASSBINCOUNT + FAILBINCOUNT AS 'Gross Dies'  " & _
 " ,d.LOTID + CASE LEN(d.WAFER_ID) WHEN 1 THEN '0'+d.WAFER_ID ELSE d.WAFER_ID END  AS 'FAB Device','' AS 'PO',e.CREATED_DATE AS 'WO上传时间'   " & _
 " FROM erptemp..tblshipreport_new a  " & _
 " LEFT JOIN erpdata .. tblStocksqfhsub c ON  a.ship_order =c.单据编号 AND a.qbox =c.箱号  " & _
 " LEFT JOIN erpdata .. tblStocksqfh b ON  b.单据编号 =c.单据编号 and b.序号=c.单据项次    " & _
 " LEFT JOIN erpbase..tblmappingdata d on c.流程卡编号=d.SUBSTRATEID  " & _
 " LEFT JOIN erpbase..tblcustomeroi e on convert(VARCHAR(20),convert(int,e.id))=d.filename AND e.SOURCE_BATCH_ID=d.LOTID  " & _
 " WHERE a.create_date>='" & date1Temp & "' AND a.create_date<'" & date2Temp & "' AND b.客户代码='GC'  "


    Set rs = Get_SqlserveRs(strSql)

    With fpS1
        .MaxRows = 0
        If rs.RecordCount > 0 Then
                Set .DataSource = rs
        End If
    End With
End Sub

Private Sub ListGCShippinglist(strbond As String, strShipTo As String)
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    Dim date1Temp   As String
    Dim date2Temp   As String
    
    date1Temp = Format(DTP(0).Value, "YYYY-MM-DD")
    date2Temp = Format(DTP(1).Value + 1, "YYYY-MM-DD")
    strSql = "SELECT row_number() over(order by z.单据编号) AS 'NO',  'HTKS' AS 'Sub Name','Galaxycore' AS 'Ship to','' as 'FAB Device' ,Z.MPN_DESC AS 'Customer Device','' as 'GC Version',  " & _
    " Z.PO_NUM AS PO_NO ,'' as 'WO','' as 'Invoice NO' ,Z.单据日期 AS 'FAB-Out Date', '' as 'FAB Lot ID',  '' as 'Wafer ID', '' as 'Gross Dies', '' as 'Sampling Qty', '' as 'Pass Dies', '' as 'Yield',   " & _
    " '' as 'Remark',sum(QTY) AS 'Total Qty','' as 'Unit Price', '' as 'Amount Price',SUM(Z.净重) AS 'Net Weight',SUM(Z.重量) AS 'Total Weight' ,sum(箱数)AS Box , z.单据编号 , Z.发货地址, Z.二级代码 FROM (  " & _
    " SELECT x.单据编号 AS 单据编号,convert(varchar(10),d.单据日期,120) AS 单据日期,  e.PO_NUM AS PO_NUM , e.MPN_DESC AS MPN_DESC,f.gcversion AS 二级代码 ,d.发货地址 AS 发货地址  " & _
    " ,round(isnull(cc.重量,0),4) as 重量,CC.箱号, round(SUM(x.数量) * 0.1/6000,4) as 净重,count(ISNULL(cc.箱号,bb.箱号)) AS QTY, count(DISTINCT ISNULL(cc.箱号,bb.箱号)) AS 箱数  " & _
    " FROM     erpdata..tblStockSQfh d   " & _
    " inner JOIN  erpdata..tblStocksqfhsub x ON x.单据编号 = d.单据编号  AND x.单据项次 = d.序号  " & _
    " INNER JOIN erpdata..tblStockNumTree bb ON bb.箱号 = x.箱号 AND bb.基层标记 = 0   " & _
    " LEFT JOIN erpdata..tblStockNumTree cc ON cc.序号 = bb.上级序号 AND cc.基层标记 = 1   " & _
    " LEFT JOIN ERPBASE..tblmappingData dd ON dd.SUBSTRATEID = x.流程卡编号   " & _
    " LEFT JOIN erpbase..tblCustomerOI e ON CONVERT(VARCHAR(20), convert(int,e.id))  = dd.FILENAME AND e.SOURCE_BATCH_ID = dd.LOTID    " & _
    " LEFT JOIN erptemp..tblshipreport_new f ON x.单据编号 = f.ship_order AND x.箱号 =f.qbox AND x.工单号 = f.lot_id    " & _
    " Where  d.单据编号 IN   (SELECT DISTINCT mm.单据编号 FROM erpdata..tblStockSQfh mm WHERE MM.客户代码 ='GC' AND mm.单据编号 LIKE 'F%'  " & _
    " AND CONVERT(VARCHAR(20), mm.单据日期,23) >= '" & date1Temp & "' AND CONVERT(VARCHAR(20), mm.单据日期,23) < '" & date2Temp & "')"
    If strbond = "保税" Then
        strSql = strSql & " and left(e.PO_NUM ,2)='HK'"
    ElseIf strbond = "非保税" Then
        strSql = strSql & " and left(e.PO_NUM ,2)='SH'"
    End If
    If strShipTo <> "" Then
        strSql = strSql & " and  d.发货地址='" & strShipTo & "'"
    End If
    strSql = strSql & " GROUP BY x.单据编号, convert(varchar(10),d.单据日期,120)  ,e.PO_NUM, e.CUSTOMERSHORTNAME , e.MPN_DESC, d.发货地址 ,f.gcversion,round(isnull(cc.重量,0),4),CC.箱号) Z  " & _
    " GROUP BY z.单据编号,Z.单据日期 ,Z.PO_NUM,Z.MPN_DESC,Z.二级代码,Z.发货地址  " & _
    " ORDER BY z.单据编号,Z.MPN_DESC,Z.PO_NUM "
   

    Set rs = Get_SqlserveRs(strSql)

    With fpS1
        .MaxRows = 0
        If rs.RecordCount > 0 Then
                Set .DataSource = rs
        End If
    End With
End Sub


Private Sub ListGCShippinglist_WLA(strbond As String)
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    Dim date1Temp   As String
    Dim date2Temp   As String
    
    date1Temp = Format(DTP(0).Value, "YYYY-MM-DD")
    date2Temp = Format(DTP(1).Value + 1, "YYYY-MM-DD")
    

 
strSql = "SELECT replace(d.MPN_DESC,'-3','-2.5')  AS 存货编码 ,'' AS 存货名称 ,'' AS 主计量单位,'' AS 库存单位,'' AS 件数, count(ISNULL(cc.箱号,bb.箱号)) AS 数量,d.IMAGER_CUSTOMER_REV   AS 二级代码, " & _
" 'SWHT' AS 贸易类型,'' AS 等级, d.PO_NUM AS 'PO NO','' AS Box " & _
" FROM  erpdata..tblstockdbsub a " & _
" LEFT JOIN erpdata..tblstockdb b ON a.调拨编号=b.调拨编号 AND a.序号=b.序号 " & _
"INNER JOIN erpdata..tblStockNumTree bb ON bb.箱号 = a.箱号 AND bb.基层标记 = 0 " & _
" LEFT JOIN erpdata..tblStockNumTree cc ON cc.序号 = bb.上级序号 AND cc.基层标记 = 1 " & _
" LEFT JOIN erpbase..tblmappingdata c ON rtrim(a.流程卡编号)=rtrim(c.SUBSTRATEID) AND rtrim(a.工单号)=rtrim(c.LOTID) " & _
" LEFT JOIN erpbase..tblCustomerOI  d ON rtrim(c.LOTID)=d.SOURCE_BATCH_ID  AND c.FILENAME =convert(VARCHAR(20),convert(int,d.id)) " & _
" WHERE  d.CUSTOMERSHORTNAME ='GC' AND b.目标仓库='72'" & _
" AND CONVERT(VARCHAR(20), b.申请时间,23) >= '" & date1Temp & "' AND CONVERT(VARCHAR(20),  b.申请时间,23) < '" & date2Temp & "'" & _
" And a.调拨编号 not in (select 关联调拨编号 from ERPTEMP..InvalidStockDb) "

    If strbond = "保税" Then
        strSql = strSql & " and left(d.PO_NUM ,2)='HK'"
    ElseIf strbond = "非保税" Then
        strSql = strSql & " and left(d.PO_NUM ,2)='SH'"
    End If

    strSql = strSql & " GROUP BY  d.MPN_DESC,d.IMAGER_CUSTOMER_REV,d.PO_NUM  "

    Set rs = Get_SqlserveRs(strSql)

    With fpS1
        .MaxRows = 0
        If rs.RecordCount > 0 Then
           Set .DataSource = rs
        End If
    End With
    If fpS1.MaxRows > 0 Then
        strSql = "SELECT count(DISTINCT ISNULL(cc.箱号,bb.箱号)) AS Box " & _
        " FROM  erpdata..tblstockdbsub a " & _
        " LEFT JOIN erpdata..tblstockdb b ON a.调拨编号=b.调拨编号 AND a.序号=b.序号 " & _
        "INNER JOIN erpdata..tblStockNumTree bb ON bb.箱号 = a.箱号 AND bb.基层标记 = 0 " & _
        " LEFT JOIN erpdata..tblStockNumTree cc ON cc.序号 = bb.上级序号 AND cc.基层标记 = 1 " & _
        " LEFT JOIN erpbase..tblmappingdata c ON rtrim(a.流程卡编号)=rtrim(c.SUBSTRATEID) AND rtrim(a.工单号)=rtrim(c.LOTID) " & _
        " LEFT JOIN erpbase..tblCustomerOI  d ON rtrim(c.LOTID)=d.SOURCE_BATCH_ID  AND c.FILENAME =convert(VARCHAR(20),convert(int,d.id)) " & _
        " WHERE  d.CUSTOMERSHORTNAME ='GC' AND b.目标仓库='72'" & _
        " AND CONVERT(VARCHAR(20), b.申请时间,23) >= '" & date1Temp & "' AND CONVERT(VARCHAR(20),  b.申请时间,23) < '" & date2Temp & "'" & _
        " And a.调拨编号 not in (select 关联调拨编号 from ERPTEMP..InvalidStockDb) "

        If strbond = "保税" Then
            strSql = strSql & " and left(d.PO_NUM ,2)='HK'"
        ElseIf strbond = "非保税" Then
            strSql = strSql & " and left(d.PO_NUM ,2)='SH'"
        End If

        fpS1.MaxRows = fpS1.MaxRows + 1
        fpS1.SetText 10, fpS1.MaxRows, "汇总"
        fpS1.SetText 11, fpS1.MaxRows, GetSqlServerStr(strSql)
    End If
    
End Sub

Private Sub ListNormal()

Select Case txtCusCode.text

    Case "GC"
        ListNormal_GC
    
    Case "SX", "TJ003", "SC081"
        ListNormal_SX
    
    Case "AC64", "SD", "SH115", "HD"
        ListNormal_AC64
        
    Case "DA69"
        ListNormal_DA69
    
    Case "HK037", "AC70"
        ListNormal_AC70
    Case "SH103"
        ListNormal_SH103
    Case "SH105"
        ListNormal_SH105
    Case "BB32"
        ListNormal_BB32
    Case "SH48"
        ListNormal_SH48
    Case "AH017"
        ListNormal_AH017
    Case "XD46", "SC057", "SC03", "B1", "TS26"
        ListNormal_XD46
        
    Case "SH07", "XD36", "BJ139", "JS161", "SC060", "FJ030", "XD88", "XD66", "HK010", "SH188", "TJ008", "RD"
        ListNormal_SH07
        
    Case "57"
        ListNormal_57
    Case "HK005"
        ListNormal_HK005
    Case "AC03"
        ListNormal_AC03
End Select
End Sub


Private Sub ListNormal_AC03()
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
strShipNO = UCase(Trim$(txtShipNo.text))

strSql = " SELECT t8.*,CASE WHEN t9.lot数>1 THEN '是'ELSE '否'END AS     是否合箱 FROM ( " & _
" SELECT t4.mpn_desc AS 产品型号,isnull(t7.箱号,t6.箱号) AS 外箱号,isnull(t4.ASSEMBLY_FACILITY,'') AS 封装形式,t4.RETICLE_LEVEL_71 AS 打印周期, " & _
" CASE WHEN t5.库房名称 LIKE '%不良%' THEN '不良品' ELSE '良品' END  AS '状态 (良品or不良品)',CASE WHEN t5.库房名称 LIKE '%不良%' THEN '散装' ELSE '编带' END  AS 包装方式, " & _
" t4.SOURCE_BATCH_ID AS 芯片批号,t4.po_num AS 订单号,sum(t2.数量) AS    出货数量 " & _
" FROM  erpdata .. tblStocksqfh t1 " & _
" INNER JOIN erpdata .. tblStocksqfhsub t2 ON t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次    " & _
" INNER JOIN erpbase .. tblmappingData t3 ON  t2.流程卡编号 = t3.SUBSTRATEID  " & _
" INNER JOIN  erpbase .. tblCustomerOI t4 ON  t3.FILENAME = convert(varchar(20),t4.id)  AND t3.LOTID = t4.SOURCE_BATCH_ID " & _
" INNER JOIN erpdata..tblstock t5 ON t1.仓库编号=t5.库房代码  " & _
" INNER JOIN erpdata .. tblPackTreeInf t6 ON   t6.箱号=t2.箱号  " & _
" INNER JOIN erpdata .. tblPackTreeInf  t7 ON   t7.序号=t6.上级序号 AND t7.基层标记=1  " & _
" where  t1.单据编号 = '" & strShipNO & "' GROUP BY  t4.mpn_desc ,t4.ASSEMBLY_FACILITY,t4.RETICLE_LEVEL_71,t5.库房名称,t4.SOURCE_BATCH_ID,t4.po_num,isnull(t7.箱号,t6.箱号)) AS t8 " & _
" INNER JOIN (SELECT  count(DISTINCT a.工单号) AS lot数,isnull(c.箱号,b.箱号) AS 大箱号 FROM  erpdata..tblPackMainInfsub a, erpdata .. tblPackTreeInf b , erpdata .. tblPackTreeInf  c " & _
" WHERE a.箱号=b.箱号 AND b.上级序号=c.序号 GROUP BY isnull(c.箱号,b.箱号) ) AS t9 ON t8.外箱号=t9.大箱号 "
Set rs = Get_SqlserveRs(strSql)
With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With
End Sub




Private Sub ListNormal_HK005()
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
strShipNO = UCase(Trim$(txtShipNo.text))

' strSql = "select CONVERT(VARCHAR(100), t1.单据日期, 23) as '委外厂入库日期', t3.PO_NUM as '采购单号',t3.MPN_DESC as '料号', reverse(substring(reverse(t3.MPN_DESC),charindex('.',reverse(t3.MPN_DESC)) +1,500)) as '品名', " & _
        ' "t3.SOURCE_BATCH_ID as '晶圆批次号',t4.wafer_id as '片号',  " & _
        ' "SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as 'BP后的良品数' from erpdata .. tblStocksqfh t1, " & _
        ' "erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, " & _
        ' "erpdata .. tblPackTreeInf t5, erpdata .. tblErpInStockDetailInfo t7, " & _
        ' "(SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) " & _
        ' " AS wafer FROM  erpdata .. tblErpInStockRelation t8 where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
        ' "erpdata .. tblErpInStockDetailInfo t11 left join erpdata .. tblPackTreeInf on 序号 = erpdata .. tblPackTreeInf.上级序号 where t1.单据编号 = '" + strShipNO + "'" & _
        ' "and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
        ' "and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
        ' "AND t5.箱号 = t2.箱号  AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  " & _
        ' "AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  " & _
        ' "AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
        ' "AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' " & _
        ' "AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID " & _
        ' "GROUP BY CONVERT(VARCHAR(100), t1.单据日期, 23), t3.PO_NUM,t3.MPN_DESC,t3.SOURCE_BATCH_ID,t4.wafer_id order by t3.SOURCE_BATCH_ID "
strSql = "select CONVERT(VARCHAR(100), t1.单据日期, 23) as '委外厂入库日期', t3.PO_NUM as '采购单号',t3.MPN_DESC as '料号', reverse(substring(reverse(t3.MPN_DESC),charindex('.',reverse(t3.MPN_DESC)) +1,500)) as '品名', " & _
        "t3.SOURCE_BATCH_ID as '晶圆批次号',t4.wafer_id as '片号',  " & _
        " sum(CONVERT(INT, t6.KEY_VALUE)) AS 'BP后的良品数' from erpdata .. tblStocksqfh t1, " & _
        " erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4 ," & _
        " erpdata .. tblErpInStockDetailInfo  t5 ," & _
        " erpdata .. tblErpInStockDetailInfo  t6  " & _
        " where t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
        " and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID  " & _
        " AND T5.key_value=T2.流程卡编号 AND T5.KEY_TYPE='WAFER' AND t5.KEY_NAME='NAME' " & _
        " AND t6.keyid=t5.keyid AND t6.BOX_ID=t5.BOX_ID AND t6. KEY_NAME='GOOD_DIE'  " & _
        " and t1.单据编号 = '" + strShipNO + "'" & _
        " GROUP BY CONVERT(VARCHAR(100), t1.单据日期, 23), t3.PO_NUM,t3.MPN_DESC,t3.SOURCE_BATCH_ID,t4.wafer_id  order by t3.SOURCE_BATCH_ID,t4.wafer_id "
Set rs = Get_SqlserveRs(strSql)
With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With
End Sub

Private Sub ListNormal_57()
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
strShipNO = UCase(Trim$(txtShipNo.text))

strSql = "select ROW_NUMBER() over(order by t3.SOURCE_BATCH_ID asc) as item,t3.PO_NUM as 'PO NO',t3.TEST_SITE as 'supplier',t3.MPN_DESC as 'Customer Device', " & _
        "t3.SOURCE_BATCH_ID as 'Lot ID',wafer_id = (STUFF((SELECT ',' + convert(varchar(2),k4.wafer_id) from " & _
        "(SELECT t4.lotid,convert(int,t4.wafer_id) as 'wafer_id'  from erpdata .. tblStocksqfh t1, " & _
        "erpdata .. tblStocksqfhsub t2,erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4 where t1.单据编号 = '" & strShipNO & "'" & _
        "and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID  " & _
        "and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID group by convert(int,t4.wafer_id) ,t4.lotid " & _
        ") k4 where k4.lotid = t4.lotid  order by k4.wafer_id for xml path('')),1,1,'')),  " & _
        "SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as 'NG  Die Qty', " & _
        "SUM(CONVERT(INT, t9.KEY_VALUE)) AS 'GOOD Die Qty', count(t4.substrateid) as 'wafer QTY', t1.发货地址 as Destination, ' ' as Forwarder, " & _
        "CONVERT(VARCHAR(100), t1.单据日期, 23) AS 'shipping date'  from erpdata .. tblStocksqfh t1, " & _
        "erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, " & _
        "erpdata .. tblPackTreeInf t5, erpdata .. tblErpInStockDetailInfo t7, " & _
        "(SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) " & _
        " AS wafer FROM  erpdata .. tblErpInStockRelation t8 where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
        "erpdata .. tblErpInStockDetailInfo t11 left join erpdata .. tblPackTreeInf on 序号 = erpdata .. tblPackTreeInf.上级序号  where t1.单据编号 = '" & strShipNO & "'" & _
        "and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
        "and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
        "AND t5.箱号 = t2.箱号  AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  " & _
        "AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  " & _
        "AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
        "AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' " & _
        "AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID " & _
        "GROUP BY t3.PO_NUM,t3.TEST_SITE,t3.MPN_DESC,t3.SOURCE_BATCH_ID,t4.lotid,CONVERT(VARCHAR(100), t1.单据日期, 23),t1.发货地址 order by t3.SOURCE_BATCH_ID "

'strSql = "select ROW_NUMBER() over(order by t3.SOURCE_BATCH_ID asc) as item,t3.PO_NUM as 'PO NO',t3.TEST_SITE as 'supplier',t3.MPN_DESC as 'Customer Device', " & _
'        "t3.SOURCE_BATCH_ID as 'Lot ID',wafer_id = (STUFF((SELECT ',' + k4.wafer_id from " & _
'        "(SELECT t4.wafer_id ,t4.lotid from erpdata .. tblStocksqfh t1, " & _
'        "erpdata .. tblStocksqfhsub t2,erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4 where t1.单据编号 = '" & strShipNO & "'" & _
'        "and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID  " & _
'        "and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
'        ") k4 where k4.lotid = t4.lotid  for xml path('')),1,1,'')),  " & _
'        "SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as 'Bin000  Die Qty', " & _
'        "SUM(CONVERT(INT, t9.KEY_VALUE)) AS 'NG  Die Qty', count(t4.substrateid) as 'wafer QTY', t1.发货地址 as Destination, ' ' as Forwarder, " & _
'        "CONVERT(VARCHAR(100), t1.单据日期, 23) AS 'shipping date'  from erpdata .. tblStocksqfh t1, " & _
'        "erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, " & _
'        "erpdata .. tblPackTreeInf t5, erpdata .. tblErpInStockDetailInfo t7, " & _
'        "(SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) " & _
'        " AS wafer FROM  erpdata .. tblErpInStockRelation t8 where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
'        "erpdata .. tblErpInStockDetailInfo t11 left join erpdata .. tblPackTreeInf on 序号 = erpdata .. tblPackTreeInf.上级序号 " & _
'        "where t1.单据编号 = '" & strShipNO & "'" & _
'        "and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
'        "and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
'        "AND t5.箱号 = t2.箱号  AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  " & _
'        "AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  " & _
'        "AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
'        "AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' " & _
'        "AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID " & _
'        "GROUP BY t3.PO_NUM,t3.TEST_SITE,t3.MPN_DESC,t3.SOURCE_BATCH_ID,t4.lotid,CONVERT(VARCHAR(100), t1.单据日期, 23),t1.发货地址 order by t3.SOURCE_BATCH_ID "
         wflag = "1"
Set rs = Get_SqlserveRs(strSql)
With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
        Set .DataSource = rs
    End If
End With

End Sub
Private Sub ListNormal_SH07()
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
strShipNO = UCase(Trim$(txtShipNo.text))

strSql = "select ROW_NUMBER() over(order by t3.SOURCE_BATCH_ID asc) as item,t3.PO_NUM as 'PO NO',t3.SHIP_SITE as 'supplier',t3.MPN_DESC as 'Customer Device', " & _
        "t3.SOURCE_BATCH_ID as 'Lot ID',wafer_id = (STUFF((SELECT ',' + convert(varchar(2),k4.wafer_id) from " & _
        "(SELECT t4.lotid,convert(int,t4.wafer_id) as 'wafer_id'  from erpdata .. tblStocksqfh t1, " & _
        "erpdata .. tblStocksqfhsub t2,erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4 where t1.单据编号 = '" & strShipNO & "'" & _
        "and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID  " & _
        "and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID group by convert(int,t4.wafer_id) ,t4.lotid " & _
        ") k4 where k4.lotid = t4.lotid  order by k4.wafer_id for xml path('')),1,1,'')),  " & _
        "SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as 'NG  Die Qty', " & _
        "SUM(CONVERT(INT, t9.KEY_VALUE)) AS 'Bin000  Die Qty', count(t4.substrateid) as 'wafer QTY', t1.发货地址 as Destination, ' ' as Forwarder, " & _
        "CONVERT(VARCHAR(100), t1.单据日期, 23) AS 'shipping date'  from erpdata .. tblStocksqfh t1, " & _
        "erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, " & _
        "erpdata .. tblPackTreeInf t5, erpdata .. tblErpInStockDetailInfo t7, " & _
        "(SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) " & _
        " AS wafer FROM  erpdata .. tblErpInStockRelation t8 where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
        "erpdata .. tblErpInStockDetailInfo t11 left join erpdata .. tblPackTreeInf on 序号 = erpdata .. tblPackTreeInf.上级序号  where t1.单据编号 = '" & strShipNO & "'" & _
        "and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
        "and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
        "AND t5.箱号 = t2.箱号  AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  " & _
        "AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  " & _
        "AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
        "AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' " & _
        "AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID " & _
        "GROUP BY t3.PO_NUM,t3.SHIP_SITE,t3.MPN_DESC,t3.SOURCE_BATCH_ID,t4.lotid,CONVERT(VARCHAR(100), t1.单据日期, 23),t1.发货地址 order by t3.SOURCE_BATCH_ID "

'strSql = "select ROW_NUMBER() over(order by t3.SOURCE_BATCH_ID asc) as item,t3.PO_NUM as 'PO NO',t3.TEST_SITE as 'supplier',t3.MPN_DESC as 'Customer Device', " & _
'        "t3.SOURCE_BATCH_ID as 'Lot ID',wafer_id = (STUFF((SELECT ',' + k4.wafer_id from " & _
'        "(SELECT t4.wafer_id ,t4.lotid from erpdata .. tblStocksqfh t1, " & _
'        "erpdata .. tblStocksqfhsub t2,erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4 where t1.单据编号 = '" & strShipNO & "'" & _
'        "and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID  " & _
'        "and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
'        ") k4 where k4.lotid = t4.lotid  for xml path('')),1,1,'')),  " & _
'        "SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as 'Bin000  Die Qty', " & _
'        "SUM(CONVERT(INT, t9.KEY_VALUE)) AS 'NG  Die Qty', count(t4.substrateid) as 'wafer QTY', t1.发货地址 as Destination, ' ' as Forwarder, " & _
'        "CONVERT(VARCHAR(100), t1.单据日期, 23) AS 'shipping date'  from erpdata .. tblStocksqfh t1, " & _
'        "erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, " & _
'        "erpdata .. tblPackTreeInf t5, erpdata .. tblErpInStockDetailInfo t7, " & _
'        "(SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) " & _
'        " AS wafer FROM  erpdata .. tblErpInStockRelation t8 where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
'        "erpdata .. tblErpInStockDetailInfo t11 left join erpdata .. tblPackTreeInf on 序号 = erpdata .. tblPackTreeInf.上级序号 " & _
'        "where t1.单据编号 = '" & strShipNO & "'" & _
'        "and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
'        "and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
'        "AND t5.箱号 = t2.箱号  AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  " & _
'        "AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  " & _
'        "AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
'        "AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' " & _
'        "AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID " & _
'        "GROUP BY t3.PO_NUM,t3.TEST_SITE,t3.MPN_DESC,t3.SOURCE_BATCH_ID,t4.lotid,CONVERT(VARCHAR(100), t1.单据日期, 23),t1.发货地址 order by t3.SOURCE_BATCH_ID "
         wflag = "1"
Set rs = Get_SqlserveRs(strSql)
With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
        Set .DataSource = rs
    End If
End With

End Sub
Private Sub ListNormal_XD46()
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
strShipNO = UCase(Trim$(txtShipNo.text))

strSql = "select t3.TEST_SITE as 'supplier',t4.customershortname as 'Customer Code',t3.MPN_DESC as 'Customer Device', t3.PO_NUM as 'PO Number', t3.SOURCE_BATCH_ID as 'Customer Lot', " & _
"t4.wafer_id as 'WaferNo', sum(t4.PASSBINCOUNT) as 'incoming good Die', " & _
"SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as 'process NGdieQty',SUM(CONVERT(INT, t9.KEY_VALUE)) AS 'Shipment GoodDie', " & _
"CONVERT(VARCHAR(100), t1.单据日期, 23) AS 'Shipmentdate', '出货天水' as remark from erpdata .. tblStocksqfh t1, " & _
"erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, " & _
"erpdata .. tblPackTreeInf t5, erpdata .. tblErpInStockDetailInfo t7, " & _
"(SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) " & _
" AS wafer FROM  erpdata .. tblErpInStockRelation t8 where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
"erpdata .. tblErpInStockDetailInfo t11 left join erpdata .. tblPackTreeInf on 序号 = erpdata .. tblPackTreeInf.上级序号 where t1.单据编号 = '" & strShipNO & "'" & _
"and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
"and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
"AND t5.箱号 = t2.箱号  AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  " & _
"AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  " & _
"AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
"AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' " & _
"AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID GROUP BY t3.TEST_SITE,t4.customershortname,t3.MPN_DESC, t3.PO_NUM, t3.SOURCE_BATCH_ID , " & _
"t3.RETICLE_LEVEL_72,t4.wafer_id,t4.PRODUCTID, CONVERT(VarChar(100), t1.单据日期, 23) "

Set rs = Get_SqlserveRs(strSql)
With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With
End Sub

Private Sub ListNormal_AH017()
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
strShipNO = UCase(Trim$(txtShipNo.text))

strSql = "select t3.TEST_SITE as 'supplier',t4.customershortname,t3.MPN_DESC as 'Customer Device',t3.PO_NUM as 'PO Number', " & _
"t3.TARGET_WAF_THICKNESS as 'Customer Lot',t4.wafer_id as 'WaferNo', " & _
"SUM(CONVERT(INT, t9.KEY_VALUE)) AS 'Shipment GoodDie',SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as 'shipping  NG die'," & _
"t4.productid as 'Laser Mark',CONVERT(VARCHAR(100), t1.单据日期, 23) AS 'Shipment Date', t6.箱号 as 'Box.',' ' as Remark from erpdata .. tblStocksqfh t1, " & _
"erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, " & _
"erpdata .. tblPackTreeInf t5,erpdata .. tblPackTreeInf t6, erpdata .. tblErpInStockDetailInfo t7, " & _
"(SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) " & _
" AS wafer FROM  erpdata .. tblErpInStockRelation t8 where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
"erpdata .. tblErpInStockDetailInfo t11 where t1.单据编号 = '" & strShipNO & "' " & _
"and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
"and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
"AND t5.箱号 = t2.箱号  AND t6.序号 = t5.上级序号 AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  " & _
"AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  " & _
"AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
"AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' " & _
"AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID GROUP BY  " & _
"t3.TEST_SITE ,t4.customershortname,t3.MPN_DESC,t3.PO_NUM, " & _
"t3.TARGET_WAF_THICKNESS ,t4.wafer_id,t6.箱号,t4.productid,CONVERT(VarChar(100), t1.单据日期, 23) "

Set rs = Get_SqlserveRs(strSql)
With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With

End Sub

Private Sub ListNormal_BB32()
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
strShipNO = UCase(Trim$(txtShipNo.text))

'strSql = "select 'HTKS' AS Supplier, 'BB32' AS  'Customer Code',t3.MPN_DESC as 'Customer Device', t3.PO_NUM as 'PO Number', t3.SOURCE_BATCH_ID as 'Customer Lot', " & _
'"t3.RETICLE_LEVEL_72 as 'LOT NO',t4.WAFER_ID as 'WaferNo', " & _
'"SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as ""process NGdieQty"",SUM(CONVERT(INT, t9.KEY_VALUE)) AS ""Shipment GoodDie"", " & _
'"t4.PRODUCTID as 'Laser Mark',CONVERT(VARCHAR(100), t1.单据日期, 23) AS 出货日期, ' ' as remark from erpdata .. tblStocksqfh t1, " & _
'"erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, " & _
'"erpdata .. tblPackTreeInf t5, erpdata .. tblErpInStockDetailInfo t7, " & _
'"(SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) " & _
'" AS wafer FROM  erpdata .. tblErpInStockRelation t8 where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
'"erpdata .. tblErpInStockDetailInfo t11 where t1.单据编号 = '" & strShipNO & "' " & _
'"and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
'"and t4.FILENAME = convert(varchar(50), t3.id)  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
'"AND t5.箱号 = t2.箱号  AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  " & _
'"AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  " & _
'"AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
'"AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' " & _
'"AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID GROUP BY t3.MPN_DESC, t3.PO_NUM, t3.SOURCE_BATCH_ID , " & _
'"t3.RETICLE_LEVEL_72,t4.WAFER_ID,t4.PRODUCTID, CONVERT(VarChar(100), t1.单据日期, 23) "

strSql = "select 'HTKS' AS Supplier, 'BB32' AS  'Customer Code',t3.MPN_DESC as 'Customer Device', t3.PO_NUM as 'PO Number', t3.SOURCE_BATCH_ID as 'Customer Lot', " & _
"t3.RETICLE_LEVEL_72 as 'LOT NO',t4.WAFER_ID as 'WaferNo', " & _
"SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as ""process NGdieQty"",SUM(CONVERT(INT, t9.KEY_VALUE)) AS ""Shipment GoodDie"", " & _
"t4.PRODUCTID as 'Laser Mark',CONVERT(VARCHAR(100), t1.单据日期, 23) AS 出货日期, ' ' as remark from erpdata .. tblStocksqfh t1, " & _
"erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, " & _
"erpdata .. tblPackTreeInf t5, erpdata .. tblErpInStockDetailInfo t7, " & _
"(SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) " & _
" AS wafer FROM  erpdata .. tblErpInStockRelation t8 where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
"erpdata .. tblErpInStockDetailInfo t11 left join erpdata .. tblPackTreeInf on 序号 = erpdata .. tblPackTreeInf.上级序号 where t1.单据编号 = '" & strShipNO & "' " & _
"and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
"and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
"AND t5.箱号 = t2.箱号  AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  " & _
"AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  " & _
"AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
"AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' " & _
"AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID GROUP BY t3.MPN_DESC, t3.PO_NUM, t3.SOURCE_BATCH_ID , " & _
"t3.RETICLE_LEVEL_72,t4.WAFER_ID,t4.PRODUCTID, CONVERT(VarChar(100), t1.单据日期, 23) "

Set rs = Get_SqlserveRs(strSql)
With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With

End Sub

Private Sub ListNormal_SH48()
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim loPOQty As Long
Dim loShipQty As Long
Dim strpo As String

strShipNO = UCase(Trim$(txtShipNo.text))

strSql = "select t3.PO_NUM as '订单编号',t3.MPN_DESC as '产品名称', t3.FAB_CONV_ID as '封装形式',t3.TARGET_WAF_THICKNESS as 'LOT NO.','' as 'D/C', " & _
"SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) + SUM(CONVERT(INT, t9.KEY_VALUE))  as '发货数量',SUM(CONVERT(INT, t9.KEY_VALUE)) AS '良品数量(PCS)', " & _
"SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as '不良品数量(PCS)', " & _
"'' as '备注一', '' as '是否结单' from erpdata .. tblStocksqfh t1, erpdata .. tblStocksqfhsub t2,erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, " & _
"erpdata .. tblPackTreeInf t5,erpdata .. tblPackTreeInf t6, erpdata .. tblErpInStockDetailInfo t7, " & _
"(SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) " & _
"AS wafer FROM  erpdata .. tblErpInStockRelation t8 where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
"erpdata .. tblErpInStockDetailInfo t11 where t1.单据编号 = '" & strShipNO & "' " & _
"and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
"and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
"AND t5.箱号 = t2.箱号  AND t6.序号 = t5.上级序号 AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  " & _
"AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  " & _
"AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
"AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' " & _
"AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID GROUP BY  t3.PO_NUM,t3.MPN_DESC,t3.RETICLE_LEVEL_71,t3.FAB_CONV_ID,t3.TARGET_WAF_THICKNESS "

'strSql = "select t3.TEST_SITE as 'supplier',t4.customershortname,t3.MPN_DESC as 'Customer Device',t3.PO_NUM as 'PO Number'," & _
'"t3.TARGET_WAF_THICKNESS as 'Customer Lot',t4.wafer_id as 'WaferNo',t6.""箱号""," & _
'"SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as ""shipping  NG die"",SUM(CONVERT(INT, t9.KEY_VALUE)) AS ""Shipment GoodDie"", " & _
'"CONVERT(VARCHAR(100), t1.单据日期, 23) AS 出货日期, ' ' as remark from erpdata .. tblStocksqfh t1, " & _
'"erpdata .. tblStocksqfhsub t2,erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, " & _
'"erpdata .. tblPackTreeInf t5,erpdata .. tblPackTreeInf t6, erpdata .. tblErpInStockDetailInfo t7, " & _
'"(SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) " & _
'" AS wafer FROM  erpdata .. tblErpInStockRelation t8 where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
'"erpdata .. tblErpInStockDetailInfo t11 where t1.单据编号 = '" & strShipNO & "' " & _
'"and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
'"and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
'"AND t5.箱号 = t2.箱号  AND t6.序号 = t5.上级序号 AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  " & _
'"AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  " & _
'"AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
'"AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' " & _
'"AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID GROUP BY  " & _
'"t3.TEST_SITE ,t4.customershortname,t3.MPN_DESC,t3.PO_NUM," & _
'"t3.TARGET_WAF_THICKNESS ,t4.wafer_id,t6.""箱号"",CONVERT(VarChar(100), t1.单据日期, 23) "
Toolbar1.Buttons("EXPORT").Enabled = False
Set rs = Get_SqlserveRs(strSql)
With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        strpo = Trim(.text)
        loPOQty = Get_SqlserverNo("select isnull(sum(isnull(a.PASSBINCOUNT,0) + isnull(a.FAILBINCOUNT,0)),0) from erpbase..tblmappingData a inner join ERPBASE..TBLCUSTOMEROI b ON  convert(VARCHAR(30),b.ID)=a.FILENAME AND b.SOURCE_BATCH_ID=a.LOTID and  right(a.SUBSTRATEID,1)<>'+' and b.customershortname='SH48' and  b.PO_NUM='" & strpo & "'")
        loShipQty = Get_SqlserverNo("select isnull(sum(c.数量*d.标记),0) from erpdata..tblstocksqfhsub c inner join  erpdata..tblstocksqfh d on  c.单据编号=d.单据编号  AND c.单据项次=d.序号   inner join erpbase..tblmappingData a  on  c.流程卡编号=a.SUBSTRATEID  inner join ERPBASE..TBLCUSTOMEROI b ON  convert(VARCHAR(30),b.ID)=a.FILENAME AND b.SOURCE_BATCH_ID=a.LOTID  where b.customershortname='SH48' and left(c.单据编号,1) IN ('F','T')  and  b.PO_NUM='" & strpo & "'")
        
        If loPOQty = loShipQty Then
            .SetText 10, i, "是"
        Else
            .SetText 10, i, "否"
        End If

    Next
    
End With
Toolbar1.Buttons("EXPORT").Enabled = True

End Sub

Private Sub ListNormal_SH105()
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
strShipNO = UCase(Trim$(txtShipNo.text))
'
'strSql = " SELECT xx.NO,xx.供货方,xx.客户,xx.产品名称,xx.客户订单号 ,xx.客户Lot,xx.WaferNo,xx.GoodDieQty,xx.Reel_Code,xx.PKG_LOT ,  " & _
'" SUBSTRING(yy.Content,CHARINDEX('""PACKING_DATE_10"",""',yy.Content)+19,6) AS DC ,xx.出货日期,xx.箱号  FROM( SELECT y.*,MAX(x.ID) AS PRINT_ID  " & _
'" FROM ( SELECT ROW_NUMBER() over(order by REPLACE(t2.工单号, ' ', ''), t4.SUBSTRATEID) AS NO, 'HTKS' AS 供货方, 'SH103' AS 客户,  " & _
'" t3.MPN_DESC AS 产品名称, t3.PO_NUM AS 客户订单号, RTRIM(t2.工单号) AS 客户Lot, RIGHT(REPLACE(t4.SUBSTRATEID, '+', ''), 2) AS WaferNo,  " & _
'" SUM(CONVERT(INT, t9.KEY_VALUE)) AS GoodDieQty, rtrim( t2.箱号) AS Reel_Code, t3.IMAGER_CUSTOMER_REV + SUBSTRING(t3.FAB_CONV_ID,CHARINDEX('-',t3.FAB_CONV_ID),30)  AS PKG_LOT,  " & _
'" '' AS DC, CONVERT(VARCHAR(100), t1.单据日期, 23) AS 出货日期, rtrim( t6.箱号) AS 箱号, t7.keyid   from erpdata .. tblStocksqfh t1, erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,  " & _
'" erpbase .. tblmappingData t4, erpdata .. tblPackTreeInf t5, erpdata .. tblPackTreeInf t6, erpdata .. tblErpInStockDetailInfo t7, (SELECT t8.BOX_ID,  t8.WAFER_ID,   " & _
'" SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''),  2,  CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) AS wafer FROM erpdata .. tblErpInStockRelation t8) t88,  " & _
'" erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, erpdata .. tblErpInStockDetailInfo t11  where t1.单据编号 IN ('" & strShipNO & "') and t1.单据编号 = t2.单据编号  " & _
'" and t1.序号 = t2.单据项次 and t2.流程卡编号 = t4.SUBSTRATEID and t4.FILENAME = convert(varchar(50), t3.id) AND t4.LOTID = t3.SOURCE_BATCH_ID AND t5.箱号 = t2.箱号 AND t6.序号 = t5.上级序号  " & _
'" AND t7.KEY_NAME = 'CONTAINER_NAME' AND t7.KEY_VALUE = t2.箱号 AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID AND t88.wafer = t2.流程卡编号 AND t9.KEYID = t88.WAFER_ID AND t9.KEY_TYPE = 'WAFER'  " & _
'" AND t9.KEY_NAME = 'GOOD_DIE' AND t10.KEYID = t88.WAFER_ID AND t10.KEY_TYPE = 'WAFER' AND t10.KEY_NAME IN ('BAD2_DIE') AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' AND t11.KEY_NAME IN ('BAD1_DIE')   " & _
'" GROUP BY t3.MPN_DESC, t3.PO_NUM, t2.工单号, t4.SUBSTRATEID, t4.WAFER_ID, t1.单据日期, t3.IMAGER_CUSTOMER_REV, t3.FAB_CONV_ID, t6.箱号, t3.PO_NUM, t2.箱号, t7.keyid ) y LEFT JOIN erpdata..tblME_PrintInfo x  " & _
'" ON x.EVENT_ID = y.keyid AND x.LABEL_ID = 'SH103IN1' GROUP BY y.NO,y.供货方,y.客户,y.产品名称,y.客户订单号,y.客户Lot ,y.WaferNo,y.GoodDieQty,y.GoodDieQty,y.Reel_Code,y.箱号, y.PKG_LOT,y.DC,y.出货日期,y.keyid ) XX   " & _
'" LEFT JOIN erpdata..tblME_PrintInfo YY ON YY.ID  = xx.PRINT_ID ORDER BY xx.NO "

strSql = "select t3.comp_code as Supplier,t3.CUSTOMERSHORTNAME as customerCode, " & _
  "t3.MPN_DESC as Customer_Device,t3.RETICLE_LEVEL_71 as '包装品名',t3.PO_NUM as PO_Number,t3.RETICLE_LEVEL_72 as ""芯片型号"", " & _
  "t3.SOURCE_BATCH_ID as Customer_Lot, SUM(CONVERT(INT, t9.KEY_VALUE)) AS GoodDieQty, " & _
  "SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as BadDieQty, " & _
  "CONVERT(VARCHAR(100), t1.单据日期, 23) AS 出货日期, " & _
  "t6.""箱号"" " & _
  "from erpdata .. tblStocksqfh t1, erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3," & _
  "erpbase .. tblmappingData t4, erpdata .. tblPackTreeInf t5," & _
  "erpdata .. tblPackTreeInf t6, erpdata .. tblErpInStockDetailInfo t7,     (SELECT t8.BOX_ID,t8.WAFER_ID," & _
  "SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) AS wafer " & _
  "FROM  erpdata .. tblErpInStockRelation t8 " & _
  "where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, " & _
  "erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, " & _
  "erpdata .. tblErpInStockDetailInfo t11 " & _
  "where t1.单据编号 = '" & strShipNO & "'  and t1.单据编号 = t2.单据编号 " & _
  "and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID " & _
  "and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID " & _
  "AND t5.箱号 = t2.箱号  AND t6.序号 = t5.上级序号 " & _
  "AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  AND t88.BOX_ID = t7.BOX_ID " & _
  "AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  AND t9.KEY_TYPE = 'WAFER' " & _
  "AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER' " & _
  "AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID " & _
  "AND t11.KEY_TYPE = 'WAFER'  AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID " & _
  "GROUP BY t3.comp_code,t3.CUSTOMERSHORTNAME, t3.MPN_DESC,t3.RETICLE_LEVEL_71,t3.PO_NUM, " & _
  "t3.RETICLE_LEVEL_72 , t3.SOURCE_BATCH_ID, t6.箱号, CONVERT(VarChar(100), t1.单据日期, 23) "
  
Set rs = Get_SqlserveRs(strSql)
With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With

End Sub

Private Sub ListNormal_SH103()

Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset

strShipNO = UCase(Trim$(txtShipNo.text))

'strSql = " SELECT xx.NO,xx.供货方,xx.客户,xx.产品名称,xx.客户订单号 ,xx.客户Lot,xx.WaferNo,xx.GoodDieQty,xx.NgDieQty,xx.Reel_Code,xx.PKG_LOT ,  " & _
'" SUBSTRING(yy.Content,CHARINDEX('""PACKING_DATE_10"",""',yy.Content)+19,6) AS DC ,xx.出货日期,xx.箱号  FROM( SELECT y.*,MAX(x.ID) AS PRINT_ID  " & _
'" FROM ( SELECT ROW_NUMBER() over(order by REPLACE(t2.工单号, ' ', ''), t4.SUBSTRATEID) AS NO, 'HTKS' AS 供货方, 'SH103' AS 客户,  " & _
'" t3.MPN_DESC AS 产品名称, t3.PO_NUM AS 客户订单号, RTRIM(t2.工单号) AS 客户Lot, RIGHT(REPLACE(t4.SUBSTRATEID, '+', ''), 2) AS WaferNo,  " & _
'" SUM(CONVERT(INT, t9.KEY_VALUE)) AS GoodDieQty,SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) AS NgDieQty, rtrim( t2.箱号) AS Reel_Code, t3.IMAGER_CUSTOMER_REV + SUBSTRING(t3.FAB_CONV_ID,CHARINDEX('-',t3.FAB_CONV_ID),30)  AS PKG_LOT,  " & _
'" '' AS DC, CONVERT(VARCHAR(100), t1.单据日期, 23) AS 出货日期, rtrim( t6.箱号) AS 箱号, t7.keyid   from erpdata .. tblStocksqfh t1, erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,  " & _
'" erpbase .. tblmappingData t4, erpdata .. tblPackTreeInf t5, erpdata .. tblPackTreeInf t6, erpdata .. tblErpInStockDetailInfo t7, (SELECT t8.BOX_ID,  t8.WAFER_ID,   " & _
'" SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''),  2,  CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) AS wafer FROM erpdata .. tblErpInStockRelation t8) t88,  " & _
'" erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, erpdata .. tblErpInStockDetailInfo t11  where t1.单据编号 IN ('" & strShipNO & "') and t1.单据编号 = t2.单据编号  " & _
'" and t1.序号 = t2.单据项次 and t2.流程卡编号 = t4.SUBSTRATEID and t4.FILENAME = convert(varchar(50), t3.id) AND t4.LOTID = t3.SOURCE_BATCH_ID AND t5.箱号 = t2.箱号 AND t6.序号 = t5.上级序号  " & _
'" AND t7.KEY_NAME = 'CONTAINER_NAME' AND t7.KEY_VALUE = t2.箱号 AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID AND t88.wafer = t2.流程卡编号 AND t9.KEYID = t88.WAFER_ID AND t9.KEY_TYPE = 'WAFER'  " & _
'" AND t9.KEY_NAME = 'GOOD_DIE' AND t10.KEYID = t88.WAFER_ID AND t10.KEY_TYPE = 'WAFER' AND t10.KEY_NAME IN ('BAD2_DIE') AND t11.KEYID = t88.WAFER_ID AND t11.KEY_TYPE = 'WAFER' AND t11.KEY_NAME IN ('BAD1_DIE')   " & _
'" GROUP BY t3.MPN_DESC, t3.PO_NUM, t2.工单号, t4.SUBSTRATEID, t4.WAFER_ID, t1.单据日期, t3.IMAGER_CUSTOMER_REV, t3.FAB_CONV_ID, t6.箱号, t3.PO_NUM, t2.箱号, t7.keyid ) y LEFT JOIN erpdata..tblME_PrintInfo x  " & _
'" ON x.EVENT_ID = y.keyid AND x.LABEL_ID = 'SH103IN1' GROUP BY y.NO,y.供货方,y.客户,y.产品名称,y.客户订单号,y.客户Lot ,y.WaferNo,y.GoodDieQty,y.GoodDieQty,y.NgDieQty,y.Reel_Code,y.箱号, y.PKG_LOT,y.DC,y.出货日期,y.keyid ) XX   " & _
'" LEFT JOIN erpdata..tblME_PrintInfo YY ON YY.ID  = xx.PRINT_ID ORDER BY xx.NO "


strSql = "SELECT xx.NO,xx.供货方,xx.客户,xx.产品名称,xx.客户订单号 ,xx.客户Lot,xx.WaferNo,xx.GoodDieQty,xx.NgDieQty,xx.Reel_Code,xx.PKG_LOT " & _
" ,SUBSTRING(yy.Content,CHARINDEX('""PACKING_DATE_10"",""',yy.Content)+19,6) AS DC ,xx.出货日期,xx.箱号  FROM( SELECT y.*,MAX(x.ID) AS PRINT_ID  " & _
" FROM ( SELECT ROW_NUMBER() over(order by REPLACE(t2.工单号, ' ', ''), t4.SUBSTRATEID) AS NO, 'HTKS' AS 供货方, 'SH103' AS 客户  " & _
" ,   t3.MPN_DESC AS 产品名称, t3.PO_NUM AS 客户订单号, RTRIM(t2.工单号) AS 客户Lot, RIGHT(REPLACE(t4.SUBSTRATEID, '+', ''), 2) AS WaferNo  " & _
" ,   SUM(CONVERT(INT, t9.KEY_VALUE)) AS GoodDieQty,SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) AS NgDieQty  " & _
" , rtrim( t2.箱号) AS Reel_Code, t3.IMAGER_CUSTOMER_REV + SUBSTRING(t3.FAB_CONV_ID,CHARINDEX('-',t3.FAB_CONV_ID),30)  AS PKG_LOT  " & _
" ,   '' AS DC, CONVERT(VARCHAR(100), t1.单据日期, 23) AS 出货日期, rtrim( t6.箱号) AS 箱号, t7.keyid   from erpdata .. tblStocksqfh t1  " & _
" , erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3,   erpbase .. tblmappingData t4, erpdata .. tblPackTreeInf t5  " & _
" , erpdata .. tblPackTreeInf t6, erpdata .. tblErpInStockDetailInfo t7, (SELECT t8.BOX_ID,  t8.WAFER_ID  " & _
" ,    SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''),  2,  CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) AS wafer  " & _
" FROM erpdata .. tblErpInStockRelation t8) t88,   erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10  " & _
" , erpdata .. tblErpInStockDetailInfo t11  where t1.单据编号 IN ('" & strShipNO & "')  and t1.单据编号 = t2.单据编号   and t1.序号 = t2.单据项次  " & _
" and t2.流程卡编号 = t4.SUBSTRATEID and t4.FILENAME = convert(varchar(50), convert(int,t3.id)) AND t4.LOTID = t3.SOURCE_BATCH_ID AND t5.箱号 = t2.箱号  " & _
" AND t6.序号 = t5.上级序号   AND t7.KEY_NAME = 'CONTAINER_NAME' AND t7.KEY_VALUE = t2.箱号 AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  " & _
" AND t88.wafer = t2.流程卡编号 AND t9.KEYID = t88.WAFER_ID AND t9.KEY_TYPE = 'WAFER'   AND t9.KEY_NAME = 'GOOD_DIE' AND t10.KEYID = t88.WAFER_ID  " & _
" AND t10.KEY_TYPE = 'WAFER' AND t10.KEY_NAME IN ('BAD2_DIE') AND t11.KEYID = t88.WAFER_ID AND t10.BOX_ID = t88.BOX_ID  AND t11.BOX_ID = t88.BOX_ID  " & _
"  AND t11.KEY_TYPE = 'WAFER' AND t11.KEY_NAME IN ('BAD1_DIE')  " & _
"  GROUP BY t3.MPN_DESC, t3.PO_NUM, t2.工单号, t4.SUBSTRATEID, t4.WAFER_ID  " & _
"  , t1.单据日期, t3.IMAGER_CUSTOMER_REV, t3.FAB_CONV_ID, t6.箱号, t3.PO_NUM, t2.箱号, t7.keyid ) y  " & _
"   LEFT JOIN erpdata..tblME_PrintInfo x   ON x.EVENT_ID = y.keyid AND x.LABEL_ID = 'SH103IN1'  " & _
"   GROUP BY y.NO,y.供货方,y.客户,y.产品名称,y.客户订单号,y.客户Lot ,y.WaferNo,y.GoodDieQty,y.GoodDieQty,y.NgDieQty,y.Reel_Code,y.箱号  " & _
" , y.PKG_LOT,y.DC,y.出货日期,y.keyid ) XX LEFT JOIN erpdata..tblME_PrintInfo YY ON YY.ID  = xx.PRINT_ID ORDER BY xx.NO "

Set rs = Get_SqlserveRs(strSql)

With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With

End Sub

Private Sub ListNormal_SX()

Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
Dim CustName As String

strShipNO = UCase(Trim$(txtShipNo.text))

'strSql = " SELECT ROW_NUMBER() over(order by REPLACE(t2.工单号, ' ', ''), t4.SUBSTRATEID) AS NO,  'HTKS' AS 供货方, 'super pix' AS 客户, t3.MPN_DESC AS 产品名称, t3.PO_NUM AS 客户订单, " & _
'" t2.工单号 AS 客户Lot, RIGHT(REPLACE(t4.SUBSTRATEID,'+',''),2) AS  WaferNo, " & _
'" SUM(CONVERT(INT, t9.KEY_VALUE)) AS GoodDieQty, SUM(CONVERT(INT, t10.KEY_VALUE))  + SUM(CONVERT(INT, t11.KEY_VALUE)) as BadDieQty, CONVERT(VARCHAR(100),CONVERT(decimal(18, 2),  " & _
'" CONVERT(INT, SUM(CONVERT(INT, t9.KEY_VALUE))) *100.0   /( SUM(CONVERT(INT, t9.KEY_VALUE))   +SUM(CONVERT(INT, t11.KEY_VALUE))  +  SUM(CONVERT(INT, t10.KEY_VALUE))))) + '%' AS Yield , CONVERT(VARCHAR(100), t1.单据日期, 23) AS 单据日期 ," & _
'" t4.PRODUCTID AS LaserMark,  t6.箱号      from erpdata .. tblStocksqfh t1, erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3, erpbase .. tblmappingData t4, " & _
'" erpdata .. tblPackTreeInf t5, erpdata .. tblPackTreeInf t6, erpdata .. tblErpInStockDetailInfo t7,    (SELECT t8.BOX_ID,t8.WAFER_ID,  SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''),2, " & _
'" CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) AS wafer " & _
'" FROM  erpdata .. tblErpInStockRelation t8 ) t88, erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, erpdata .. tblErpInStockDetailInfo t11 " & _
'" where t1.单据编号 = '" & strShipNO & "' and t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次 and t2.流程卡编号 = t4.SUBSTRATEID  " & _
'" and t4.FILENAME = convert(varchar(50), t3.id) AND t4.LOTID = t3.SOURCE_BATCH_ID AND t5.箱号 = t2.箱号 AND t6.序号 = t5.上级序号 " & _
'" AND t7.KEY_NAME = 'CONTAINER_NAME' AND t7.KEY_VALUE = t2.箱号 AND t88.BOX_ID = t7.BOX_ID AND t9.BOX_ID = t88.BOX_ID  " & _
'" AND t88.wafer =t2.流程卡编号 AND t9.KEYID = t88.WAFER_ID AND t9.KEY_TYPE = 'WAFER' AND t9.KEY_NAME = 'GOOD_DIE' " & _
'" AND t10.KEYID = t88.WAFER_ID AND t10.KEY_TYPE = 'WAFER' AND t10.KEY_NAME IN ('BAD2_DIE') AND t11.KEYID = t88.WAFER_ID aND t11.KEY_TYPE = 'WAFER' " & _
'" AND t11.KEY_NAME IN ( 'BAD1_DIE') GROUP BY t3.MPN_DESC, t3.PO_NUM, t2.工单号, t4.SUBSTRATEID, t4.WAFER_ID, t1.单据日期, t4.PRODUCTID, t6.箱号  "
'
Select Case txtCusCode.text

    Case "SC081"
        CustName = "SC081"
    
    Case "SX", "TJ003"
        CustName = "super pix"
        
    Case Else
        CustName = ""
        
    End Select
    
strSql = "select ROW_NUMBER() over(order by REPLACE(t2.工单号, ' ', ''), t4.SUBSTRATEID) AS NO,'HTKS' AS 供货方,'" & CustName & "' AS 客户,t3.MPN_DESC AS 产品名称,t3.PO_NUM AS 客户订单, " & _
"t2.工单号 AS 客户Lot,RIGHT(REPLACE(t4.SUBSTRATEID, '+', ''), 2) AS WaferNo,SUM(CONVERT(INT, t9.KEY_VALUE)) AS GoodDieQty,SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as BadDieQty, " & _
"CONVERT(VARCHAR(100),CONVERT(decimal(18, 2),CONVERT(INT, SUM(CONVERT(INT, t9.KEY_VALUE))) * 100.0 /(SUM(CONVERT(INT, t9.KEY_VALUE)) +SUM(CONVERT(INT, t11.KEY_VALUE)) +SUM(CONVERT(INT, t10.KEY_VALUE))))) + '%' AS Yield, " & _
"CONVERT(VARCHAR(100), t1.单据日期, 23) AS 单据日期,t4.PRODUCTID AS LaserMark,t6.箱号 from erpdata .. tblStocksqfh t1  " & _
"inner join erpdata .. tblStocksqfhsub t2 on t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次 inner join erpbase .. tblmappingData t4 on t4.SUBSTRATEID = t2.流程卡编号  " & _
"inner join erpbase .. tblCustomerOI t3 on t4.FILENAME = convert(varchar(50), t3.id)  AND t4.LOTID = t3.SOURCE_BATCH_ID inner join erpdata .. tblPackTreeInf t5 on t5.箱号 = t2.箱号 " & _
"inner join erpdata .. tblPackTreeInf t6 on t6.序号 = t5.上级序号 inner join erpdata .. tblErpInStockDetailInfo t7 on t7.KEY_VALUE = t2.箱号  AND t7.KEY_NAME = 'CONTAINER_NAME' " & _
"inner join  (SELECT t8.BOX_ID,t8.WAFER_ID,SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''),2,CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) AS wafer " & _
"FROM erpdata .. tblErpInStockRelation t8) t88 on  t88.BOX_ID = t7.BOX_ID and t88.wafer = t2.流程卡编号 " & _
"inner join erpdata .. tblErpInStockDetailInfo t9 on t9.BOX_ID = t88.BOX_ID  and  t9.KEYID = t88.WAFER_ID  and t9.KEY_TYPE = 'WAFER' and t9.KEY_NAME = 'GOOD_DIE' " & _
"inner join erpdata .. tblErpInStockDetailInfo t10 on t10.BOX_ID = t88.BOX_ID  and  t10.KEYID = t88.WAFER_ID  and t10.KEY_TYPE = 'WAFER' and t10.KEY_NAME = 'BAD2_DIE' " & _
"inner join erpdata .. tblErpInStockDetailInfo t11 on t11.BOX_ID = t88.BOX_ID  and  t11.KEYID = t88.WAFER_ID  and t11.KEY_TYPE = 'WAFER' and t11.KEY_NAME = 'BAD1_DIE' " & _
"where t1.单据编号 =  '" & strShipNO & "'  GROUP BY t3.MPN_DESC,t3.PO_NUM,t2.工单号,t4.SUBSTRATEID,t4.WAFER_ID,t1.单据日期,t4.PRODUCTID,t6.箱号 "

Set rs = Get_SqlserveRs(strSql)

With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With

End Sub

Private Sub ListNormal_AC64()

Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset

strShipNO = UCase(Trim$(txtShipNo.text))

'strSql = " SELECT ROW_NUMBER() over(order by REPLACE(t2.工单号, ' ', ''), t4.SUBSTRATEID) AS NO, 'HTKS' AS 供货方, t3.MPN_DESC AS 产品名称, t3.FAB_CONV_ID AS Part_NO, t2.工单号 AS 客户Lot, " & _
'" RIGHT(REPLACE(t4.SUBSTRATEID,'+',''),2) AS  WaferNo, SUM(CONVERT(INT, t9.KEY_VALUE)) AS GoodDieQty, SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as BadDieQty, " & _
'" CONVERT(VARCHAR(100), t1.单据日期, 23) AS 单据日期, t3.IMAGER_CUSTOMER_REV AS DC, t6.箱号,t3.PO_NUM from erpdata .. tblStocksqfh t1, erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3, " & _
'" erpbase .. tblmappingData t4, erpdata .. tblPackTreeInf t5, erpdata .. tblPackTreeInf t6, erpdata .. tblErpInStockDetailInfo t7,     (SELECT t8.BOX_ID,t8.WAFER_ID, " & _
'" SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) AS wafer     FROM  erpdata .. tblErpInStockRelation t8 ) t88, " & _
'" erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, erpdata .. tblErpInStockDetailInfo t11   where t1.单据编号 = '" & strShipNO & "'  and t1.单据编号 = t2.单据编号 " & _
'" and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID  and t4.FILENAME = convert(varchar(50), t3.id)  AND t4.LOTID = t3.SOURCE_BATCH_ID  AND t5.箱号 = t2.箱号  AND t6.序号 = t5.上级序号 " & _
'" AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  AND t88.BOX_ID = t7.BOX_ID  AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  AND t9.KEY_TYPE = 'WAFER' " & _
'" AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER'  AND t10.KEY_NAME IN ('BAD2_DIE')  AND t11.KEYID = t88.WAFER_ID  AND t11.KEY_TYPE = 'WAFER'  AND t11.KEY_NAME IN ( 'BAD1_DIE') " & _
'" GROUP BY t3.MPN_DESC,   t3.PO_NUM,   t2.工单号,   t4.SUBSTRATEID,   t4.WAFER_ID,   t1.单据日期,   t3.IMAGER_CUSTOMER_REV,   t3.FAB_CONV_ID,   t6.箱号,t3.PO_NUM "



strSql = " SELECT ROW_NUMBER() over(order by REPLACE(t2.工单号, ' ', ''), t4.SUBSTRATEID) AS NO, 'HTKS' AS 供货方, t3.MPN_DESC AS 产品名称, t3.FAB_CONV_ID AS Part_NO, t2.工单号 AS 客户Lot, " & _
" RIGHT(REPLACE(t4.SUBSTRATEID,'+',''),2) AS  WaferNo, SUM(CONVERT(INT, t9.KEY_VALUE)) AS GoodDieQty, SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as BadDieQty, " & _
" CONVERT(VARCHAR(100), t1.单据日期, 23) AS 单据日期, t3.IMAGER_CUSTOMER_REV AS DC, t6.箱号,t3.PO_NUM from erpdata .. tblStocksqfh t1, erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3, " & _
" erpbase .. tblmappingData t4, erpdata .. tblPackTreeInf t5, erpdata .. tblPackTreeInf t6, erpdata .. tblErpInStockDetailInfo t7,     (SELECT t8.BOX_ID,t8.WAFER_ID, " & _
" SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) AS wafer     FROM  erpdata .. tblErpInStockRelation t8  where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, " & _
" erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, erpdata .. tblErpInStockDetailInfo t11   where t1.单据编号 = '" & strShipNO & "'  and t1.单据编号 = t2.单据编号 " & _
" and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID  and t4.FILENAME = convert(varchar(50), convert(int,t3.id))  AND t4.LOTID = t3.SOURCE_BATCH_ID  AND t5.箱号 = t2.箱号  AND t6.序号 = t5.上级序号 " & _
" AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  AND t88.BOX_ID = t7.BOX_ID  AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  AND t9.KEY_TYPE = 'WAFER' " & _
" AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER'  AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID  AND t11.KEY_TYPE = 'WAFER'  AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID " & _
" GROUP BY t3.MPN_DESC,   t3.PO_NUM,   t2.工单号,   t4.SUBSTRATEID,   t4.WAFER_ID,   t1.单据日期,   t3.IMAGER_CUSTOMER_REV,   t3.FAB_CONV_ID,   t6.箱号,t3.PO_NUM "

Set rs = Get_SqlserveRs(strSql)

With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With

End Sub

Private Sub ListNormal_DA69()

Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset

strShipNO = UCase(Trim$(txtShipNo.text))

strSql = " SELECT ROW_NUMBER() over(order by REPLACE(t2.工单号, ' ', ''), t4.SUBSTRATEID) AS NO, 'HTKS' AS 供货方, SUBSTRING(t3.MPN_DESC,1,CHARINDEX('$$',t3.MPN_DESC)-1) as Device, SUBSTRING(replace(t3.MPN_DESC,SUBSTRING(t3.MPN_DESC,1,CHARINDEX('$$',t3.MPN_DESC)-1) + '$$',''),1,CHARINDEX('$$',replace(t3.MPN_DESC,SUBSTRING(t3.MPN_DESC,1,CHARINDEX('$$',t3.MPN_DESC)-1) + '$$',''))-1) as Item, " & _
" replace(replace(t3.MPN_DESC,SUBSTRING(t3.MPN_DESC,1,CHARINDEX('$$',t3.MPN_DESC)-1) + '$$',''),SUBSTRING(replace(t3.MPN_DESC,SUBSTRING(t3.MPN_DESC,1,CHARINDEX('$$',t3.MPN_DESC)-1) + '$$',''),1,CHARINDEX('$$',replace(t3.MPN_DESC,SUBSTRING(t3.MPN_DESC,1,CHARINDEX('$$',t3.MPN_DESC)-1) + '$$',''))-1) + '$$','') As SPA, t2.工单号 AS 客户Lot, " & _
" RIGHT(REPLACE(t4.SUBSTRATEID,'+',''),2) AS  WaferNo, SUM(CONVERT(INT, t9.KEY_VALUE)) AS GoodDieQty, SUM(CONVERT(INT, t10.KEY_VALUE)) + SUM(CONVERT(INT, t11.KEY_VALUE)) as BadDieQty, " & _
" CONVERT(VARCHAR(100), t1.单据日期, 23) AS 单据日期,     '1' + right(t4.PRODUCTID,3)  AS DC, t6.箱号,t3.PO_NUM from erpdata .. tblStocksqfh t1, erpdata .. tblStocksqfhsub t2, erpbase .. tblCustomerOI t3, " & _
" erpbase .. tblmappingData t4, erpdata .. tblPackTreeInf t5, erpdata .. tblPackTreeInf t6, erpdata .. tblErpInStockDetailInfo t7,     (SELECT t8.BOX_ID,t8.WAFER_ID, " & _
" SUBSTRING(REPLACE(t8.WAFER_ID, t8.SFC_ID, ''), 2, CHARINDEX('::', REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2) AS wafer     FROM  erpdata .. tblErpInStockRelation t8  where CHARINDEX('::',REPLACE(t8. WAFER_ID, t8.SFC_ID, '')) - 2 > 0) t88, " & _
" erpdata .. tblErpInStockDetailInfo t9, erpdata .. tblErpInStockDetailInfo t10, erpdata .. tblErpInStockDetailInfo t11   where t1.单据编号 = '" & strShipNO & "'  and t1.单据编号 = t2.单据编号 " & _
" and t1.序号 = t2.单据项次  and t2.流程卡编号 = t4.SUBSTRATEID  and t4.FILENAME = convert(varchar(50), t3.id)  AND t4.LOTID = t3.SOURCE_BATCH_ID  AND t5.箱号 = t2.箱号  AND t6.序号 = t5.上级序号 " & _
" AND t7.KEY_NAME = 'CONTAINER_NAME'  AND t7.KEY_VALUE = t2.箱号  AND t88.BOX_ID = t7.BOX_ID  AND t9.BOX_ID = t88.BOX_ID  AND t88.wafer =t2.流程卡编号  AND t9.KEYID = t88.WAFER_ID  AND t9.KEY_TYPE = 'WAFER' " & _
" AND t9.KEY_NAME = 'GOOD_DIE'  AND t10.KEYID = t88.WAFER_ID  AND t10.KEY_TYPE = 'WAFER'  AND t10.KEY_NAME IN ('BAD2_DIE') AND t10.BOX_ID = t9.BOX_ID  AND t11.KEYID = t88.WAFER_ID  AND t11.KEY_TYPE = 'WAFER'  AND t11.KEY_NAME IN ( 'BAD1_DIE')  AND t11.BOX_ID = t9.BOX_ID " & _
" GROUP BY t3.MPN_DESC,   t3.PO_NUM,   t2.工单号,   t4.SUBSTRATEID,   t4.WAFER_ID,   t1.单据日期,   t3.IMAGER_CUSTOMER_REV,   t3.FAB_CONV_ID,   t6.箱号,t3.PO_NUM,t4.PRODUCTID "

Set rs = Get_SqlserveRs(strSql)

With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With

End Sub



Private Sub ListNormal_AC70()
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Dim main_lot As String
Dim lot As String
Dim Qbox As String
Dim QTY As Long
Dim dc As String


strShipNO = UCase(Trim$(txtShipNo.text))

'strSql = " select  b.箱号, d.MPN_DESC as 产品名称, d.probe_ship_part_type as 封装批次号, substring( convert(varchar(100), datepart(YY,f.ERPCREATEDATE)),3,2) + right('0' + convert(varchar(100) " & _
'" , datepart(WW,f.ERPCREATEDATE)),2) as 日期,SUM( b.数量) AS 数量,d.reticle_level_72 as 封装形式,  d.PO_NUM as 工单单号,'' as 发票号,CASE b.合格标记 WHEN '0' THEN '良品' Else '不良品' END as 良品非良品 " & _
'" From erpdata..tblStockSQfh a ,erpdata..tblStocksqfhsub b ,ERPBASE..tblmappingData c " & _
'" ,ERPBASE..tblCustomerOI d,erpdata..tblTSVwaferlist e,erpdata..tblTSVworkorder f " & _
'" WHERE b.单据编号 = a.单据编号 and b.单据项次 = a.序号 and c.SUBSTRATEID = b.流程卡编号 " & _
'"  and d.ID = c.FILENAME and e.WAFERID = c.SUBSTRATEID  and f.ORDERNAME = e.ORDERNAME and a.单据编号 = '" & strShipNO & "' " & _
'"  GROUP BY b.箱号, d.MPN_DESC, substring( convert(varchar(100), datepart(YY,f.ERPCREATEDATE)),3,2) + right('0'+convert(varchar(100) " & _
'" , datepart(WW,f.ERPCREATEDATE)),2),  d.PO_NUM, b.合格标记 ,b.流程卡编号,d.probe_ship_part_type,d.reticle_level_72 "
'
'strSql = " select  b.箱号, d.MPN_DESC as 产品名称, isnull(d.probe_ship_part_type,d.ZX_INVOICE) as 封装批次号, substring( convert(varchar(100), datepart(YY,f.ERPCREATEDATE)),3,2) + right('0' + convert(varchar(100)  " & _
'", datepart(WW,f.ERPCREATEDATE)),2) as 日期,SUM( b.数量) AS 数量,isnull(d.reticle_level_72,d.comp_code) as 封装形式,  d.PO_NUM as 工单单号,'' as 发票号,CASE b.合格标记 WHEN '0' THEN '良品' Else '不良品' END as 良品非良品  " & _
'" From erpdata..tblStockSQfh a ,erpdata..tblStocksqfhsub b ,ERPBASE..tblmappingData c  " & _
'" ,ERPBASE..tblCustomerOI d,erpdata..tblTSVwaferlist e,erpdata..tblTSVworkorder f  " & _
'" WHERE b.单据编号 = a.单据编号 and b.单据项次 = a.序号 and c.SUBSTRATEID = b.流程卡编号  " & _
'"  and d.ID = c.FILENAME and e.WAFERID = c.SUBSTRATEID  and f.ORDERNAME = e.ORDERNAME and a.单据编号 = '" & strShipNO & "'  " & _
'"  GROUP BY b.箱号, d.MPN_DESC, substring( convert(varchar(100), datepart(YY,f.ERPCREATEDATE)),3,2) + right('0'+convert(varchar(100)  " & _
'" , datepart(WW,f.ERPCREATEDATE)),2),  d.PO_NUM, b.合格标记 ,b.流程卡编号,d.probe_ship_part_type,d.reticle_level_72,d.ZX_INVOICE,d.comp_code "

'
' strsql = "SELECT mm.箱号,mm.SOURCE_batch_id,mm.产品名称,mm.封装批次号,kk.封装批次号 as '合并后批次号' ,mm.日期, mm.数量,mm.封装形式,mm.工单单号,mm.发票号,mm.良品非良品  from " & _
'"(SELECT ss.箱号,ss.SOURCE_batch_id,ss.产品名称, ss.封装批次号,ss.日期, sum(ss.数量) as 数量,ss.封装形式,ss.工单单号,ss.发票号,ss.良品非良品 from " & _
'"(select  b.箱号,d.SOURCE_batch_id,case  when d.MPN_DESC like '%\_%' escape '\' then substring(d.MPN_DESC,1,charindex('_',d.MPN_DESC) - 1) else d.MPN_DESC end as '产品名称', " & _
'"isnull(d.probe_ship_part_type,d.ZX_INVOICE) as 封装批次号, substring( convert(varchar(100) " & _
'", datepart(YY,f.ERPCREATEDATE)),3,2) + right('0' + convert(varchar(100) " & _
'", datepart(WW,f.ERPCREATEDATE)),2) as 日期,SUM(b.数量) AS 数量,g.package as 封装形式, " & _
'"d.PO_NUM as 工单单号,'' as 发票号,CASE b.合格标记 WHEN '0' THEN '良品' Else '不良品' END as 良品非良品  " & _
'"From erpdata..tblStockSQfh a ,erpdata..tblStocksqfhsub b ,ERPBASE..tblmappingData c  " & _
'",ERPBASE..tblCustomerOI d,erpdata..tblTSVwaferlist e,erpdata..tblTSVworkorder f  ,erptemp .. EU010_reference g  " & _
'"Where b.单据编号 = a.单据编号 And b.单据项次 = a.序号 And c.SUBSTRATEID = b.流程卡编号  " & _
'"and convert(varchar(50), convert(int,d.id)) = c.FILENAME and e.WAFERID = c.SUBSTRATEID  and f.ORDERNAME = e.ORDERNAME and  a.单据编号 = '" & strShipNO & "' AND g.cust_device = d.MPN_DESC  " & _
'"GROUP BY b.箱号,d.SOURCE_batch_id,d.MPN_DESC, substring( convert(varchar(100), datepart(YY,f.ERPCREATEDATE)),3,2) + right('0'+convert(varchar(100) " & _
'",datepart(WW,f.ERPCREATEDATE)),2),  d.PO_NUM, b.合格标记 ,b.流程卡编号,d.probe_ship_part_type,d.ZX_INVOICE,g.package ) ss " & _
'"group by  ss.箱号,ss.SOURCE_batch_id,ss.产品名称,ss.封装批次号, ss.日期,ss.封装形式,ss.工单单号,ss.发票号,ss.良品非良品 ) mm " & _
'"left join (SELECT zz.箱号,zz.封装批次号,Max(zz.数量) as 数量 from (SELECT b.箱号,d.SOURCE_batch_id,sum(b.数量) as 数量 ,isnull(d.probe_ship_part_type,d.ZX_INVOICE) as 封装批次号 " & _
'"from erpdata..tblStockSQfh a ,erpdata..tblStocksqfhsub b ,ERPBASE..tblmappingData c,ERPBASE..tblCustomerOI d " & _
'"Where b.单据编号 = a.单据编号 And b.单据项次 = a.序号 And c.SUBSTRATEID = b.流程卡编号  " & _
'"and convert(varchar(50), convert(int,d.id)) = c.FILENAME and a.单据编号 = '" & strShipNO & "'  group by  b.箱号,d.SOURCE_batch_id,isnull(d.probe_ship_part_type,d.ZX_INVOICE) " & _
'") zz group by zz.箱号,zz.封装批次号) kk on mm.箱号 = kk.箱号 "


strSql = "  select '' AS 'PackingList No.', b.箱号,case when d.MPN_DESC like '%\_%'  then substring(d.MPN_DESC, 1, charindex('_', d.MPN_DESC) - 1) else d.MPN_DESC end as '产品名称', " & _
         "  isnull(d.probe_ship_part_type, d.ZX_INVOICE) as 封装批次号, '' AS 合并后批次号, substring(convert(varchar(100), datepart(YY, f.ERPCREATEDATE)), 3, 2) + right('0' + convert(varchar(100) " & _
         "  , datepart(WW, f.ERPCREATEDATE)), 2) as 日期, SUM(b.数量) AS 数量, g.package as 封装形式, d.PO_NUM as 工单单号,'' as 发票号, CASE b.合格标记 WHEN '0' THEN'良品' ELSE '不良品' END as 良品非良品, " & _
         "  '' AS '毛重', '' AS '外箱尺寸','' AS 'BIN TYPE','' AS '测试程序版本','' AS '最终入库储存位置', '' AS '中转储存位置(ex:在途仓)'  FROM erpdata .. tblStockSQfh a, erpdata .. tblStocksqfhsub b, " & _
         "  ERPBASE .. tblmappingData c, ERPBASE .. tblCustomerOI d, erpdata .. tblTSVwaferlist e,erpdata .. tblTSVworkorder f,erptemp .. EU010_reference g Where b.单据编号 = a.单据编号 And b.单据项次 = a.序号 " & _
         "   And c.SUBSTRATEID = b.流程卡编号 and convert(varchar(100),  d.id) = c.FILENAME AND e.WAFERID = c.SUBSTRATEID and f.ORDERNAME = e.ORDERNAME and a.单据编号 = '" & strShipNO & "'  " & _
         "   AND g.cust_device = d.MPN_DESC GROUP BY b.箱号, d.MPN_DESC,f.ERPCREATEDATE,d.PO_NUM,b.合格标记,d.probe_ship_part_type, d.ZX_INVOICE,  b.工单号, g.package order by  b.箱号 "


Set rs = Get_SqlserveRs(strSql)




With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
    
End With

With fpS1
     For i = 1 To .MaxRows
            .Row = i
            .Col = 4
            lot = Trim(.text)
            
            .Col = 2
        Qbox = " select b.箱号,isnull(d.probe_ship_part_type, d.ZX_INVOICE) as 封装批次号, substring(convert(varchar(100), datepart(YY, f.ERPCREATEDATE)), 3, 2) + " & _
               " right('0' + convert(varchar(100), datepart(WW, f.ERPCREATEDATE)), 2) as 日期,SUM(b.数量) AS 数量 FROM erpdata .. tblStocksqfhsub b, ERPBASE .. tblmappingData c, " & _
               "  ERPBASE .. tblCustomerOI d,erpdata .. tblTSVwaferlist e, erpdata .. tblTSVworkorder f Where c.SUBSTRATEID = b.流程卡编号 and convert(varchar(100), d.id) = c.FILENAME " & _
               "   AND e.WAFERID = c.SUBSTRATEID AND f.ORDERNAME = e.ORDERNAME AND b.单据编号 = '" & strShipNO & "'  AND b.箱号 = '" & Trim(.text) & "' " & _
               "  GROUP BY b.箱号, d.probe_ship_part_type, d.ZX_INVOICE, f.ERPCREATEDATE "
               
        Set rs1 = Get_SqlserveRs(Qbox)
        
        QTY = 0
        dc = "3000"
        
         If rs1.RecordCount > 0 Then
        'If Not rs1.EOF Then
           Do While Not rs1.EOF
           If QTY < Val(rs1.Fields(3).Value) Then
            main_lot = rs1.Fields(1).Value
            
           ElseIf QTY = Val(rs1.Fields(3).Value) Then
            If Val(dc) > Val(rs1.Fields(2).Value) Then
            main_lot = rs1.Fields(1).Value
            End If
           
           End If
           QTY = Val(rs1.Fields(3).Value)
           dc = Val(rs1.Fields(2).Value)
           
           rs1.MoveNext
          Loop
            
        End If
        
        .Col = 5
        .text = main_lot
         
        Next
    
End With




End Sub





Private Sub ListNormal_GC()
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
Dim GCVersion_firstrow As String
Dim DifCodeExist As Boolean

strShipNO = UCase(Trim$(txtShipNo.text))

' strSql = "  SELECT row_number() over(order by t.lot_id, t.waferid) AS 'NO', t.sub_name AS 'Sub NAME',  'GCSH' AS 'Ship To', t.FAB_CONV_ID AS 'Fab Device',  t.cust_device AS 'Customer Device', t.gcversion AS 'GC Version', " & _
 ' " t.PO_NUM AS 'PO NO', t.invoice AS 'Invoice NO', t.create_date AS 'Ship Out Date', t.lot_id AS 'FAB Lot ID', t.waferid AS 'Wafer ID', t.gross_die AS 'Gross Dies', '' as 'Sampling Qty', " & _
 ' " ISNULL(ISNULL(t.BIN1, t.A), K.NDPW) as 'Pass Dies',  ISNULL(ISNULL(T.E, n.NDPW), 0) as 'NG Die', CONVERT(VARCHAR(10),  CONVERT(decimal(18, 2), (t.gross_die - ISNULL(ISNULL(T.E, n.NDPW), 0)) * 1.0 / (t.gross_die) * 100)) + '%' AS 'Yield', " & _
 ' " ISNULL(t.sfc, k.FIRSTNAME) as 'Pack Lot ID',  t.PRODUCTID AS 'Wafer Mark', 'A' as 'Grade', rtrim(t.箱号) as 'Carton NO',  t.大工单 AS 'WO',  '' AS 'Remark' FROM ( SELECT 'HTKS' AS sub_name, d.SHIP_SITE, " & _
 ' " RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID,  a.cust_device, d.IMAGER_CUSTOMER_REV as gcversion,d.PO_NUM, '' AS invoice, a.create_date, rtrim(a.lot_id) as lot_id, SUBSTRING(REPLACE(b.流程卡编号, '+', ''), LEN(a.lot_id) + 1, 2) as waferid, " & _
 ' " c.FAILBINCOUNT + c.PASSBINCOUNT AS gross_die, CASE WHEN n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE') THEN 'E'  ELSE 'A' END Grade, CONVERT(INT,n.KEY_VALUE ) AS qty,  c.PRODUCTID, rtrim(ay.箱号) as 箱号, " & _
 ' "  b.大工单,  a.qbox, b.流程卡编号, SUBSTRING(ee.SFC_ID, 12, 8) AS SFC  FROM erptemp .. tblshipreport_new a  INNER JOIN erpdata .. tblStockNumTree ax  ON ax.箱号 = a.qbox  INNER JOIN erpdata .. tblStockNumTree ay " & _
 ' " ON ay.序号 = ax.上级序号  INNER JOIN erpdata .. tblStocksqfhsub b ON b.单据编号 = a.ship_order  AND b.箱号 = a.qbox   AND b.工单号 = a.lot_id  INNER JOIN ERPBASE .. tblmappingData c  ON c.SUBSTRATEID = b.流程卡编号 " & _
 ' " AND c.LOTID = b.工单号 INNER JOIN erpbase .. tblCustomerOI d  ON CONVERT(VARCHAR(20),convert(int, d.ID)) = c.FILENAME  AND d.SOURCE_BATCH_ID = c.LOTID  LEFT JOIN  erpdata..tblErpInStockDetailInfo m ON m.KEY_VALUE = b.箱号 " & _
 ' " LEFT JOIN  erpdata..tblErpInStockDetailInfo n  ON n.BOX_ID = m.BOX_ID  and n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE','GOOD_DIE') and n.KEY_TYPE = 'WAFER' AND   CHARINDEX(c.SUBSTRATEID , n.KEYID ) <> 0 " & _
 ' " inner JOIN erpdata .. tblErpInStockRelation ee ON    ee.BOX_ID = n.BOX_ID AND  ee.WAFER_ID = n.KEYID  WHERE a.ship_order = '" & UCase(Trim(txtShipNo.Text)) & "' )  AS p  PIVOT(sum(qty) FOR Grade IN(A,BIN1, E)) AS T " & _
 ' " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV k  ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.流程卡编号 AND k.CONTAINERNAME LIKE '%-A' LEFT JOIN erpdata .. TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox " & _
 ' " AND L.WAFERSCRIBENUMBER = t.流程卡编号 AND L.CONTAINERNAME LIKE '%-A-01' LEFT JOIN erpdata .. TblQBOXNUMBER_TSV m  ON m.QBOXNUMBER = t.qbox  AND m.WAFERSCRIBENUMBER = t.流程卡编号 AND m.CONTAINERNAME LIKE '%-A-02' " & _
 ' " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV n  ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.流程卡编号 AND n.CONTAINERNAME LIKE '%-E' "
 
 
 strSql = "  SELECT row_number() over(order by t.lot_id, t.waferid) AS 'NO', t.sub_name AS 'Sub NAME',  " & _
 " case t.发货地址 when '嘉善' then ( case left(t.PO_NUM,2)  when 'HK' THEN 'GCZJ' WHEN 'SH' THEN 'GCZJ' Else 'GCSH'  end )else 'GCSH' end AS 'Ship To',  " & _
" t.FAB_CONV_ID AS 'Fab Device',  t.cust_device AS 'Customer Device', t.gcversion AS 'GC Version', " & _
 " t.PO_NUM AS 'PO NO', t.invoice AS 'Invoice NO', t.create_date AS 'Ship Out Date', t.lot_id AS 'FAB Lot ID', t.waferid AS 'Wafer ID', t.gross_die AS 'Gross Dies', '' as 'Sampling Qty', " & _
 " ISNULL(ISNULL(t.BIN1, t.A), K.NDPW) as 'Pass Dies',  ISNULL(ISNULL(T.E, n.NDPW), 0) as 'NG Die', CONVERT(VARCHAR(10),  CONVERT(decimal(18, 2), (t.gross_die - ISNULL(ISNULL(T.E, n.NDPW), 0)) * 1.0 / (t.gross_die) * 100)) + '%' AS 'Yield', " & _
 " ISNULL(t.sfc, k.FIRSTNAME) as 'Pack Lot ID',  t.PRODUCTID AS 'Wafer Mark', 'A' as 'Grade', rtrim(t.箱号) as 'Carton NO',  t.大工单 AS 'WO',  '' AS 'Remark' FROM ( SELECT 'HTKS' AS sub_name, d.SHIP_SITE, " & _
 " RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID,  a.cust_device, a.gcversion as gcversion,d.PO_NUM, '' AS invoice, a.create_date, rtrim(a.lot_id) as lot_id, SUBSTRING(REPLACE(b.流程卡编号, '+', ''), LEN(a.lot_id) + 1, 2) as waferid, " & _
 " c.FAILBINCOUNT + c.PASSBINCOUNT AS gross_die, CASE WHEN n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE') THEN 'E'  ELSE 'A' END Grade, CONVERT(INT,n.KEY_VALUE ) AS qty,  c.PRODUCTID, rtrim(ay.箱号) as 箱号, " & _
 "  b.大工单,f.发货地址,a.qbox, b.流程卡编号, SUBSTRING(ee.SFC_ID, 12, 8) AS SFC  FROM erptemp .. tblshipreport_new a  INNER JOIN erpdata .. tblStockNumTree ax  ON ax.箱号 = a.qbox  INNER JOIN erpdata .. tblStockNumTree ay " & _
 " ON ay.序号 = ax.上级序号  INNER JOIN erpdata .. tblStocksqfhsub b ON b.单据编号 = a.ship_order  AND b.箱号 = a.qbox   AND b.工单号 = a.lot_id   " & _
 " INNER JOIN erpdata .. tblStocksqfh f ON f.单据编号 =b.单据编号 and f.序号=b.单据项次  " & _
 " INNER JOIN ERPBASE .. tblmappingData c  ON c.SUBSTRATEID = b.流程卡编号 " & _
 " AND c.LOTID = b.工单号 INNER JOIN erpbase .. tblCustomerOI d  ON d.ID = c.FILENAME  AND d.SOURCE_BATCH_ID = c.LOTID  LEFT JOIN  erpdata..tblErpInStockDetailInfo m ON m.KEY_VALUE = b.箱号 " & _
 " LEFT JOIN  erpdata..tblErpInStockDetailInfo n  ON n.BOX_ID = m.BOX_ID  and n.KEY_NAME in ( 'BAD1_DIE','BAD2_DIE','GOOD_DIE') and n.KEY_TYPE = 'WAFER' AND   CHARINDEX(c.SUBSTRATEID , n.KEYID ) <> 0 " & _
 " inner JOIN erpdata .. tblErpInStockRelation ee ON    ee.BOX_ID = n.BOX_ID AND  ee.WAFER_ID = n.KEYID  WHERE a.ship_order = '" & UCase(Trim(txtShipNo.text)) & "'"
 If cbTpe.ListIndex = 8 Then
     strSql = strSql & " and SUBSTRING(isnull(a.gcversion,''),3,1)='" & gcrev_normal & "'"
 End If
 strSql = strSql & " )  AS p  PIVOT(sum(qty) FOR Grade IN(A,BIN1, E)) AS T " & _
 " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV k  ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.流程卡编号 AND k.CONTAINERNAME LIKE '%-A' LEFT JOIN erpdata .. TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox " & _
 " AND L.WAFERSCRIBENUMBER = t.流程卡编号 AND L.CONTAINERNAME LIKE '%-A-01' LEFT JOIN erpdata .. TblQBOXNUMBER_TSV m  ON m.QBOXNUMBER = t.qbox  AND m.WAFERSCRIBENUMBER = t.流程卡编号 AND m.CONTAINERNAME LIKE '%-A-02' " & _
 " LEFT JOIN erpdata .. TblQBOXNUMBER_TSV n  ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.流程卡编号 AND n.CONTAINERNAME LIKE '%-E' "


Set rs = Get_SqlserveRs(strSql)

With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With

With fpS1
    '查找二级代码不同的，标记出来
    If .MaxRows = 0 Then Exit Sub
    DifCodeExist = False
    .Row = 1
    .Col = E_GC_Normal.E_GCVersion
            
    GCVersion_firstrow = Trim(.text)
    For i = 2 To .MaxRows
        .Row = i
        .Col = E_GC_Normal.E_GCVersion
        If Trim(.text) <> GCVersion_firstrow Then
            .Col = -1
            .BackColor = &H80FFFF
            DifCodeExist = True
        End If
    Next
End With
If DifCodeExist = True Then
    MsgBox "单号" & UCase(Trim(txtShipNo.text)) & " 存在多种不同的二级代码！", vbInformation, "提示"
End If
End Sub

Private Sub ListWLT()

Select Case txtCusCode.text

    Case "GC"
        ListWLT_GC
    Case "SX", "TJ003", "SC081"
        ListWLT_SX

End Select

End Sub


Private Sub ListWLT_GC()

Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
Dim GCVersion_firstrow As String
Dim DifCodeExist As Boolean

strShipNO = UCase(Trim$(txtShipNo.text))

strSql = "SELECT row_number() over(order by t.lot_id,t.waferid) AS 'NO' ,t.sub_name AS 'Sub NAME', " & _
 " case t.发货地址 when '嘉善' then ( case left(t.PO_NUM,2)  when 'HK' THEN 'GCZJ' WHEN 'SH' THEN 'GCZJ' Else t.SHIP_SITE  end )else t.SHIP_SITE end AS 'Ship To',  " & _
" t.FAB_CONV_ID AS 'Fab Device',t.cust_device AS 'Customer Device',t.gcversion AS 'GC Version',t.PO_NUM AS 'PO NO',t.invoice AS 'Invoice NO',t.create_date AS 'Ship Out Date',t.lot_id AS 'FAB Lot ID',t.waferid AS 'Wafer ID',t.gross_die AS 'Gross Dies',ISNULL(t.BIN3,L.NDPW) as 'Sampling Qty' " & _
",ISNULL(ISNULL(t.BIN1,t.A),K.NDPW) as 'Pass Dies1',ISNULL(T.BIN2,m.NDPW) as 'Pass Dies2','' AS 'Pass Dies3',ISNULL(ISNULL(T.E,n.NDPW),0) as 'NG Die',CONVERT(VARCHAR(10),CONVERT(decimal(18,2), (t.gross_die - ISNULL(ISNULL(T.E,n.NDPW),0))*1.0/(t.gross_die )*100)) + '%' AS 'Yield'   " & _
",ISNULL(t.sfc,k.FIRSTNAME) as 'Pack Lot ID',t.PRODUCTID AS 'Wafer Mark','A' as 'Grade',rtrim(t.箱号) as 'Carton NO',t.大工单 AS 'WO', '' AS 'Remark' FROM ( SELECT  'HTKS' AS sub_name,d.SHIP_SITE,RTRIM(d.FAB_CONV_ID) as FAB_CONV_ID " & _
",a.cust_device,a.gcversion,d.PO_NUM,'' AS invoice, a.create_date,rtrim(a.lot_id) as lot_id,SUBSTRING(REPLACE(b.流程卡编号,'+',''),LEN(a.lot_id)+1,2) as waferid,c.FAILBINCOUNT+c.PASSBINCOUNT AS gross_die " & _
",e.GRADES,e.QTY,c.PRODUCTID,'A' as Grade,rtrim(ay.箱号) as 箱号,b.大工单,f.发货地址,a.qbox,b.流程卡编号,SUBSTRING( e.SFC,12,8) AS SFC " & _
"FROM erptemp..tblshipreport_new a INNER JOIN erpdata..tblStockNumTree ax  ON ax.箱号 =a.qbox  INNER JOIN erpdata..tblStockNumTree ay  ON ay.序号 = ax.上级序号 " & _
"INNER JOIN  erpdata..tblStocksqfhsub b  ON b.单据编号 = a.ship_order AND b.箱号 = a.qbox AND b.工单号 = a.lot_id " & _
 " INNER JOIN erpdata .. tblStocksqfh f ON f.单据编号 =b.单据编号 and f.序号=b.单据项次  " & _
"INNER JOIN  ERPBASE..tblmappingData c ON  c.SUBSTRATEID = b.流程卡编号 AND c.LOTID = b.工单号 INNER JOIN  erpbase..tblCustomerOI d ON  CONVERT(VARCHAR(20),convert(int,d.ID)) = c.FILENAME AND d.SOURCE_BATCH_ID = c.LOTID  " & _
"left JOIN  erptemp..WAFER_BIN_LIST e ON e.WAFER_ID = b.流程卡编号 inner JOIN erpdata..tblErpInStockRelation ee ON ee.SFC_ID = e.SFC  AND CHARINDEX(e.WAFER_ID,ee.WAFER_ID) <> 0    " & _
"WHERE a.ship_order = '" & UCase(Trim(txtShipNo.text)) & "'"
 If cbTpe.ListIndex = 8 Then
     strSql = strSql & " and SUBSTRING(isnull(a.gcversion,''),3,1)='" & gcrev_wlt & "'"
 End If
 strSql = strSql & " ) AS p PIVOT(sum(qty) FOR grades IN(A,BIN1,BIN2,BIN3, E)) AS T " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV k ON k.QBOXNUMBER = t.qbox AND k.WAFERSCRIBENUMBER = t.流程卡编号 AND k.CONTAINERNAME LIKE '%-A' " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV L ON L.QBOXNUMBER = t.qbox AND L.WAFERSCRIBENUMBER = t.流程卡编号 AND L.CONTAINERNAME LIKE '%-A-01' " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV m ON m.QBOXNUMBER = t.qbox AND m.WAFERSCRIBENUMBER = t.流程卡编号 AND m.CONTAINERNAME LIKE '%-A-02' " & _
"LEFT JOIN erpdata..TblQBOXNUMBER_TSV n ON n.QBOXNUMBER = t.qbox AND n.WAFERSCRIBENUMBER = t.流程卡编号 AND n.CONTAINERNAME LIKE '%-E' "

Set rs = Get_SqlserveRs(strSql)
 If cbTpe.ListIndex = 8 Then
     With fpS2
        .MaxRows = 0
        If rs.RecordCount > 0 Then
            Set .DataSource = rs
        End If
    End With
 Else
 
 
    With fpS1
        .MaxRows = 0
        If rs.RecordCount > 0 Then
            Set .DataSource = rs
        End If
    End With

    With fpS1
        '查找二级代码不同的，标记出来
        If .MaxRows = 0 Then Exit Sub
        DifCodeExist = False
        .Row = 1
        .Col = E_GC_WLT.E_GCVersion
                
        GCVersion_firstrow = Trim(.text)
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = E_GC_WLT.E_GCVersion
            If Trim(.text) <> GCVersion_firstrow Then
                .Col = -1
                .BackColor = &H80FFFF
                DifCodeExist = True
            End If
        Next
    End With
    If DifCodeExist = True Then
        MsgBox "单号" & UCase(Trim(txtShipNo.text)) & " 存在多种不同的二级代码！", vbInformation, "提示"
    End If
End If

End Sub

Private Sub ListWLT_SX()

Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset

strShipNO = UCase(Trim$(txtShipNo.text))

strSql = " SELECT ROW_NUMBER() over(order by REPLACE(x.工单号, ' ', ''), x.WAFER_ID) AS NO, x.Supplier AS 供货方, x.customer AS 客户,x.MPN_DESC AS 产品名称,x.PO_NUM AS 客户订单,REPLACE(x.工单号, ' ', '') AS 客户Lot, " & _
" x.WAFER_ID AS WaferNo,SUM(a + x.BIN1) AS GoodDieQty,SUM(BIN2) AS BIN2,SUM(BIN3) AS BIN3,SUM(BIN4) AS BIN4,SUM(BIN5) AS BIN5,SUM(BIN6) AS BIN6,SUM(BIN7) AS BIN7,SUM(BIN8) AS BIN8,SUM(BIN9) AS BIN9, " & _
" SUM(BIN10) AS BIN10, SUM(BIN11) AS BIN11,SUM(BIN12) AS BIN12,SUM(BIN13) AS BIN13,SUM(BIN14) AS BIN14,SUM(BIN15) AS BIN15,SUM(BIN16) AS BIN16,SUM(BIN17) AS BIN17,SUM(BIN18) AS BIN18,SUM(E) AS BadDieQty, " & _
" CONVERT(VARCHAR(20),CONVERT(decimal(18, 2), (x.pass - SUM(E)) * 100.0 / x.pass)) + '%' AS Yield,x.ship_date AS 出货日期,x.PRODUCTID AS LaserMark,REPLACE(x.箱号, ' ', '') AS 箱号 " & _
" FROM (SELECT t.Supplier,t.customer,t.MPN_DESC,t.PO_NUM,t.工单号,t.WAFER_ID,ISNULL(t.a, 0) AS a,ISNULL(t.BIN1, 0) AS BIN1,ISNULL(t.BIN2, 0) AS BIN2,ISNULL(t.BIN3, 0) AS BIN3,ISNULL(t.BIN4, 0) AS BIN4, " & _
" ISNULL(t.BIN5, 0) AS BIN5,ISNULL(t.BIN6, 0) AS BIN6,ISNULL(t.BIN7, 0) AS BIN7,ISNULL(t.BIN8, 0) AS BIN8,ISNULL(t.BIN9, 0) AS BIN9,ISNULL(t.BIN10, 0) AS BIN10,ISNULL(t.BIN11, 0) AS BIN11, " & _
" ISNULL(t.BIN12, 0) AS BIN12,ISNULL(t.BIN13, 0) AS BIN13,ISNULL(t.BIN14, 0) AS BIN14,ISNULL(t.BIN15, 0) AS BIN15,ISNULL(t.BIN16, 0) AS BIN16,ISNULL(t.BIN17, 0) AS BIN17,ISNULL(t.BIN18, 0) AS BIN18, " & _
" ISNULL(t.E, 0) AS E,t.pass,CONVERT(VARCHAR(100), t.审核日期, 111) AS ship_date, t.PRODUCTID,t.箱号 " & _
" FROM (SELECT 'HTKS' AS Supplier,'super pix' AS customer,t3.MPN_DESC,t3.PO_NUM,t2.工单号,t4.WAFER_ID, t6.QTY,t6.GRADES,t4.PASSBINCOUNT + t4.FAILBINCOUNT AS pass,t1.审核日期,t4.PRODUCTID,t8.箱号,t2.流程卡编号 " & _
" from erpdata .. tblStocksqfh t1,erpdata .. tblStocksqfhsub t2,erpbase .. tblCustomerOI t3,erpbase .. tblmappingData t4, erptemp .. WAFER_BIN_LIST t6,erpdata .. tblErpInStockRelation t66,erpdata .. tblPackTreeInf t7, " & _
" erpdata .. tblPackTreeInf t8, erpdata .. tblErpInStockDetailInfo t9 where t1.单据编号 = t2.单据编号 and t1.序号 = t2.单据项次 and t2.流程卡编号 = t4.SUBSTRATEID and t4.FILENAME = convert(varchar(50), convert(int,t3.id) " & _
" AND t4.LOTID = t3.SOURCE_BATCH_ID and t6.WAFER_ID = t2.流程卡编号 and t2.箱号 = t7.箱号 and t7.上级序号 = t8.序号 AND t66.SFC_ID = t6.SFC AND SUBSTRING(REPLACE(t66.WAFER_ID, t66.SFC_ID, ''),2,CHARINDEX('::',REPLACE(t66. WAFER_ID, t66.SFC_ID,'')) - 2) = t6.WAFER_ID " & _
" AND t9.KEY_VALUE = t2.箱号 AND t9.BOX_ID = t66.BOX_ID and t1.单据编号 = '" & strShipNO & "') as tt PIVOT(sum(qty) FOR grades IN(A,BIN1,BIN2,BIN3,BIN4,BIN5,BIN6,BIN7,BIN8,BIN9,BIN10,BIN11,BIN12,BIN13,BIN14,BIN15,BIN16,BIN17,BIN18,E)) AS T) x " & _
" GROUP BY x.Supplier,x.customer,x.MPN_DESC, x.PO_NUM,REPLACE(x.工单号, ' ', ''),x.WAFER_ID,x.ship_date,x.PRODUCTID,REPLACE(x.箱号, ' ', ''),x.pass "

Set rs = Get_SqlserveRs(strSql)

With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With

End Sub

Private Sub ListWLA()

Select Case txtCusCode.text

    Case "GC"
        ListWLA_GC
    Case "SX", "TJ003", "SC081"
        ListWLA_SX

End Select

End Sub

Private Sub ListWLA_GC()
Dim strShipNO_SQ As String
Dim strShipNO As String
Dim ShipTo As String
Dim strSql As String
Dim rs As New ADODB.Recordset
Dim SMR As New ADODB.Recordset
Dim i As Integer

Dim GCVersion_firstrow As String
Dim CustPN_firstrow As String
Dim DifCodeExist As Boolean
Dim DifCustPnExist As Boolean

Dim lot As String
Dim WAFER As String
Dim sqllot As String
Dim Rs_lot As New ADODB.Recordset




strShipNO = UCase(Trim$(txtShipNo.text))
ShipTo = UCase(Trim$(Cbshipto.text))

If strShipNO = "" Then Exit Sub

If UCase(Left(strShipNO, 3)) = "FWW" Then
    strShipNO_SQ = strShipNO
    strShipNO = ""
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Source = "select distinct remark2 from erptemp..tblstockdb_temp where ORDER_NUM='" & strShipNO_SQ & "'"
    SMR.Open , INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            If strShipNO = "" Then
                strShipNO = "'" & (Trim(SMR("remark2"))) & "'"
            Else
                strShipNO = strShipNO & "，'" & (Trim(SMR("remark2"))) & "'"
            End If
            SMR.MoveNext
        Next
    Else
        MsgBox "单号有误或还未执行调拨", vbInformation, "提示"
        Exit Sub
    End If
    SMR.Close
    Set SMR = Nothing
Else
    strShipNO = "'" & strShipNO & "'"
End If
'
'    strSql = "SELECT  row_number() OVER(ORDER BY X.[CST ID],X.[Wafer ID]) AS [NO],X.* FROM ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID],[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)), " & _
'" rtrim(ltrim(a.流程卡编号))) as [Wafer Qty], " & _
'" CASE '" & ShipTo & "' when '上海' then (case  LEFT(B.PO_NUM,2) WHEN 'HK' THEN 'WHT' WHEN 'SH' THEN (CASE LEFT(B.PO_NUM,5)  WHEN 'SHGCW' THEN 'SH' ELSE 'SWHT' END ) ELSE 'SH' END)  when '嘉善' then ( case  LEFT(B.PO_NUM,2) WHEN 'SH' Then ( CASE LEFT(B.PO_NUM,5)  WHEN 'SHGCW' THEN 'ZSH' ELSE 'SWHT' END) else 'SH' end )ELSE 'SH' END   as [Bond Pro], " & _
'" B.PO_NUM AS [PO NO],'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.工单号 as [FAB Lot ID], " & _
'" right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID], a.合格数 as [Gross Dies], '0' as [Sampling Qty],a.合格数 as [Pass Dies],0 as [NG Die],''as [Yield],c.FIRSTNAME as [Pack Lot ID], d.productid as [Wafer Mark], " & _
'" 'A' AS Grade,c.QBOXNUMBER as [Carton NO],f.ORDERNAME as WO,'' as [Remark] FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV c  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.调拨编号 ='" & billNoTemp & "' " & _
'" and b.SOURCE_BATCH_ID=a.工单号 and d.filename = cast(b.ID as nvarchar) and c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 and d.SUBSTRATEID=a.流程卡编号 and f.WAFERID=a.流程卡编号 )X" & _
'" union SELECT  row_number() OVER(ORDER BY Y.[CST ID],Y.[Wafer ID]) AS [NO],Y.* FROM ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device], replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
'" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID],[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)), " & _
'" rtrim(ltrim(a.流程卡编号))) as [Wafer Qty],  " & _
'" CASE '" & ShipTo & "' when '上海' then (case  LEFT(B.PO_NUM,2) WHEN 'HK' THEN 'WHT' WHEN 'SH' THEN (CASE LEFT(B.PO_NUM,5)  WHEN 'SHGCW' THEN 'SH' ELSE 'SWHT' END ) ELSE 'SH' END)  when '嘉善' then ( case  LEFT(B.PO_NUM,2) WHEN 'SH' Then ( CASE LEFT(B.PO_NUM,5)  WHEN 'SHGCW' THEN 'ZSH' ELSE 'SWHT' END) else 'SH' end )ELSE 'SH' END   as [Bond Pro], " & _
'" B.PO_NUM AS [PO NO],'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.工单号 as [FAB Lot ID], " & _
'" right(rtrim(ltrim(replace(a.流程卡编号,'+',''))),2) as [Wafer ID], a.合格数 as [Gross Dies], '0' as [Sampling Qty],a.合格数 as [Pass Dies],0 as [NG Die],''as [Yield],REPLACE(BB.SFC_ID,'SFCBO:1020,','') as [Pack Lot ID], d.productid as [Wafer Mark], " & _
'" 'A' AS Grade,A.箱号 as [Carton NO],f.ORDERNAME as WO,'' as [Remark] FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata..tblErpInStockDetailInfo aa,erpdata..tblErpInStockRelation bb  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.调拨编号 in (" & strShipNO & ") " & _
'" and b.SOURCE_BATCH_ID=a.工单号 and d.filename = cast(b.ID as nvarchar)  and d.SUBSTRATEID=a.流程卡编号 and f.WAFERID=a.流程卡编号  and a.箱号 = aa.KEY_VALUE and bb.BOX_ID = aa.BOX_ID and  SUBSTRING(replace(bb.WAFER_ID,'SFCBO:1020,','') " & _
'", CHARINDEX(',', replace(bb.WAFER_ID,'SFCBO:1020,',''))+1,CHARINDEX('::', replace(bb.WAFER_ID,'SFCBO:1020,',''))-CHARINDEX(',', replace(bb.WAFER_ID,'SFCBO:1020,',''))-1) = a.流程卡编号 )Y"
'

   strSql = "SELECT  row_number() OVER(ORDER BY X.[FAB Lot ID],X.[Wafer ID]) AS [NO],X.* FROM( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device] " & _
           " , replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version],  [erpdata].[dbo].[Get_TSV_GCWLA_LotID_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)) " & _
           " ,rtrim(ltrim(a.流程卡编号))) as [CST ID] ,[erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_NewDB](a.调拨编号,rtrim(ltrim(a.工单号)),  rtrim(ltrim(a.流程卡编号))) as [Wafer Qty] " & _
          " ,  CASE '嘉善' when '上海' then (case  LEFT(B.PO_NUM,2) WHEN 'HK' THEN 'WHT' WHEN 'SH' THEN (CASE LEFT(B.PO_NUM,5) WHEN 'SHGCW' THEN 'SH' ELSE 'SWHT' END ) ELSE 'SH' END) " & _
         "  WHEN '嘉善' then ( case  LEFT(B.PO_NUM,2)  WHEN 'SH' Then ( CASE LEFT(B.PO_NUM,5)  WHEN 'SHGCW' THEN 'ZSH' ELSE 'SWHT' END) else 'SH' end )ELSE 'SH' END   as [Bond Pro] " & _
         "  ,  B.PO_NUM AS [PO NO],'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.工单号 as [FAB Lot ID] ,  right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID] " & _
         " , a.合格数 as [Gross Dies], '0' as [Sampling Qty],a.合格数 as [Pass Dies],0 as [NG Die],''as [Yield],c.FIRSTNAME as [Pack Lot ID], d.productid as [Wafer Mark],  'A' AS Grade " & _
        " ,c.QBOXNUMBER as [Carton NO],f.ORDERNAME as WO,'' as [Remark],a.流程卡编号 FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV c " & _
        "  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.调拨编号 in (" & strShipNO & ")  and b.SOURCE_BATCH_ID=a.工单号 and d.filename = cast(b.ID as nvarchar) " & _
        " AND c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 and d.SUBSTRATEID=a.流程卡编号  AND f.WAFERID=a.流程卡编号 )X  UNION " & _
       "   SELECT  row_number() OVER(ORDER BY Y.[FAB Lot ID],Y.[Wafer ID]) AS [NO],Y.*  FROM ( SELECT DISTINCT 'HTKS' as [Sub Name],'GC_LG' as [Ship To],B.FAB_CONV_ID AS [Fab Device] " & _
       "  , replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version] , '' AS  [CST ID] ,'' AS [Wafer Qty] ,   CASE '嘉善' when '上海' then (case  LEFT(B.PO_NUM,2) " & _
       "  WHEN 'HK' THEN 'WHT' WHEN 'SH' THEN (CASE LEFT(B.PO_NUM,5)   WHEN 'SHGCW' THEN 'SH' ELSE 'SWHT' END ) ELSE 'SH' END)  when '嘉善' then ( case  LEFT(B.PO_NUM,2) " & _
       "  WHEN 'SH' Then ( CASE LEFT(B.PO_NUM,5)  WHEN 'SHGCW' THEN 'ZSH' ELSE 'SWHT' END) else 'SH' end )ELSE 'SH' END   as [Bond Pro] ,  B.PO_NUM AS [PO NO] " & _
      "  ,'' [Invoice NO],convert(varchar(100), getdate(), 111) AS [Ship Out Date], A.工单号 as [FAB Lot ID] ,  right(rtrim(ltrim(replace(a.流程卡编号,'+',''))),2) as [Wafer ID] " & _
      "   , a.合格数 as [Gross Dies], '0' as [Sampling Qty]  ,a.合格数 as [Pass Dies],0 as [NG Die],''as [Yield],SUBSTRING( REPLACE(ab.KEYID, 'SFCBO:1020,', ''),1 " & _
      "   ,CHARINDEX(rtrim(a.流程卡编号),REPLACE(ab.KEYID, 'SFCBO:1020,', ''))-2) as 'Pack Lot ID' , d.productid as [Wafer Mark],  'A' AS Grade,A.箱号 as [Carton NO],f.ORDERNAME as WO " & _
      "  ,'' as [Remark],a.流程卡编号  FROM   erpdata.dbo.tblStockdbsub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata..tblErpInStockDetailInfo aa,erpdata..tblErpInStockDetailInfo ab " & _
      "  ,[ERPBASE].[dbo].[tblmappingData] d,[erpdata].[dbo].[tblTSVwaferlist] f WHERE a.调拨编号 in (" & strShipNO & ")  and b.SOURCE_BATCH_ID=a.工单号  AND d.filename = cast(b.ID as nvarchar) " & _
      "  and d.SUBSTRATEID=a.流程卡编号 and f.WAFERID=a.流程卡编号  and a.箱号 = aa.KEY_VALUE  AND ab.BOX_ID = aa.BOX_ID AND ab.KEY_TYPE = 'WAFER' AND ab.KEY_VALUE = a.流程卡编号 )Y "






Set rs = Get_SqlserveRs(strSql)

With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
    
  For i = 1 To .MaxRows
       .Row = i
       .Col = 13
        lot = .text
       .Col = 26
        WAFER = .text
        sqllot = "SELECT  ISNULL(erpdata.dbo.Get_TSV_GCWLA_LotID_NewDB(" & strShipNO & ",rtrim(ltrim('" & lot & "')) ,rtrim(ltrim('" & WAFER & "'))),'') as [CST ID] " & _
                  " , ISNULL(erpdata.dbo.Get_TSV_GCWLA_LotIDQty_NewDB(" & strShipNO & ",rtrim(ltrim('" & lot & "'))  ,  rtrim(ltrim('" & WAFER & "'))),'')  as [Wafer Qty]"
        
        If Rs_lot.State = adStateOpen Then Rs_lot.Close
        Rs_lot.Open sqllot, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        
        .Col = 7
        .text = Rs_lot.Fields(0).Value
        .Col = 8
        .text = Rs_lot.Fields(1).Value
         .Col = 26
        .text = ""
  
  Next
  
    
  MsgBox "查询完成", vbInformation, "提示"


End With


With fpS1
    '查找二级代码不同的，标记出来
    If .MaxRows = 0 Then Exit Sub
    DifCodeExist = False
    DifCustPnExist = False
    .Row = 1
    .Col = E_GC_WLA.E_GCVersion
            
    GCVersion_firstrow = Trim(.text)
    
    .Row = 1
    .Col = E_GC_WLA.E_CustomerDevice

    CustPN_firstrow = Trim(.text)
    For i = 1 To .MaxRows
        .Row = i
        .Col = E_GC_WLA.E_GCVersion
        If Trim(.text) <> GCVersion_firstrow Then
            .Col = -1
            .BackColor = &H80FFFF
            DifCodeExist = True
        End If
        
        .Col = E_GC_WLA.E_CustomerDevice
        If Trim(.text) <> CustPN_firstrow Then
            .Col = -1
            .BackColor = &H80FFFF
            DifCustPnExist = True
        End If
        
    Next
End With
If DifCodeExist = True Then
    MsgBox "单号" & UCase(Trim(txtShipNo.text)) & " 存在多种不同的二级代码！", vbInformation, "提示"
End If
If DifCustPnExist = True Then
    MsgBox "单号" & UCase(Trim(txtShipNo.text)) & " 存在多种不同的客户机种！", vbInformation, "提示"
End If

End Sub

Private Sub ListWLA_SX()
Dim strShipNO_SQ As String
Dim strShipNO As String
Dim strSql As String
Dim rs As New ADODB.Recordset
Dim SMR As New ADODB.Recordset
Dim i As Integer


strShipNO = UCase(Trim$(txtShipNo.text))

If strShipNO = "" Then Exit Sub

If UCase(Left(strShipNO, 3)) = "FWW" Then
    strShipNO_SQ = strShipNO
    strShipNO = ""
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Source = "select distinct remark2 from erptemp..tblstockdb_temp where ORDER_NUM='" & strShipNO_SQ & "'"
    SMR.Open , INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            If strShipNO = "" Then
                strShipNO = "'" & (Trim(SMR("remark2"))) & "'"
            Else
                strShipNO = strShipNO & "，'" & (Trim(SMR("remark2"))) & "'"
            End If
            SMR.MoveNext
        Next
    Else
        MsgBox "单号有误或还未执行调拨", vbInformation, "提示"
        Exit Sub
    End If
    SMR.Close
    Set SMR = Nothing
Else
    strShipNO = "'" & strShipNO & "'"
End If


strSql = "SELECT row_number() OVER(ORDER BY X.[客户Lot], X.[WaferNo]) AS [ NO ], " & _
"X.* FROM (SELECT DISTINCT 'HTKS' as [供货方], " & _
"'super pix' as [客户],  " & _
"B.FAB_CONV_ID AS [产品名称],  " & _
"B.PO_NUM AS [客户订单号], " & _
"rtrim(A.工单号) as [客户Lot], " & _
"right(rtrim(ltrim(a.流程卡编号)), 2) as [WaferNo], " & _
"a.合格数 as [GoodDieQty], " & _
"0 as [BadDieQty], " & _
"'100%' as [Yield], " & _
"convert(varchar(100), getdate(), 111) AS [出货日期],  " & _
"d.productid as [LaserMark], " & _
"a.箱号, " & _
"'' as [备注]  " & _
"FROM erpdata.dbo.tblStockdbsub a, " & _
"[ERPBASE].[dbo].[tblCustomeroi] b,  " & _
"[ERPBASE].[dbo].[tblmappingData] d, " & _
"[erpdata].[dbo].[tblTSVwaferlist] f " & _
"WHERE a.调拨编号 = '" & strShipNO & "' " & _
"AND b.SOURCE_BATCH_ID = a.工单号 and d.filename = cast(b.ID as nvarchar) " & _
"AND d.SUBSTRATEID = a.流程卡编号 and f.WAFERID = a.流程卡编号) X"

Set rs = Get_SqlserveRs(strSql)

With fpS1
    .MaxRows = 0
    If rs.RecordCount > 0 Then
            Set .DataSource = rs
    End If
End With

End Sub

Private Sub ListBUMP()

    Dim strShipNO As String

    Dim strSql    As String

    Dim rs        As New ADODB.Recordset

    Dim i         As Integer, strLotIDTmp As String

    strShipNO = UCase(Trim$(txtShipNo.text))
    
strSql = "select '' as 入库时间,'' as 入库单编号,'' as 工单号,'' as 箱号,'' as 品名,'' as 料号,'' as 厂内批号, Lot_id,'小计' as WaferID,SUM(CONVERT(INT, 客户设计GOODDIE)) as 客户设计GROSSDIE " & _
 ", SUM(CONVERT(INT,GOODDIE数量)) as GOODDIE数量,SUM(CONVERT(INT,良品进站片数)) as 良品进站片数,SUM(CONVERT(INT, 厂内NGDIE数量)) as 厂内NGDIE数量, " & _
 "SUM(CONVERT(INT,客户NGDIE数量)) as 客户NGDIE数量, '' as 镭射码,'' as 客户代码 from " & _
 "(select distinct CONVERT(VARCHAR(100), aa.入库时间, 23) AS 入库时间,aa.入库单编号 as 入库单编号,cc.大工单 as 工单号,bb.箱号 as 箱号,ee.SALESORDER as 品名, " & _
 "ee.PRODUCT as 料号,REPLACE(ff.SFC_ID, 'SFCBO:1020,', '') AS 厂内批号,cc.工单号 as LOT_ID,cc.流程卡编号 as WAFER_ID,dd.DIEQTY as 客户设计GOODDIE, " & _
 "gg.KEY_VALUE as GOODDIE数量,1 as 良品进站片数,hh.KEY_VALUE as 厂内NGDIE数量,jj.KEY_VALUE as 客户NGDIE数量,dd.MARKINGCODE as 镭射码内容,ee.CUSTOMER as 客户代码 " & _
 "from erpdata .. tblPackToHouse aa " & _
 "inner join erpdata .. tblPackTreeInf bb on aa.入库单编号 = bb.入库单编号 and bb.上级序号 <> '0' " & _
 "inner join erpdata .. tblPackMainInfSub cc on bb.箱号 = cc.箱号 " & _
 "inner join erpdata .. tblTSVwaferlist dd on dd.WAFERLOT = cc.工单号 and dd.WAFERID = cc.流程卡编号 " & _
 "inner join erpdata .. tblTSVworkorder ee on dd.ORDERNAME = ee.ORDERNAME " & _
 "INNER JOIN erpdata .. tblErpInStockDetailInfo kk ON kk.KEY_VALUE = cc.箱号 AND kk.KEY_NAME = 'CONTAINER_NAME' " & _
 "inner join erpdata .. tblErpInStockRelation ff on SUBSTRING(REPLACE(ff.WAFER_ID, 'SFCBO:1020,', ''),CHARINDEX(',', REPLACE(ff.WAFER_ID, 'SFCBO:1020,', '')) + 1, " & _
 "CHARINDEX('::', REPLACE(ff.WAFER_ID, 'SFCBO:1020,', '')) -CHARINDEX(',', REPLACE(ff.WAFER_ID, 'SFCBO:1020,', '')) - 1) =cc.流程卡编号 AND kk.BOX_ID = ff.BOX_ID " & _
 "inner join erpdata .. tblErpInStockDetailInfo gg on gg.KEYID = ff.WAFER_ID and gg.BOX_ID = ff.BOX_ID and gg.KEY_NAME = 'GOOD_DIE' " & _
 "inner join erpdata .. tblErpInStockDetailInfo hh on hh.KEYID = ff.WAFER_ID and hh.BOX_ID = ff.BOX_ID and hh.KEY_NAME = 'BAD1_DIE' " & _
 "inner join erpdata .. tblErpInStockDetailInfo jj on jj.KEYID = ff.WAFER_ID and jj.BOX_ID = ff.BOX_ID and jj.KEY_NAME = 'BAD2_DIE' " & _
 "where aa.入库单编号 = '" & strShipNO & "') SS group by LOT_ID union "

strSql = strSql & "select distinct CONVERT(VARCHAR(100), aa.入库时间, 23) AS 入库时间,aa.入库单编号 as 入库单编号,cc.大工单 as 工单号,bb.箱号 as 箱号,ee.SALESORDER as 品名, " & _
 "ee.PRODUCT as 料号,REPLACE(ff.SFC_ID, 'SFCBO:1020,', '') AS 厂内批号,cc.工单号 as LOT_ID,cc.流程卡编号 as WaferID,dd.DIEQTY as 客户设计GOODDIE, " & _
 "gg.KEY_VALUE as GOODDIE数量,1 as 良品进站片数,hh.KEY_VALUE as 厂内NGDIE数量,jj.KEY_VALUE as 客户NGDIE数量,dd.MARKINGCODE as 镭射码内容,ee.CUSTOMER as 客户代码 " & _
 "from erpdata .. tblPackToHouse aa " & _
 "inner join erpdata .. tblPackTreeInf bb on aa.入库单编号 = bb.入库单编号 and bb.上级序号 <> '0' " & _
 "inner join erpdata .. tblPackMainInfSub cc on bb.箱号 = cc.箱号 " & _
 "inner join erpdata .. tblTSVwaferlist dd on dd.WAFERLOT = cc.工单号 and dd.WAFERID = cc.流程卡编号 " & _
 "inner join erpdata .. tblTSVworkorder ee on dd.ORDERNAME = ee.ORDERNAME " & _
 "INNER JOIN erpdata .. tblErpInStockDetailInfo kk ON kk.KEY_VALUE = cc.箱号 AND kk.KEY_NAME = 'CONTAINER_NAME' " & _
 "inner join erpdata .. tblErpInStockRelation ff on SUBSTRING(REPLACE(ff.WAFER_ID, 'SFCBO:1020,', ''),CHARINDEX(',', REPLACE(ff.WAFER_ID, 'SFCBO:1020,', '')) + 1, " & _
 "CHARINDEX('::', REPLACE(ff.WAFER_ID, 'SFCBO:1020,', '')) -CHARINDEX(',', REPLACE(ff.WAFER_ID, 'SFCBO:1020,', '')) - 1) =cc.流程卡编号 AND kk.BOX_ID = ff.BOX_ID " & _
 "inner join erpdata .. tblErpInStockDetailInfo gg on gg.KEYID = ff.WAFER_ID and gg.BOX_ID = ff.BOX_ID and gg.KEY_NAME = 'GOOD_DIE' " & _
 "inner join erpdata .. tblErpInStockDetailInfo hh on hh.KEYID = ff.WAFER_ID and hh.BOX_ID = ff.BOX_ID and hh.KEY_NAME = 'BAD1_DIE' " & _
 "inner join erpdata .. tblErpInStockDetailInfo jj on jj.KEYID = ff.WAFER_ID and jj.BOX_ID = ff.BOX_ID and jj.KEY_NAME = 'BAD2_DIE' " & _
 "where aa.入库单编号 = '" & strShipNO & "' order by LOT_ID,WaferID "

    Set rs = Get_SqlserveRs(strSql)
    
    With fpS1
        .MaxRows = 0
        If rs.RecordCount > 0 Then
            Set .DataSource = rs
        End If
    End With

End Sub

Private Sub OnExport()
    Dim strbond_shipto_temp As String
    Dim strbond_temp As String
    Dim strshipto_temp As String
    Dim DirShare1 As String
    Dim DirShare2 As String
    Dim str_Date As String
    If fpS1.MaxRows = 0 And fpS2.MaxRows = 0 Then
        MsgBox "没有数据可以导出", vbInformation, "提示"
        Exit Sub
    End If

    If cbExcel.text = "" Then
        MsgBox "请选择导出格式", vbInformation, "提示"
        Exit Sub
    End If
    str_Date = Month(Now()) & "." & Day(Now())
  
    DirShare1 = "C:\others\" & str_Date '输出excel路径
    
    If cbTpe.text = "GC Shippinglist" Then '分保税非保税及出货地再分别查询，分别输出
        Dim strbond_shipto As String
        Dim strbond_shipto_list As String
        With fpS1
        For i = 1 To .MaxRows
            strbond_shipto = ""
            .Col = 7
            .Row = i
            If Left(Trim(.text), 2) = "HK" Then
                strbond_shipto = "保税"
            ElseIf Left(Trim(.text), 2) = "SH" Then
                strbond_shipto = "非保税"
            Else
                strbond_shipto = "不确定"
            End If
            .Col = 25
            .Row = i
             strbond_shipto = strbond_shipto & "#" & Trim(.text)
            If InStr(strbond_shipto_list, strbond_shipto) Then
            Else
                If strbond_shipto_list = "" Then
                    strbond_shipto_list = strbond_shipto
                Else
                    strbond_shipto_list = strbond_shipto_list & "," & strbond_shipto
                End If
            End If
        Next
        End With
    

        For i = 0 To UBound(Split(strbond_shipto_list, ","))
            strbond_shipto_temp = Split(strbond_shipto_list, ",")(i)
            strbond_temp = Split(strbond_shipto_temp, "#")(0)
            strshipto_temp = Split(strbond_shipto_temp, "#")(1)

            If Dir(DirShare1, vbDirectory) = "" Then
                MkDir DirShare1                                      '日期文件夹
            End If
            DirShare2 = DirShare1 & "\" & str_Date & strbond_temp & strshipto_temp
           
            If Dir(DirShare2, vbDirectory) = "" Then
                MkDir DirShare2                                      '保税及出货地文件夹
            End If
            Call ListGCShippinglist(strbond_temp, strshipto_temp)
            Call ExportExcel(cbTpe.text, DirShare2, fpS1)
            
        Next
        Call ListGCShippinglist("", "")

    ElseIf cbTpe.text = "GC WLA Shippinglist" Then '分保税非保税,分别输出
        DirShare1 = "C:\others\" & str_Date
        If Dir(DirShare1, vbDirectory) = "" Then
            MkDir DirShare1                                      '日期文件夹
        End If
        
        Call ListGCShippinglist_WLA("保税")
        If fpS1.MaxRows > 0 Then
            DirShare2 = DirShare1 & "\" & str_Date & "WLA保税"
           
            If Dir(DirShare2, vbDirectory) = "" Then
                MkDir DirShare2                                      '保税及出货地文件夹
            End If
            Call ExportExcel(cbTpe.text, DirShare2, fpS1)
        End If
        Call ListGCShippinglist_WLA("非保税")
        If fpS1.MaxRows > 0 Then
            DirShare2 = DirShare1 & "\" & str_Date & "WLA非保税"
            If Dir(DirShare2, vbDirectory) = "" Then
                MkDir DirShare2                                      '保税及出货地文件夹
            End If
            Call ExportExcel(cbTpe.text, DirShare2, fpS1)
        End If
        Call ListGCShippinglist_WLA("")
    
    Else

        DirShare1 = "C:\others\" & str_Date
        If Dir(DirShare1, vbDirectory) = "" Then
            MkDir DirShare1                                      '日期文件夹
        End If
        If UCase(cbTpe.text) = "GC NOR/WLT" Then
            If fpS1.MaxRows > 0 Then 'NORMAL
                Call ExportExcel("正常出货", DirShare1, fpS1)
            End If
            If fpS2.MaxRows > 0 Then 'WLT
                Call ExportExcel("WLT出货", DirShare1, fpS2)
            End If
        
        Else

        Call ExportExcel(cbTpe.text, DirShare1, fpS1)
        End If
    End If

End Sub

Private Sub ExportExcel(excelfomat As String, outputpath As String, fps As fpSpread)
    Dim xlsApp      As Excel.Application
    Dim xlsBook     As Excel.Workbook
    Dim xlsSheet    As Excel.Worksheet
    Dim i           As Long
    Dim j           As Long
    Dim strFileName As String
    Dim strPartName As String
    Dim strCustPN As String
    Dim xlsApp1      As Excel.Application
    Dim str_eachrow  As String
    Dim objStream    As ADODB.Stream
    
    On Error GoTo Ert
    
    
    With fps
    Select Case excelfomat
        
        Case "正常出货"
            .Row = 1
            .Col = E_GC_Normal.E_CustomerDevice
            strCustPN = Trim(.text)
            strPartName = "\" & strCustPN & "_PL_HTKS_CSP_"

        Case "WLT出货"
            .Row = 1
            .Col = E_GC_WLT.E_CustomerDevice
            strCustPN = Trim(.text)
            strPartName = "\" & strCustPN & "_PL_HTKS_WLT_"

        Case "WLA出货"
            .Row = 1
            .Col = E_GC_WLA.E_CustomerDevice
            strCustPN = Trim(.text)
            strPartName = "\" & strCustPN & "_PL_HTKS_WLA_"
        
        Case "BUMPING入库"
            strPartName = "\WH_HTKS_BUMPING_"
            
        Case "GC Shippinglist"
            strPartName = "\Shippinglist_HTKS_"
        Case "68 Shippinglist", "HK006 Shippinglist"
            strPartName = "\" & txtCusCode.text & "Shippinglist_HTKS_"
        Case "GC WLA Shippinglist"
            strPartName = "\WLA_Shippinglist_HTKS_"
        Case "GC list汇总"
            strPartName = "\list_汇总_" & Month(DTP(0).Value) & "." & Day(DTP(0).Value) & "_" & Month(DTP(1).Value) & "." & Day(DTP(1).Value)
        
            
    End Select
    End With
   
    Select Case UCase(cbExcel.text)
    Case ".XLSX", ".XLS"


    Set xlsApp = CreateObject("Excel.Application")
    Set xlsBook = xlsApp.Workbooks.Add
    Set xlsSheet = xlsBook.Worksheets(1)

    With xlsApp
        .Rows(1).Font.Bold = True

    End With
   
    With fps

        For i = 0 To .MaxRows
            For j = 1 To .MaxCols
                .Col = j
                .Row = i
                
                If .Col = 9 Then
                    If cbTpe.text = "BUMPING入库" Then
                    
                        xlsSheet.Cells(i + 1, j) = Trim$(("'" & .text))
                    Else
                        
                        xlsSheet.Cells(i + 1, j) = Trim$(("" & .text))
                    End If
                
                Else
                    xlsSheet.Cells(i + 1, j) = Trim$(("" & .text))

                End If
                If cbTpe.text = "GC Shippinglist" Then
                    If InStr(outputpath, "非保税") > 0 And i > 0 Then '非保税不需要重量
                        If j = E_GC_Shipping.E_NetWeight Or j = E_GC_Shipping.E_TotalWeight Then
                            xlsSheet.Cells(i + 1, j) = ""
                        End If
                    End If
                    If j > E_GC_Shipping.E_BOX Then ''后面的仅辅助，不输出
                         xlsSheet.Cells(i + 1, j) = ""
                    End If
                End If
    
                If wflag = "1" And .Col = 6 And i > 0 Then
                    xlsSheet.Cells(i + 1, j) = wafer_to_string(Trim$(("" & .text)))
                End If
                 If (UCase(Trim(txtCusCode.text)) = "68" Or UCase(Trim(txtCusCode.text)) = "HK006") And .Col = 6 And i > 0 Then
                    If InStr(.text, ",") > 0 Then
                        xlsSheet.Cells(i + 1, j) = wafer_to_string(Trim$(("" & .text)))
                    Else
                        xlsSheet.Cells(i + 1, j) = "#" & (Trim$(("" & .text)))
                    End If
                End If
                If (UCase(Trim(txtCusCode.text)) = "TJ003" Or UCase(Trim(txtCusCode.text)) = "SX") And .Col = 12 And i > 0 Then
                    xlsSheet.Cells(i + 1, j) = Trim$(("'" & .text))
                End If
            Next j
       
        Next i

    End With

    xlsApp.Visible = True
    

    strFileName = outputpath & strPartName & Format(Now, "YYYYMMDD") & GetGC_FileNoNew(UCase(Trim(txtCusCode.text))) & cbExcel.text
    
    xlsBook.SaveAs strFileName

    Set xlsApp = Nothing
    Set xlsSheet = Nothing
    Set xlsBook = Nothing
    
    Case ".CSV"

    Set objStream = CreateObject("ADODB.Stream")
    objStream.type = 2
    objStream.Charset = "UTF-8"
    objStream.Open
   

    With fps

        For i = 0 To .MaxRows
            str_eachrow = ""
            For j = 1 To .MaxCols
                .Col = j
                .Row = i
            
                If .Col = 9 Then
                    If cbTpe.text = "BUMPING入库" Then
                        str_eachrow = str_eachrow & "," & Trim$(("'" & .text))
                    Else
                        
                        str_eachrow = str_eachrow & "," & Trim$(("" & .text))
                    End If
                
                Else
                    If j = 1 Then
                        str_eachrow = Trim$(("" & .text))
                    Else
                        If cbTpe.text = "GC Shippinglist" Then
                            If j <= E_GC_Shipping.E_BOX Then '后面的仅辅助，不输出
                                If InStr(outputpath, "非保税") > 0 And i > 0 Then '非保税不需要重量
                                    If j = E_GC_Shipping.E_NetWeight Or j = E_GC_Shipping.E_TotalWeight Then
                                     
                                      str_eachrow = str_eachrow & ","
                                    Else
                                        str_eachrow = str_eachrow & "," & Trim$(("" & .text))
                                    End If
                                Else
                                    str_eachrow = str_eachrow & "," & Trim$(("" & .text))
                                End If
                            End If
                        ElseIf cbTpe.text = "GC list汇总" Then 'wafer_id,比如01,02
                            If j = 11 Then
                                str_eachrow = str_eachrow & ",'" & Trim$(("" & .text))
                             
                            Else
                                str_eachrow = str_eachrow & "," & Trim$(("" & .text))
                            End If
                        Else
                            str_eachrow = str_eachrow & "," & Trim$(("" & .text))
                        End If
                    End If

                End If

    
            Next j
            
            objStream.WriteText str_eachrow & vbCrLf
        Next i

    End With
    strFileName = outputpath & strPartName & Format(Now, "YYYYMMDD") & GetGC_FileNoNew(UCase(Trim(txtCusCode.text))) & cbExcel.text
    objStream.SaveToFile strFileName, 2
    Set objStream = Nothing
    
    '打开csv文件
    Set xlsApp1 = CreateObject("excel.application")
    xlsApp1.Workbooks.Open strFileName
    xlsApp1.Visible = True
    
    Set xlsApp1 = Nothing
    
    End Select

        AddSql2 ("insert into [erpdata].[dbo].[GR_GC_SendHistory](单据编号,SendTime,Flag,createdby,createdDate,customername) values ('" + UCase(Trim(txtShipNo.text)) + "','" & Format(Now, "YYYY-MM-DD") & "','Y','Auto',getdate(),'" & UCase(Trim(txtCusCode.text)) & "') ")

    Exit Sub

Ert:

    If Not (xlsApp Is Nothing) Then
        
        Set xlsApp = Nothing
        Set xlsSheet = Nothing
        Set xlsBook = Nothing

    End If
    
    If Not (xlsApp1 Is Nothing) Then
        
        Set xlsApp1 = Nothing

    End If
    If Not (objStream Is Nothing) Then
        Set objStream = Nothing
    End If

End Sub

Private Sub txtCusCode_Change()

Select Case txtCusCode.text

    Case "GC"
        cbExcel.ListIndex = 1
    Case "SX", "TJ003", "SC081"
        cbExcel.ListIndex = 0
     Case "68", "HK006"
        cbExcel.ListIndex = 2
End Select

End Sub


Private Sub txtShipNo_Change()
Cbshipto.text = ""
End Sub

Function GetGcVer_Type(strShipNO As String)

    Dim gcrev_w As String
    Dim gcrev_n As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    GetGcVer_Type = ""
    gcrev_wlt = ""
    gcrev_normal = ""
    If rs.State = adStateOpen Then rs.Close
    rs.Source = "select DISTINCT SUBSTRING(isnull(a.gcversion,''),3,1) as 二级代码第三位 ,ISNULL(b.制程,'' ) AS 制程 FROM erptemp..tblshipreport_new a " & _
    " LEFT JOIN erptemp..GcCode_Reference b ON a.remark3 =b.成品料号 AND( SUBSTRING(isnull(a.gcversion,''),3,1)=b.二级代码 OR SUBSTRING(isnull(a.gcversion,''),3,1)=b.分bin二级代码 )" & _
    " WHERE  a.ship_order=  '" & strShipNO & "'"
    rs.Open , INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            If Trim(rs("二级代码第三位")) = "" Then
                MsgBox "二级代码没有维护", vbInformation, 提示
                GetGcVer_Type = ""
                Exit Function
            End If
            If Trim(rs("制程")) = "" Then
                MsgBox "二级代码对应形式维护错误，请确认！", vbInformation, 提示
                GetGcVer_Type = ""
                Exit Function
            End If
         
            If UCase(Trim(rs("制程"))) = "NORMAL" Or UCase(Trim(rs("制程"))) = "转NORMAL" Then
                gcrev_n = Trim(rs("二级代码第三位"))
            ElseIf UCase(Trim(rs("制程"))) = "WLT" Then
                gcrev_w = Trim(rs("二级代码第三位"))
            Else
                MsgBox "二级代码对应形式维护错误，请确认！", vbInformation, 提示
            End If
            rs.MoveNext
        Next
    Else
        MsgBox "系统数据缺失，请反馈！", vbInformation, 提示
    End If
    GetGcVer_Type = gcrev_n & "," & gcrev_w
    rs.Close
    Set rs = Nothing


End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       wafer_to_string
' Description:       很无聊的方法 把wafer片数1,2,3,4，8,10,11这种变成 #1-4,8,#10-11
' Created by :       Project Administrator
' Machine    :       DESKTOP-F6L8S2V
' Date-Time  :       2020/1/17-14:40:51
'
' Parameters :       waferlist (String)
'--------------------------------------------------------------------------------
Private Function wafer_to_string(WAFERLIST As String) As String
Dim TEMP As String
Dim String2 As String
Dim bb() As String
Dim b() As String
Dim i As Integer
Dim j As Integer
b = Split(WAFERLIST, ",")

Last = UBound(b) - LBound(b) + 1  '获取数组大小

If Last = 1 Then
    wafer_to_string = b(0)
    Exit Function
ElseIf Last = 2 Then
    wafer_to_string = b(0) + "," + b(1)
End If

Last = Last - 2
String2 = "#" + b(0)
TEMP = b(0)
For i = 0 To Last
    j = i + 1
    If (b(j) - b(i)) > 1 Then
        If b(i) <> TEMP Then
            String2 = String2 + "-" + b(i) + ",#" + b(j)
        Else
            bb = Split(String2, b(j))
           ' String2 = Mid(bb(0), 1, Len(bb(0)) - 4) + "," + TEMP + ",#" + b(j)
            String2 = String2 + ",#" + b(j)
        End If
        TEMP = b(j)
    End If
Next i
    Last = Last + 1
    String2 = String2 + "-" + b(Last)
    wafer_to_string = String2
End Function






