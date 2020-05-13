VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmHK037ShipCheck 
   Caption         =   "艾为出货标签和出货资料比对"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13710
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
   ScaleHeight     =   9045
   ScaleWidth      =   13710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.TextBox txtCheckResult 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2213
         Width           =   4215
      End
      Begin VB.TextBox txtQrCode 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   1613
         Width           =   4215
      End
      Begin VB.CommandButton cmdBegin 
         Caption         =   "开始比对"
         Height          =   360
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   990
      End
      Begin VB.TextBox txtQBoxID 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1020
         Width           =   4215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   990
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   5535
         Index           =   0
         Left            =   600
         TabIndex        =   10
         Top             =   2880
         Width           =   11055
         _Version        =   524288
         _ExtentX        =   19500
         _ExtentY        =   9763
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
         MaxCols         =   2
         MaxRows         =   0
         SpreadDesigner  =   "FrmHK037ShipCheck.frx":0000
         Appearance      =   1
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "比对结果"
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
         Left            =   840
         TabIndex        =   8
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "二维码箱号"
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
         Left            =   600
         TabIndex        =   6
         Top             =   1680
         Width           =   1140
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   10320
         TabIndex        =   5
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   1085
         _cy             =   873
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Q开头大箱号"
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
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   1260
      End
   End
End
Attribute VB_Name = "FrmHK037ShipCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBegin_Click()
txtQBoxID.text = ""
txtQrCode.text = ""

If txtQBoxID.text = "" Then
    Call PlaySound("请扫描Q箱号")
    fpS(0).MaxRows = 0
End If

txtQBoxID.Enabled = True
txtQBoxID.SetFocus

End Sub

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub Form_Load()
With fpS(0)
    .Col = -1
    .Row = -1
    .Lock = True
    .Col = 1
    .Row = 0
    .FontSize = 10
    .Col = 2
    .Row = 0
    .FontSize = 10

    .SetText 1, 0, "出货资料"
    .SetText 2, 0, "出货外箱标签"
 
    .ColWidth(1) = 31
    .ColWidth(2) = 31
End With
End Sub

Private Sub txtQBoxID_KeyPress(KeyAscii As Integer)
Dim strQrCode As String
Dim strPN     As String
Dim strLblPN  As String
Dim strLC     As String
Dim strLblLC  As String
Dim strDC     As String
Dim strLblDC  As String
Dim strPKG    As String
Dim strLblPKG As String
Dim strqty    As String
Dim strLblQty As String
Dim strsql    As String
Dim strArr()  As String
Dim strQBoxID As String
Dim rs        As New ADODB.Recordset

If KeyAscii <> vbKeyReturn Or Len(Trim(txtQBoxID.text)) = 0 Then Exit Sub
strQBoxID = UCase(Trim$(txtQBoxID.text))

If Left$(strQBoxID, 1) <> "Q" Then
    MsgBox "标签扫描出错", vbCritical, "警告"
    Exit Sub

End If


'strSql = " select b.箱号,d.MPN_DESC as 产品名称,isnull(d.probe_ship_part_type, d.ZX_INVOICE) as 封装批次号,substring(convert(varchar(100), datepart(YY, f.ERPCREATEDATE)),3,2) + " & " right('0' + convert(varchar(100), datepart(WW, f.ERPCREATEDATE)), 2) as 日期,SUM(b.数量) AS 数量,isnull(d.reticle_level_72, d.comp_code) as 封装形式 " & " From erpdata .. tblStockSQfh a,erpdata .. tblStocksqfhsub b,ERPBASE .. tblmappingData c,ERPBASE .. tblCustomerOI d,erpdata .. tblTSVwaferlist e,erpdata .. tblTSVworkorder f " & " Where b.单据编号 = a.单据编号 And b.单据项次 = a.序号 And c.SUBSTRATEID = b.流程卡编号 And D.ID = c.filename And e.WAFERID = c.SUBSTRATEID And f.ORDERNAME = e.ORDERNAME " & " and b.箱号 = '" & strQBoxID & "' GROUP BY b.箱号, d.MPN_DESC,substring(convert(varchar(100), datepart(YY, f.ERPCREATEDATE)),3,2) +right('0' + convert(varchar(100), datepart(WW, f.ERPCREATEDATE)), 2), " & " D.probe_ship_part_type , D.reticle_level_72, D.ZX_INVOICE, D.comp_code "

 strsql = " select b.箱号,d.MPN_DESC as 产品名称,isnull(d.probe_ship_part_type, d.ZX_INVOICE) as 封装批次号 ,substring(convert(varchar(100), datepart(YY, f.ERPCREATEDATE)),3,2) + " & _
          " RIGHT('0' + convert(varchar(100), datepart(WW, f.ERPCREATEDATE)), 2) as 日期 ,SUM(b.数量) AS 数量,g.package as 封装形式  From erpdata .. tblStockSQfh a,erpdata .. tblStocksqfhsub b " & _
          " ,ERPBASE .. tblmappingData c,ERPBASE .. tblCustomerOI d,erpdata .. tblTSVwaferlist e,erpdata .. tblTSVworkorder f ,erptemp .. EU010_reference g " & _
          " Where b.单据编号 = a.单据编号 And b.单据项次 = a.序号 And c.SUBSTRATEID = b.流程卡编号 And D.ID = c.filename And e.waferid = c.SUBSTRATEID And f.ORDERNAME = e.ORDERNAME " & _
          " AND b.箱号 = '" & strQBoxID & "' AND g.cust_device = d.MPN_DESC  GROUP BY b.箱号, d.MPN_DESC,substring(convert(varchar(100)  , datepart(YY, f.ERPCREATEDATE)),3,2) " & _
          " +right('0' + convert(varchar(100), datepart(WW, f.ERPCREATEDATE)), 2),  D.probe_ship_part_type , D.ZX_INVOICE , g.package "

Set rs = Get_SqlserveRs(strsql)
If rs.EOF Then
    MsgBox "查询不到该箱号", vbCritical, "警告"
    Exit Sub

End If

strPN = rs!产品名称
strLC = rs!封装批次号
strDC = rs!日期
strPKG = rs!封装形式
strqty = rs!数量

With fpS(0)
    .MaxRows = 5
    .SetText 1, 1, UCase(strPN)
    .SetText 1, 2, UCase(strLC)
    .SetText 1, 3, UCase(strDC)
    .SetText 1, 4, UCase(strPKG)
    .SetText 1, 5, strqty

End With

Call PlaySound("Q箱号已扫描")
txtQrCode.SetFocus

End Sub

Private Sub txtQrCode_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Or Len(Trim(txtQrCode.text)) = 0 Then Exit Sub
If Not CheckHandler Then Exit Sub
txtQBoxID.text = ""
txtQrCode.text = ""
txtQBoxID.SetFocus

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckHandler
' Description:       比对
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/10/15-15:34:56
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function CheckHandler() As Boolean
Dim strQrCode As String
Dim strPN     As String
Dim strLblPN  As String
Dim strLC     As String
Dim strLblLC  As String
Dim strDC     As String
Dim strLblDC  As String
Dim strPKG    As String
Dim strLblPKG As String
Dim strqty    As String
Dim strLblQty As String
Dim strsql    As String
Dim strArr()  As String
Dim rs        As New ADODB.Recordset

CheckHandler = False
strQrCode = UCase(Trim$(txtQrCode.text))
If Left(strQrCode, 2) <> "O;" Then
    MsgBox "标签扫描出错", vbCritical, "警告"
    Exit Function

End If

strArr = Split(strQrCode, ";")
strLblPN = strArr(1)
strLblLC = strArr(2)
strLblDC = strArr(3)
strLblPKG = strArr(6)
strLblQty = strArr(4)

With fpS(0)
    .SetText 2, 1, strLblPN
    .SetText 2, 2, strLblLC
    .SetText 2, 3, strLblDC
    .SetText 2, 4, strLblPKG
    .SetText 2, 5, strLblQty
    .Col = 1
    .Row = 1
    strPN = .text
    .Row = 2
    strLC = .text
    .Row = 3
    strDC = .text
    .Row = 4
    strPKG = .text
    .Row = 5
    strqty = .text

End With

If InStr(UCase(strPN), UCase(strLblPN)) <= 0 Then
    fpS(0).Row = 1
    fpS(0).BackColor = vbRed
    MsgBox "出货资料和出货标签的产品名称不一致", vbCritical, "警告"
    Exit Function

End If

If UCase(strLC) <> UCase(strLblLC) Then
    fpS(0).Row = 2
    fpS(0).BackColor = vbRed
    MsgBox "出货资料和出货标签的封装批次号不一致", vbCritical, "警告"
    Exit Function

End If

If UCase(strDC) <> UCase(strLblDC) Then
    fpS(0).Row = 3
    fpS(0).BackColor = vbRed
    MsgBox "出货资料和出货标签的D/C不一致", vbCritical, "警告"
    Exit Function

End If

If UCase(strPKG) <> UCase(strLblPKG) Then
    fpS(0).Row = 4
    fpS(0).BackColor = vbRed
    MsgBox "出货资料和出货标签的封装形式不一致", vbCritical, "警告"
    Exit Function

End If

If strqty <> strLblQty Then
    fpS(0).Row = 5
    fpS(0).BackColor = vbRed
    MsgBox "出货资料和出货标签的数量不一致", vbCritical, "警告"
    Exit Function

End If

Call PlaySound("比对通过")
fpS(0).MaxRows = 0
CheckHandler = True

End Function

Private Sub PlaySound(sFileName As String)
Dim sPath   As String
Dim sSuffix As String

sPath = "\\10.160.1.84\public\media_source\"
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix

End Sub
