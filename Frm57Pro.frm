VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.ocx"
Begin VB.Form Frm57Pro 
   Caption         =   "57出矽力杰"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13260
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
   ScaleHeight     =   10575
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   18653
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   8421504
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "内包转换"
      TabPicture(0)   =   "Frm57Pro.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "player1"
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(5)=   "txtTempQRCode"
      Tab(0).Control(6)=   "txtSuccess"
      Tab(0).Control(7)=   "txtFailed"
      Tab(0).Control(8)=   "txtPrintQty"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "外包合箱"
      TabPicture(1)   =   "Frm57Pro.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label11"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label12"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "fps"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtQRCode"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "btnBegin"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "btnClose"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtPrintQty2"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtShipDate"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtDieQty"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtTrayQty"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtOuterPkgNO"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtOuterPkg"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "btnPrint"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).ControlCount=   19
      Begin VB.CommandButton btnPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "补打(&P)"
         Height          =   360
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox txtOuterPkg 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   10920
         TabIndex        =   27
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox txtOuterPkgNO 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtTrayQty 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1650
         Width           =   1215
      End
      Begin VB.TextBox txtDieQty 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2010
         Width           =   1215
      End
      Begin VB.TextBox txtShipDate 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2370
         Width           =   1215
      End
      Begin VB.TextBox txtPrintQty2 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   10920
         TabIndex        =   17
         Text            =   "1"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtPrintQty 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   -66240
         TabIndex        =   15
         Text            =   "3"
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton btnClose 
         BackColor       =   &H00FFC0C0&
         Caption         =   "合 箱(&C)"
         Height          =   360
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   562
         Width           =   975
      End
      Begin VB.CommandButton btnBegin 
         BackColor       =   &H00C0C0C0&
         Caption         =   "开 始(&B)"
         Height          =   360
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   562
         Width           =   975
      End
      Begin VB.TextBox txtQRCode 
         BackColor       =   &H00FFC0FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   600
         Width           =   5175
      End
      Begin VB.TextBox txtFailed 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   7815
         Left            =   -70200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtSuccess 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   7815
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtTempQRCode 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   -73080
         TabIndex        =   1
         Top             =   600
         Width           =   5175
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   7575
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   8655
         _Version        =   524288
         _ExtentX        =   15266
         _ExtentY        =   13361
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
         MaxCols         =   8
         MaxRows         =   0
         SpreadDesigner  =   "Frm57Pro.frx":0038
         Appearance      =   1
         AppearanceStyle =   0
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱标签补打:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9720
         TabIndex        =   26
         Top             =   5280
         Width           =   1140
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前外箱虚拟箱号:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9360
         TabIndex        =   24
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前外箱卷盘数量:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9360
         TabIndex        =   23
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前外箱总数量:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9480
         TabIndex        =   22
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出货日期DC:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9840
         TabIndex        =   21
         Top             =   2400
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打印份数:"
         Height          =   195
         Left            =   10080
         TabIndex        =   16
         Top             =   1005
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打印份数:"
         Height          =   195
         Left            =   -67080
         TabIndex        =   14
         Top             =   645
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前外箱已扫描卷盘明细:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   2040
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘临时标签二维码:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1680
      End
      Begin WMPLibCtl.WindowsMediaPlayer player1 
         Height          =   495
         Left            =   -66600
         TabIndex        =   7
         Top             =   1560
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转换失败:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -70200
         TabIndex        =   6
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转换成功:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74760
         TabIndex        =   4
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘临时标签二维码:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74760
         TabIndex        =   2
         Top             =   630
         Width           =   1680
      End
   End
End
Attribute VB_Name = "Frm57Pro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strTmpQRCodeList As String
Private strFormalQRCodeList As String
Private Type FORMAL_INNERBOX_LABEL

po_no As String
part_no As String
Quantity As String
DATE_CODE As String
LOT_NO As String
QR_CODE As String
SERIAL_NO As String
End Type

Private Type FORMAL_OUTERBOX_LABEL

po_no As String
part_no As String
Quantity As String
DATE_CODE As String
LOT_NO As String
QR_CODE As String
SHIP_DATE As String
End Type

Private Sub btnBegin_Click()
Call InitData
txtQRCode.Enabled = True
txtQRCode.SetFocus
txtShipDate.text = Right(Year(Now), 2) & Right$("00" & Month(Now), 2) & Right$("00" & Day(Now), 2)
txtOuterPkgNO.text = Get_OracleStr("select trglabelseq.QTSeq_NotMesQbox(SEQ_57SHIPDN.NEXTVAL)  from dual")

Call PlaySound("请依次扫描卷盘临时标签二维码")
End Sub

Private Sub InitData()
fps.MaxRows = 0
strTmpQRCodeList = ""
strFormalQRCodeList = ""
txtTrayQty.text = "0"
txtDieQty.text = "0"
txtShipDate.text = ""
txtOuterPkgNO.text = ""

End Sub

Private Sub btnClose_Click()

If fps.MaxRows = 0 Then
    MsgBox "当前外箱未录入铝箔袋信息,不可合箱", vbCritical, "警告"
    Exit Sub
End If

Dim rs As New ADODB.Recordset
Dim strSql As String
Dim outerLabel As FORMAL_OUTERBOX_LABEL

strSql = "select po,pn,dc,lotid,sum(qty) as qty from TBL_57XLJ where outerbox_no = '" & txtOuterPkgNO & "' group by po,pn,dc,lotid "
Set rs = Get_OracleRs(strSql)
Do While Not rs.EOF
    outerLabel.po_no = rs!PO
    outerLabel.part_no = rs!PN
    outerLabel.DATE_CODE = rs!DC
    outerLabel.LOT_NO = rs!LOTID
    outerLabel.Quantity = rs!QTY
    outerLabel.SHIP_DATE = txtShipDate.text
    outerLabel.QR_CODE = outerLabel.po_no & "!" & outerLabel.part_no & "!" & outerLabel.Quantity & "!" & outerLabel.DATE_CODE & "!" & outerLabel.LOT_NO
    
    Call PrintOuterLabel(outerLabel)
    Call PrintOuterLabel2(outerLabel)
    rs.MoveNext
Loop

Call TransToERP
Call InitData
MsgBox "外箱标签已经打印完成", vbInformation, "提示"
End Sub

Private Sub PrintOuterLabel2(outerLabel As FORMAL_OUTERBOX_LABEL)
Dim strContent As String
Dim strSql As String

strContent = strContent & Chr(34) & "CARTON_ID" & Chr(34) & "," & Chr(34) & txtOuterPkgNO.text & Chr(34) & ";"
strContent = strContent & Chr(34) & "CUSTOMER_LOT" & Chr(34) & "," & Chr(34) & outerLabel.LOT_NO & Chr(34) & ";"
strContent = strContent & Chr(34) & "CUSTOMER_LOT_SLASH" & Chr(34) & "," & Chr(34) & outerLabel.LOT_NO & Chr(34) & ";"
strContent = strContent & Chr(34) & "OUT_GOOD_DIE" & Chr(34) & "," & Chr(34) & outerLabel.Quantity & Chr(34)

strSql = " insert into erpdata.dbo.tblME_PrintInfo(PrinterNameID,BartenderName,Content,Flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & _
" values('ALL_OUT_2B1F_1','57OUT1.btw','" & strContent & "','0',GetDate(),'PKG','" & txtOuterPkgNO.text & "','57OUT1',2)"

AddSql2 (strSql)

End Sub

Private Sub TransToERP()
Dim strSql As String
Dim strCartonID As String
Dim lCartonQty As Long
Dim id As Long

strCartonID = txtOuterPkgNO.text
lCartonQty = CLng(txtDieQty.text)

strSql = "insert into [erpdata].[dbo].[tblPackMainInf](箱号,客户代码,数量,合格标记,装箱标记,产线标记) values('" & strCartonID & "','57'," & lCartonQty & ",'0','1','1')"
AddSql2 (strSql)

'2 insert - update [erpdata].[dbo].[tblPackTreeInf]
strSql = "insert into [erpdata].[dbo].[tblPackTreeInf](箱号,上级序号,基层标记 ,Memo) values('" & strCartonID & "',0,1,'57')"
AddSql2 (strSql)

id = Get_SqlserverNo("select 序号 as ID from [erpdata].[dbo].[tblPackTreeInf] a where a.箱号='" & strCartonID & "' and Memo='57' ")
strSql = "Update [erpdata].[dbo].[tblPackTreeInf] set 上级序号='" & id & "',Memo='57' where 箱号 in (select * from  OPENQUERY(ORACLEDB, 'SELECT tray_id from TBL_57XLJ where outerbox_no = ''" & strCartonID & "'' ' ) ) "
AddSql2 (strSql)

'3 insert - update [erpdata].[dbo].[tblStockNumTree]
'strSql = "insert into [erpdata].[dbo].[tblStockNumTree](序号,箱号,上级序号,基层标记 ,尺寸,重量,Memo,DN) values( " & id & ",'" & strCartonID & "',0,1,'','','57','')"
'AddSql2 (strSql)
'
'strSql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & id & "',Memo='57' where 箱号 in (select * from  OPENQUERY(ORACLEDB, 'SELECT tray_id from TBL_57XLJ where outerbox_no = ''" & strCartonID & "'' ' ) ) "
'AddSql2 (strSql)

End Sub

Private Sub PrintOuterLabel(outerLabel As FORMAL_OUTERBOX_LABEL)
Dim strContent As String
Dim strSql As String

strContent = strContent & Chr(34) & "57XLJ_DC" & Chr(34) & "," & Chr(34) & outerLabel.DATE_CODE & Chr(34) & ";"
strContent = strContent & Chr(34) & "57XLJ_PN" & Chr(34) & "," & Chr(34) & outerLabel.part_no & Chr(34) & ";"
strContent = strContent & Chr(34) & "57XLJ_QR_CODE" & Chr(34) & "," & Chr(34) & outerLabel.QR_CODE & Chr(34) & ";"
strContent = strContent & Chr(34) & "CUSTOMER_LOT" & Chr(34) & "," & Chr(34) & outerLabel.LOT_NO & Chr(34) & ";"
strContent = strContent & Chr(34) & "IN_GOOD_DIE" & Chr(34) & "," & Chr(34) & outerLabel.Quantity & Chr(34) & ";"
strContent = strContent & Chr(34) & "PACKING_DATE_10" & Chr(34) & "," & Chr(34) & outerLabel.SHIP_DATE & Chr(34) & ";"
strContent = strContent & Chr(34) & "PO_NO" & Chr(34) & "," & Chr(34) & outerLabel.po_no & Chr(34)

strSql = " insert into erpdata.dbo.tblME_PrintInfo(PrinterNameID,BartenderName,Content,Flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & _
" values('ALL_OUT_2B1F_2','57XLJ-OUT1.btw','" & strContent & "','0',GetDate(),'PKG','" & txtOuterPkgNO.text & "','57XLJ-OUT1'," & txtPrintQty2.text & ")"

AddSql2 (strSql)

End Sub

Private Sub btnPrint_Click()
Dim strEventID As String
Dim strSql As String

If txtOuterPkg.text = "" Then
    MsgBox "请输入要补打的外箱箱号(Q*********)", vbCritical, "警告"
    Exit Sub
End If

strEventID = UCase(Trim$(txtOuterPkg.text))

strSql = "insert into erpdata.dbo.tblME_PrintInfo(PrinterNameID,BartenderName,Flag,Createdate,EVENT_ID,PRINT_QTY,Content) " & " select top 1 PrinterNameID,BartenderName,'0',GETDATE(),EVENT_ID,1,Content  from erpdata.dbo.tblME_PrintInfo where EVENT_ID = '" & strEventID & "' and LABEL_ID = '57XLJ-OUT1' order by ID desc"
AddSql2 (strSql)

strSql = "insert into erpdata.dbo.tblME_PrintInfo(PrinterNameID,BartenderName,Flag,Createdate,EVENT_ID,PRINT_QTY,Content) " & " select top 1 PrinterNameID,BartenderName,'0',GETDATE(),EVENT_ID,1,Content  from erpdata.dbo.tblME_PrintInfo where EVENT_ID = '" & strEventID & "' and LABEL_ID = '57OUT1' order by ID desc"
AddSql2 (strSql)

MsgBox "补打成功", vbInformation, "提示"

End Sub

Private Sub Form_Activate()
txtTempQRCode.SetFocus
End Sub

Private Sub Form_Load()
With fps
    .MaxRows = 0
    .MaxCols = 5
    .Col = -1
    .Row = -1
    .Lock = True
    
    .SetText 1, 0, "W/O#"
    .SetText 2, 0, "P/N"
    .SetText 3, 0, "QTY"
    .SetText 4, 0, "D/C"
    .SetText 5, 0, "LOT NO"
    .SetText 6, 0, "SERIAL NO"

    .ColWidth(1) = 15
    .ColWidth(2) = 15
    .ColWidth(3) = 15
    .ColWidth(6) = 15

End With

End Sub

Private Sub txtQrCode_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then Exit Sub
If Len(Trim$(txtQRCode.text)) = 0 Then Exit Sub

Dim strCode As String
strCode = UCase(Trim$(txtQRCode.text))
txtQRCode.text = ""

If InStr(strCode, "@$") > 0 Then
    Call GetOuterLabelByTempCode(strCode)
'ElseIf InStr(strCode, "!") > 0 Then
'    Call GetOuterLabelByFormalCode(strCode)
Else
    MsgBox "请扫描正确的卷盘临时标签二维码", vbCritical, "扫描错误"
End If

End Sub

'扫描临时标签
Private Function GetOuterLabelByTempCode(strCode As String) As Boolean
GetOuterLabelByTempCode = False
If InStr(strTmpQRCodeList, strCode) > 0 Then
    MsgBox "该标签已经扫描过,请不要重复扫描", vbCritical, "警告"
    Exit Function
End If

Dim trayLabel As FORMAL_INNERBOX_LABEL
Dim strArray() As String

strArray = Split(strCode, "@$")

trayLabel.po_no = strArray(3)
trayLabel.part_no = strArray(1)
trayLabel.Quantity = strArray(7)
trayLabel.DATE_CODE = strArray(4)
trayLabel.LOT_NO = strArray(5)
trayLabel.SERIAL_NO = strArray(8)

If Not CheckOuterCode(trayLabel) Then Exit Function
strTmpQRCodeList = strTmpQRCodeList & strCode & "@@"
Call ListLabelHistory(trayLabel)
Call SaveToDB(trayLabel)
Call PlaySound("铝箔袋已扫描")
GetOuterLabelByTempCode = True
End Function

Private Sub SaveToDB(trayLabel As FORMAL_INNERBOX_LABEL)
Dim strSql As String

strSql = "insert into TBL_57XLJ(OUTERBOX_NO,TRAY_ID,QTY,LOTID,PO,DC,PN,CREATE_DATE) values('" & txtOuterPkgNO.text & "','" & trayLabel.SERIAL_NO & "'," & trayLabel.Quantity & ",'" & trayLabel.LOT_NO & "','" & trayLabel.po_no & "','" & trayLabel.DATE_CODE & "','" & trayLabel.part_no & "',sysdate)"
AddSql (strSql)

End Sub

'扫描正式标签
Private Function GetOuterLabelByFormalCode(strCode As String)
GetOuterLabelByFormalCode = False

Dim trayLabel As FORMAL_INNERBOX_LABEL
Dim strArray() As String

strArray = Split(strCode, "!")

trayLabel.po_no = strArray(0)
trayLabel.part_no = strArray(1)
trayLabel.Quantity = strArray(2)
trayLabel.DATE_CODE = strArray(3)
trayLabel.LOT_NO = strArray(4)

If Not CheckOuterCode(trayLabel) Then Exit Function
strFormalQRCodeList = strFormalQRCodeList & strCode & "@@"
Call ListLabelHistory(trayLabel)
Call PlaySound("铝箔袋已扫描")
GetOuterLabelByFormalCode = True
End Function

Private Function CheckOuterCode(trayLabel As FORMAL_INNERBOX_LABEL) As Boolean
CheckOuterCode = False
Dim i As Integer

'With fps
'    For i = 1 To .MaxRows
'        .Row = i
'        '机种
'        .Col = 2
'        If .text <> trayLabel.PART_NO Then
'            MsgBox "每个外箱只能装相同P/N的产品,装箱错误", vbCritical, "警告"
'            Exit Function
'        End If
'        'DateCode
''        .Col = 4
''        If .text <> trayLabel.DATE_CODE Then
''            MsgBox "每个外箱只能装相同D/C的产品,装箱错误", vbCritical, "警告"
''            Exit Function
''        End If
'        'LOT
''        .Col = 5
''        If .text <> trayLabel.LOT_NO Then
''
''            If bTwice = True Then
''                MsgBox "每个外箱最多只能装2个不同Lot的产品,扫描错误", vbCritical, "警告"
''                Exit Function
''
''            End If
''
''        End If
'
'    Next i
'
'End With

CheckOuterCode = True
End Function

Private Sub ListLabelHistory(trayLabel As FORMAL_INNERBOX_LABEL)

Dim i As Integer
Dim lQty As Long

With fps
    .MaxRows = .MaxRows + 1
    .SetText 1, .MaxRows, trayLabel.po_no
    .SetText 2, .MaxRows, trayLabel.part_no
    .SetText 3, .MaxRows, trayLabel.Quantity
    .SetText 4, .MaxRows, trayLabel.DATE_CODE
    .SetText 5, .MaxRows, trayLabel.LOT_NO
    .SetText 6, .MaxRows, trayLabel.SERIAL_NO

End With

With fps
    txtTrayQty.text = .MaxRows
    
    For i = 1 To .MaxRows
        .Row = i
        .Col = 3
        lQty = lQty + Val(.text)
    Next
End With

txtDieQty.text = lQty
End Sub

Private Sub txtTempQRCode_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then Exit Sub
If Len(Trim$(txtTempQRCode.text)) = 0 Then Exit Sub

Dim strTempCode As String
strTempCode = UCase$(Trim$(txtTempQRCode.text))
txtTempQRCode.text = ""

If Not CheckTempLabel(strTempCode) Then Exit Sub

If ConverFormalLabel(strTempCode) = True Then
    txtSuccess.text = Split(strTempCode, "@$")(8) + vbCrLf + txtSuccess.text
Else
    txtFailed.text = Split(strTempCode, "@$")(8) + vbCrLf + txtFailed.text
End If
End Sub

Private Function CheckTempLabel(strCode As String) As Boolean
CheckTempLabel = False
If InStr(strCode, "@$") = 0 Then
    MsgBox "请扫描正确的临时标签二维码信息", vbCritical, "二维码格式错误"
    Exit Function
End If

CheckTempLabel = True
End Function

Private Function ConverFormalLabel(strCode As String) As Boolean
ConverFormalLabel = False

Dim trayLabel As FORMAL_INNERBOX_LABEL
Dim strArray() As String

strArray = Split(strCode, "@$")

trayLabel.po_no = strArray(3)
trayLabel.part_no = strArray(1)
trayLabel.Quantity = strArray(7)
trayLabel.DATE_CODE = strArray(4)
trayLabel.LOT_NO = strArray(5)
trayLabel.SERIAL_NO = strArray(8)
trayLabel.QR_CODE = trayLabel.po_no + "!" + trayLabel.part_no + "!" + trayLabel.Quantity + "!" + trayLabel.DATE_CODE + "!" + trayLabel.LOT_NO

Call PrintInnerLabel(trayLabel)
Call PlaySound("标签已转换完成")
ConverFormalLabel = True
End Function

Private Sub PrintInnerLabel(trayLabel As FORMAL_INNERBOX_LABEL)
Dim strContent As String
Dim strSql As String

strContent = strContent & Chr(34) & "57XLJ_DC" & Chr(34) & "," & Chr(34) & trayLabel.DATE_CODE & Chr(34) & ";"
strContent = strContent & Chr(34) & "57XLJ_PN" & Chr(34) & "," & Chr(34) & trayLabel.part_no & Chr(34) & ";"
strContent = strContent & Chr(34) & "57XLJ_QR_CODE" & Chr(34) & "," & Chr(34) & trayLabel.QR_CODE & Chr(34) & ";"
strContent = strContent & Chr(34) & "CUSTOMER_LOT" & Chr(34) & "," & Chr(34) & trayLabel.LOT_NO & Chr(34) & ";"
strContent = strContent & Chr(34) & "IN_GOOD_DIE" & Chr(34) & "," & Chr(34) & trayLabel.Quantity & Chr(34) & ";"
strContent = strContent & Chr(34) & "PO_NO" & Chr(34) & "," & Chr(34) & trayLabel.po_no & Chr(34)

strSql = " insert into erpdata.dbo.tblME_PrintInfo(PrinterNameID,BartenderName,Content,Flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & _
" values('W_IN_2B5F_2','57XLJ-IN1.btw','" & strContent & "','0',GetDate(),'PKG','" & trayLabel.SERIAL_NO & "','57XLJ-IN1'," & txtPrintQty.text & ")"

AddSql2 (strSql)
End Sub

Private Sub PlaySound(strSound As String)
player1.url = "\\10.160.1.84\public\media_source\" & strSound & ".wav"

End Sub
