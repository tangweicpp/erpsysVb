VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmOuterPkgLblSys 
   Caption         =   "WLP外包出货标签打印系统(通用版1.0)"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16320
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
   ScaleHeight     =   8220
   ScaleWidth      =   16320
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frm 
      Caption         =   "出货卷盘(铝箔袋)明细"
      ForeColor       =   &H00FF0000&
      Height          =   12375
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   16935
      Begin VB.Frame Frame1 
         Caption         =   "外箱标签预览"
         ForeColor       =   &H000000FF&
         Height          =   2415
         Left            =   6360
         TabIndex        =   19
         Top             =   5160
         Visible         =   0   'False
         Width           =   9255
         Begin VB.TextBox txtPartNo 
            BackColor       =   &H00FFC0FF&
            Height          =   390
            Left            =   1680
            TabIndex        =   25
            Top             =   375
            Width           =   2775
         End
         Begin VB.TextBox txtLotNo 
            BackColor       =   &H00FFC0FF&
            Height          =   390
            Left            =   1680
            TabIndex        =   24
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox txtQty 
            BackColor       =   &H00FFC0FF&
            Height          =   390
            Left            =   1680
            TabIndex        =   23
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox txtDateCode 
            BackColor       =   &H00FFC0FF&
            Height          =   390
            Left            =   1680
            TabIndex        =   22
            Top             =   1830
            Width           =   2775
         End
         Begin VB.TextBox txtSN 
            BackColor       =   &H00FFC0FF&
            Height          =   390
            Left            =   6000
            TabIndex        =   21
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox txtSealDate 
            BackColor       =   &H00FFC0FF&
            Height          =   390
            Left            =   6000
            TabIndex        =   20
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part Number:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lot Number:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   360
            TabIndex        =   30
            Top             =   960
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   600
            TabIndex        =   29
            Top             =   1440
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Code:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   480
            TabIndex        =   28
            Top             =   1920
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SN:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   5640
            TabIndex        =   27
            Top             =   1410
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seal Date:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   4800
            TabIndex        =   26
            Top             =   1920
            Width           =   1440
         End
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   11775
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   10215
         _Version        =   524288
         _ExtentX        =   18018
         _ExtentY        =   20770
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
         SpreadDesigner  =   "FrmOuterPkgLblSys.frx":0000
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   2655
         Index           =   1
         Left            =   10560
         TabIndex        =   11
         Top             =   240
         Width           =   6135
         _Version        =   524288
         _ExtentX        =   10821
         _ExtentY        =   4683
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
         SpreadDesigner  =   "FrmOuterPkgLblSys.frx":03F8
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Frm 
      Caption         =   "扫描"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16935
      Begin VB.TextBox txtCID 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   8040
         TabIndex        =   18
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton cmdAgain 
         BackColor       =   &H0000FFFF&
         Caption         =   "补打外箱(C)"
         Height          =   360
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtDN 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   683
         Width           =   2175
      End
      Begin VB.TextBox txtMediaDir 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   13
         Text            =   "C:\media_source\"
         Top             =   263
         Width           =   1695
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "退出界面"
         Height          =   360
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Width           =   990
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0C0C0&
         Caption         =   "结束扫描"
         Height          =   360
         Left            =   3390
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1440
         Width           =   990
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0000FFFF&
         Caption         =   "合箱>>>"
         Height          =   360
         Left            =   1770
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00C0C0FF&
         Caption         =   "开始扫描"
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   990
      End
      Begin VB.TextBox txtScan 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   1065
         Width           =   5175
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
         ItemData        =   "FrmOuterPkgLblSys.frx":07F0
         Left            =   1080
         List            =   "FrmOuterPkgLblSys.frx":07F7
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblShipID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出货虚拟单号"
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
         Left            =   120
         TabIndex        =   16
         Top             =   705
         Width           =   1350
      End
      Begin VB.Label lblMediaDir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "声音文件目录"
         Height          =   195
         Left            =   6000
         TabIndex        =   14
         Top             =   308
         Width           =   1080
      End
      Begin WMPLibCtl.WindowsMediaPlayer player1 
         Height          =   495
         Left            =   12480
         TabIndex        =   12
         Top             =   1440
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
      Begin VB.Label lblScan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   450
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
         Left            =   120
         TabIndex        =   4
         Top             =   165
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmOuterPkgLblSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum E_INNER_BOX_INFO

    E_PN = 1
    E_LOTNO
    E_qty
    E_DATECODE
    E_SN
    E_SEALDATE
    E_END

End Enum

Private Type T_INNER_BOX_INFO

    T_PN As String
    T_LOTNO As String
    T_QTY As Long
    T_DATECODE As String
    T_SN As String
    T_SEALDATE As String

End Type

Private Enum E_OUTER_BOX_INFO

    E_CID = 1
    E_QID
    E_OUTBOX_QTY
    E_END

End Enum

Private Type T_OUTER_BOX_INFO

    T_PN As String
    T_HT_PN As String
    T_LOTNO As String
    T_QTY As Long
    T_DATECODE As String
    T_SN As String
    T_QID As String
    T_PACKING_DATE1 As String

End Type

Private Sub cmdAgain_Click()
Dim strSql     As String
Dim strEventID As String

If txtCID.text = "" Then
    MsgBox "请扫描或者输入需要补打的外箱-C条码", vbCritical, "警告"
    Exit Sub

End If

strEventID = UCase(Trim$(txtCID.text))
strSql = "insert into erpdata.dbo.tblME_PrintInfo(PrinterNameID,BartenderName,Flag,Createdate,EVENT_ID,PRINT_QTY,Content) " & " select top 1 PrinterNameID,BartenderName,'0',GETDATE(),EVENT_ID,1,Content  from erpdata.dbo.tblME_PrintInfo where EVENT_ID = '" & strEventID & "' order by ID desc"
If AddSql2(strSql) Then
    MsgBox "补打成功", vbInformation, "提示"
Else
    MsgBox "补打失败", vbInformation, "提示"

End If

End Sub

Private Sub cmdClose_Click()
txtScan.Visible = False

End Sub

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub Form_Load()
Call InitData
Call InitCtrls

End Sub

Private Sub InitData()

End Sub

Private Sub InitCtrls()
txtScan.Visible = False
cboCustCode.ListIndex = 0
Call InitFps

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitFps
' Description:       初始化Fps
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/31-11:12:18
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitFps()

With Fps(0)
    .ReDraw = False
    .MaxCols = E_INNER_BOX_INFO.E_END - 1
    .MaxRows = 0
    .FontBold = False
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .SelForeColor = &HFF8080
    .SetText E_INNER_BOX_INFO.E_PN, 0, "Part Number"
    .SetText E_INNER_BOX_INFO.E_LOTNO, 0, "Lot Number"
    .SetText E_INNER_BOX_INFO.E_qty, 0, "Quantity"
    .SetText E_INNER_BOX_INFO.E_DATECODE, 0, "Date Code"
    .SetText E_INNER_BOX_INFO.E_SN, 0, "SN"
    .SetText E_INNER_BOX_INFO.E_SEALDATE, 0, "Seal Date"
    .ColWidth(E_INNER_BOX_INFO.E_PN) = 16
    .ColWidth(E_INNER_BOX_INFO.E_LOTNO) = 10
    .ColWidth(E_INNER_BOX_INFO.E_qty) = 8
    .ColWidth(E_INNER_BOX_INFO.E_DATECODE) = 8
    .ColWidth(E_INNER_BOX_INFO.E_SN) = 14
    .ColWidth(E_INNER_BOX_INFO.E_SEALDATE) = 8
    .Col = E_INNER_BOX_INFO.E_SN
    .BackColor = &HFF00&
    .ReDraw = True

End With

With Fps(1)
    .ReDraw = False
    .MaxCols = E_OUTER_BOX_INFO.E_END - 1
    .MaxRows = 0
    .FontBold = False
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    .TypeHAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
    .SetText E_OUTER_BOX_INFO.E_CID, 0, "外箱SN"
    .SetText E_OUTER_BOX_INFO.E_QID, 0, "外箱Q箱号"
    .SetText E_OUTER_BOX_INFO.E_OUTBOX_QTY, 0, "外箱数量"
    .ColWidth(E_OUTER_BOX_INFO.E_CID) = 12
    .ColWidth(E_OUTER_BOX_INFO.E_QID) = 12
    .ColWidth(E_OUTER_BOX_INFO.E_OUTBOX_QTY) = 10
    .Col = E_OUTER_BOX_INFO.E_CID
    .BackColor = &H80FFFF
    .ReDraw = True

End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cmdOpen_Click
' Description:       开始扫描
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/31-11:24:14
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdOpen_Click()

If cboCustCode.text = "" Then
    MsgBox "请选择客户代码", vbInformation, "提示"
    Exit Sub
Else
    cboCustCode.Locked = True

End If

If txtDN.text = "" Then
    MsgBox "DN不可为空", vbInformation, "提示"
    Exit Sub
Else
    If Get_SqlStr("SELECT * FROM erptemp..ht_dn where DN_NUM = '" & UCase$(Trim$(txtDN.text)) & "'") = "" Then
        MsgBox "DN不存在", vbInformation, "提示"
        Exit Sub

    End If

    If Get_OracleCnt("select * from packing_detailed_gd108 where ship_dn = '" & UCase$(Trim$(txtDN.text)) & "'") > 0 Then
        If MsgBox("该DN:" & UCase$(Trim$(txtDN.text)) & "已经有历史扫描打印记录,是否删除原纪录?", vbYesNo, "是否删除?") = vbYes Then
            AddSql ("delete from packing_detailed_gd108 where ship_dn = '" & UCase$(Trim$(txtDN.text)) & "'")

        End If

    End If

    txtDN.Locked = True

End If

txtScan.Visible = True
txtScan.SetFocus
Call PlaySound("请扫描卷盘号")

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       txtScan_KeyPress
' Description:       扫描入口
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/31-11:33:48
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub txtScan_KeyPress(KeyAscii As Integer)
Dim strScan As String

If KeyAscii <> vbKeyReturn Then Exit Sub
strScan = UCase(Trim$(txtScan.text))

Select Case cboCustCode.text

    Case "GD108"
        Call ScanningHandler_GD108(strScan)

End Select

txtScan.text = ""

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ScanningHandler_GD108
' Description:       扫描处理_GD108
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/31-11:27:32
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ScanningHandler_GD108(strScan As String)
Dim tLblInfo As T_INNER_BOX_INFO

If Not GetInnerBoxLblContent_GD108(tLblInfo, strScan) Then Exit Sub
If Not ChkInnerBoxLblContent_GD108(tLblInfo) Then Exit Sub
Call LstInnerBoxLblContent_GD108(tLblInfo)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetInnerBoxLblContent_GD108
' Description:       获取内盒标签内容
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/6-14:39:04
'
' Parameters :       strInnerBoxSN (String)
'--------------------------------------------------------------------------------
Private Function GetInnerBoxLblContent_GD108(ByRef tLblInfo As T_INNER_BOX_INFO, _
                                             ByRef strInnerBoxSN As String) As Boolean
Dim strSql As String
Dim rs     As New ADODB.Recordset

GetInnerBoxLblContent_GD108 = False
If InStr(strInnerBoxSN, "-B") = 0 Then
    MsgBox "请扫描-B的条形码", vbInformation, "提示"
    Exit Function

End If

strSql = "select * from packing_detailed_gd108 where ship_dn = '" & UCase$(Trim$(txtDN.text)) & "' and reel_code = '" & strInnerBoxSN & "'"
If Get_OracleCnt(strSql) > 0 Then
    MsgBox "该箱号已经打印过,请勿重复扫描", vbInformation, "提示"
    Exit Function

End If

If Left$(strInnerBoxSN, 2) = "SI" Then
    strSql = "select t2.CUST_DEVICE,t2.CUST_LOT,t2.DC,t1.箱号,SUM(t1.数量) as QTY from erpdata..tblStockNumSub t1 ,erptemp..ht_dn t2  where t2.DN_NUM = '" & UCase(Trim(txtDN.text)) & "' and t1.箱号 = '" & strInnerBoxSN & "'  " & "group by t2.CUST_DEVICE,t2.CUST_LOT,t2.DC,t1.箱号 "
Else
    strSql = "select t2.CUST_DEVICE,t2.CUST_LOT,t2.DC,t1.箱号,SUM(t1.数量) as QTY from erpdata..tblStockNumSub t1 " & "inner join erptemp..ht_dn t2 on t1.工单号 = t2.CUST_LOT  " & "where t2.DN_NUM = '" & UCase(Trim(txtDN.text)) & "' and t1.箱号 = '" & strInnerBoxSN & "'  " & "group by t2.CUST_DEVICE,t2.CUST_LOT,t2.DC,t1.箱号 "
End If

Set rs = Get_SqlserveRs(strSql)
If rs.RecordCount = 0 Then
    MsgBox "库存中找不到该箱号: " & strInnerBoxSN, vbCritical, "警告"
    Exit Function

End If

tLblInfo.T_PN = Trim("" & rs!cust_device)
tLblInfo.T_LOTNO = Trim("" & rs!CUST_LOT)
tLblInfo.T_QTY = rs!QTY
tLblInfo.T_DATECODE = Trim("" & rs!DC)
tLblInfo.T_SN = Trim("" & rs!箱号)
tLblInfo.T_SEALDATE = ""
GetInnerBoxLblContent_GD108 = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ChkInnerBoxLblContent_GD108
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/6-11:30:54
'
' Parameters :       strInnerBoxSN (String)
'--------------------------------------------------------------------------------
Private Function ChkInnerBoxLblContent_GD108(ByRef tLblInfo As T_INNER_BOX_INFO) As Boolean
Dim strSql  As String
Dim i       As Integer
Dim strHTPN As String

ChkInnerBoxLblContent_GD108 = False

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_INNER_BOX_INFO.E_SN
        If .text = tLblInfo.T_SN Then
            Call PlaySound("请勿重复扫描")
            Exit Function

        End If

        .Col = E_INNER_BOX_INFO.E_DATECODE
        If .text <> tLblInfo.T_DATECODE Then
            MsgBox "DATECODE不一致,不可包在一起", vbCritical, "警告"
            Exit Function

        End If

        .Col = E_INNER_BOX_INFO.E_LOTNO
        If .text <> tLblInfo.T_LOTNO Then
            MsgBox "批号不一致,请扫描同一批号的卷盘", vbCritical, "警告"
            Exit Function

        End If

    Next

End With

ChkInnerBoxLblContent_GD108 = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ShowBoxSN
' Description:       显示已扫描内盒SN
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/5-17:01:39
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LstInnerBoxLblContent_GD108(ByRef tLblInfo As T_INNER_BOX_INFO)
Dim i As Long

With Fps(0)
    .MaxRows = .MaxRows + 1
    i = .MaxRows
    .SetText E_INNER_BOX_INFO.E_PN, i, tLblInfo.T_PN
    .SetText E_INNER_BOX_INFO.E_LOTNO, i, tLblInfo.T_LOTNO
    .SetText E_INNER_BOX_INFO.E_qty, i, tLblInfo.T_QTY
    .SetText E_INNER_BOX_INFO.E_DATECODE, i, tLblInfo.T_DATECODE
    .SetText E_INNER_BOX_INFO.E_SN, i, tLblInfo.T_SN
    .SetText E_INNER_BOX_INFO.E_SEALDATE, i, tLblInfo.T_SEALDATE

End With

Call PlaySound("卷盘已扫描")

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cmdNext_Click
' Description:       整合-打印-切换外箱
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/5-17:34:36
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdNext_Click()
If Fps(0).MaxRows = 0 Then
    MsgBox "请先扫描需要合箱的内盒", vbInformation, "提示"
    Exit Sub

End If

Select Case cboCustCode.text

    Case "GD108"
        Call NextHandler_GD108

End Select

End Sub

Private Sub NextHandler_GD108()
Dim tOuterLblInfo As T_OUTER_BOX_INFO

Call PrintOuterBoxLbl_GD108(tOuterLblInfo)
Call LstOuterBoxLblContent_GD108(tOuterLblInfo)
Call SavePackingDetail_GD108(tOuterLblInfo)
Call UpdateERP_CARTON_NO(tOuterLblInfo.T_SN, tOuterLblInfo.T_QTY, Trim(txtDN.text))
Call NextOuterBox

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PrintOuterBoxLbl_GD108
' Description:       获取外箱标签内容
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/6-15:49:54
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function PrintOuterBoxLbl_GD108(ByRef tOuterLblInfo As T_OUTER_BOX_INFO) As Boolean
Dim lQty As Long
Dim i    As Integer
Dim strSN As String

PrintOuterBoxLbl_GD108 = False

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_INNER_BOX_INFO.E_qty
        lQty = lQty + CLng(.text)
        .Col = E_INNER_BOX_INFO.E_PN
        tOuterLblInfo.T_PN = .text
        .Col = E_INNER_BOX_INFO.E_LOTNO
        tOuterLblInfo.T_LOTNO = .text
        .Col = E_INNER_BOX_INFO.E_DATECODE
        tOuterLblInfo.T_DATECODE = .text
        .Col = E_INNER_BOX_INFO.E_SN
        strSN = .text
        
    Next

End With

tOuterLblInfo.T_QTY = lQty

If Left(strSN, 2) = "SI" Then
    tOuterLblInfo.T_SN = GetOuterBoxLblSN(Left$(strSN, 12))
Else
    tOuterLblInfo.T_SN = GetOuterBoxLblSN(tOuterLblInfo.T_LOTNO)
End If

tOuterLblInfo.T_QID = GetOuterBoxLblQID
tOuterLblInfo.T_PACKING_DATE1 = Year(Now) & Right$("00" & Month(Now), 2) & Right$("00" & Day(Now), 2)
txtPartNo.text = tOuterLblInfo.T_PN

txtLotNo.text = tOuterLblInfo.T_LOTNO
txtDateCode.text = tOuterLblInfo.T_DATECODE
txtQTY.text = tOuterLblInfo.T_QTY
txtSN.text = tOuterLblInfo.T_SN
txtSealDate.text = tOuterLblInfo.T_PACKING_DATE1
Dialog_OuterBoxLbl_GD108.Show 1
PrintOuterBoxLbl_GD108 = True
Call PlaySound("外箱标签已打印完成")

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       LstOuterBoxLblContent_GD108
' Description:       显示外箱列表
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/6-16:54:26
'
' Parameters :       tOuterLblInfo (T_INNER_BOX_INFO)
'--------------------------------------------------------------------------------
Private Sub LstOuterBoxLblContent_GD108(ByRef tOuterLblInfo As T_OUTER_BOX_INFO)
Dim i As Long

With Fps(1)
    .MaxRows = .MaxRows + 1
    i = .MaxRows
    .SetText E_OUTER_BOX_INFO.E_CID, i, tOuterLblInfo.T_SN
    .SetText E_OUTER_BOX_INFO.E_QID, i, tOuterLblInfo.T_QID
    .SetText E_OUTER_BOX_INFO.E_OUTBOX_QTY, i, tOuterLblInfo.T_QTY

End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetOuterBoxLblSN
' Description:       获取外箱C标签SN
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/6-15:59:07
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function GetOuterBoxLblSN(strLotID As String) As String
Dim strSql   As String
Dim strMaxID As String

strSql = "select * from erptemp..tblGD108lotboxidseq where lotid = '" & strLotID & "' "
If Get_SqlserverCnt(strSql) = 0 Then
    strMaxID = "1"
    AddSql2 ("insert into erptemp..tblGD108lotboxidseq(LOTID,BOXTYPE,MAXSEQ) values('" & strLotID & "','-C','" & strMaxID & "')")
Else
    strSql = "select maxseq+1 from erptemp..tblGD108lotboxidseq where lotid = '" & strLotID & "' and boxtype = '-C'"
    strMaxID = Get_SqlStr(strSql)
    AddSql2 ("update erptemp..tblGD108lotboxidseq set MAXSEQ = '" & strMaxID & "' where lotid = '" & strLotID & "' and boxtype = '-C' ")

End If

GetOuterBoxLblSN = strLotID & "-C" & Right$("0000" & strMaxID, 2)

End Function

Private Function GetOuterBoxLblQID() As String
Dim strKey As String

With Fps(0)
    .Row = 1
    .Col = E_INNER_BOX_INFO.E_SN
    strKey = .text

End With

GetOuterBoxLblQID = Get_OracleStr("select trglabelseq.QTSeq_NotMesQbox('" & strKey & "')  from dual")

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       SavePackingDetail_GD108
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/8/6-16:18:49
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub SavePackingDetail_GD108(ByRef tOuterLblInfo As T_OUTER_BOX_INFO)
Dim strSql        As String
Dim i             As Integer
Dim tInnerLblInfo As T_INNER_BOX_INFO

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = E_INNER_BOX_INFO.E_PN
        tInnerLblInfo.T_PN = Trim$("" & .text)
        .Col = E_INNER_BOX_INFO.E_LOTNO
        tInnerLblInfo.T_LOTNO = Trim$("" & .text)
        .Col = E_INNER_BOX_INFO.E_qty
        tInnerLblInfo.T_QTY = Trim$("" & .text)
        .Col = E_INNER_BOX_INFO.E_DATECODE
        tInnerLblInfo.T_DATECODE = Trim$("" & .text)
        .Col = E_INNER_BOX_INFO.E_SN
        tInnerLblInfo.T_SN = Trim$("" & .text)
        strSql = "delete from PACKING_DETAILED_GD108 where reel_code = '" & tInnerLblInfo.T_SN & "'"
        AddSql (strSql)
        strSql = "insert into PACKING_DETAILED_GD108(ID,REEL_CODE,PART_NO,REEL_QTY,DATE_CODE,OUTER_BOX_CID,OUTER_BOX_QID,SHIP_DN,CREATE_DATE,CREATE_BY,PRINT_FLAG,LOTID,WAFERIDLST) " & " values(GD108_REEL_SEQ.Nextval,'" & tInnerLblInfo.T_SN & "','" & tInnerLblInfo.T_PN & "','" & tInnerLblInfo.T_QTY & "','" & tInnerLblInfo.T_DATECODE & "','" & tOuterLblInfo.T_SN & "','" & tOuterLblInfo.T_QID & "','" & txtDN.text & "',sysdate,'" & gUserName & "','1','" & tOuterLblInfo.T_LOTNO & "','')   "
        AddSql (strSql)
    Next

End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       UpdateERP_CARTON_NO
' Description:       更新ERP箱号对照关系
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/10-12:03:00
'
' Parameters :       strDN (String)
'--------------------------------------------------------------------------------
Private Sub UpdateERP_CARTON_NO(strCartonID As String, _
                                lCartonQty As Long, _
                                strDN As String)
Dim strSql As String
Dim rs     As ADODB.Recordset
Dim id     As String

On Error GoTo ERRON

INIadoCon.BeginTrans
' ---------------------------------------------------删除
'0
strSql = "delete from [erpdata].[dbo].[tblPackTreeInf] where 箱号 = '" & strCartonID & "'"
AddSql2 (strSql)
strSql = "delete from [erpdata].[dbo].[tblPackMainInf] where 箱号 = '" & strCartonID & "'"
AddSql2 (strSql)
strSql = "update [erpdata].[dbo].[tblPackTreeInf] set 上级序号 = '', Memo = '' where 箱号 in ( select * from  OPENQUERY(ORACLEDB, 'select REEL_CODE from PACKING_DETAILED_GD108 where OUTER_BOX_CID = ''" & strCartonID & "'' and SHIP_DN = ''" & strDN & "'' '))  "
AddSql2 (strSql)
strSql = "delete from [erpdata].[dbo].[tblStockNumTree] where 箱号 = '" & strCartonID & "'"
AddSql2 (strSql)
strSql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='',Memo='', dn='' where 箱号 in ( select * from  OPENQUERY(ORACLEDB, 'select REEL_CODE from PACKING_DETAILED_GD108 where OUTER_BOX_CID = ''" & strCartonID & "'' and SHIP_DN = ''" & strDN & "'' '))    "
AddSql2 (strSql)
' --------------------------------------------------更新
'1 insert [erpdata].[dbo].[tblPackMainInf]
strSql = "insert into [erpdata].[dbo].[tblPackMainInf](箱号,客户代码,装箱时间,数量,合格标记,装箱标记,标签标记,真空标记,完成标记,基层标记,产线标记) values('" & strCartonID & "','GD108',GetDate()," & lCartonQty & ",'0','1','1','0','1','0','1')"
If AddSql2(strSql) = 0 Then
    MsgBox "1 insert [erpdata].[dbo].[tblPackMainInf]:failed!!! ", vbCritical, "提示"
    Exit Sub

End If

'2 insert - update [erpdata].[dbo].[tblPackTreeInf]
strSql = "insert into [erpdata].[dbo].[tblPackTreeInf](箱号,上级序号,基层标记 ,Memo) values('" & strCartonID & "',0,1,'GD108')"
If AddSql2(strSql) = 0 Then
    MsgBox "2 insert [erpdata].[dbo].[tblPackTreeInf]:failed!!!", vbCritical, "提示"
    Exit Sub

End If

id = Get_SqlserverNo("select 序号 as ID from [erpdata].[dbo].[tblPackTreeInf] a where a.箱号='" & strCartonID & "' and Memo='GD108' ")
strSql = "Update [erpdata].[dbo].[tblPackTreeInf] set 上级序号='" & id & "',Memo='GD108' " & " where 箱号 in ( select * from  OPENQUERY(ORACLEDB, 'select REEL_CODE from PACKING_DETAILED_GD108 where OUTER_BOX_CID = ''" & strCartonID & "'' and SHIP_DN = ''" & strDN & "'' ')) "
If AddSql2(strSql) = 0 Then
    MsgBox "2 update [erpdata].[dbo].[tblPackTreeInf]:failed!!!", vbCritical, "提示"
    Exit Sub

End If

'3 insert - update [erpdata].[dbo].[tblStockNumTree]
strSql = "insert into [erpdata].[dbo].[tblStockNumTree](序号,箱号,上级序号,基层标记 ,尺寸,重量,Memo,DN,仓位) values( " & id & ",'" & strCartonID & "',0,1,'','','GD108','" & strDN & "','走空账，实物在产线')"
If AddSql2(strSql) = 0 Then
    MsgBox "3 insert [erpdata].[dbo].[tblStockNumTree]", vbCritical, "提示"
    Exit Sub

End If

strSql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & id & "',Memo='GD108', dn='" & strDN & "'  where 箱号 in ( select * from  OPENQUERY(ORACLEDB, 'select REEL_CODE from PACKING_DETAILED_GD108 where OUTER_BOX_CID = ''" & strCartonID & "'' and SHIP_DN = ''" & strDN & "'' '))  "
If AddSql2(strSql) = 0 Then
    MsgBox "3 update [erpdata].[dbo].[tblStockNumTree]", vbCritical, "提示"

    'Exit Sub
End If

INIadoCon.CommitTrans
MsgBox strCartonID & ":箱号已更新", vbInformation, "提示"
Exit Sub
ERRON:
INIadoCon.RollbackTrans
MsgBox "错误:" & Err.DESCRIPTION, vbCritical, "警告"

End Sub

Private Sub NextOuterBox()
Fps(0).MaxRows = 0
txtScan.SetFocus
Call PlaySound("请继续扫描下一个外箱")

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PlaySound
' Description:       播放声音文件
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-9:51:39
'
' Parameters :       strSound (String)
'--------------------------------------------------------------------------------
Private Sub PlaySound(strSound As String)
player1.url = Trim(txtMediaDir.text) & strSound & ".wav"

End Sub


