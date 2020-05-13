VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmCheckLblSys_57 
   Caption         =   "57标签核对绑定系统"
   ClientHeight    =   13290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15615
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
   ScaleHeight     =   13290
   ScaleWidth      =   15615
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      Caption         =   "菜单选项"
      ForeColor       =   &H00800000&
      Height          =   1935
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15615
      Begin VB.TextBox txtShipID 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7440
         TabIndex        =   16
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtScan 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   720
         TabIndex        =   14
         Top             =   240
         Width           =   5055
      End
      Begin VB.TextBox txtMediaDir 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Text            =   "C:\media_source\"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtASNLblVal 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   9855
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出"
         Height          =   360
         Left            =   8160
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtPSNLblVal 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   9855
      End
      Begin VB.TextBox txtReelLblVal 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   9855
      End
      Begin VB.Label lblShipID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出货虚拟单号"
         Height          =   195
         Left            =   6240
         TabIndex        =   15
         Top             =   1605
         Width           =   1080
      End
      Begin VB.Label lblScan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   285
         Width           =   420
      End
      Begin VB.Label lblMediaDir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "声音文件目录"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   1605
         Width           =   1080
      End
      Begin WMPLibCtl.WindowsMediaPlayer player1 
         Height          =   495
         Left            =   11280
         TabIndex        =   10
         Top             =   120
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
      Begin VB.Label lblASN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ASN标签二维码"
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1245
         Width           =   1200
      End
      Begin VB.Label lblPSN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PSN标签二维码"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1005
         Width           =   1185
      End
      Begin VB.Label lblReelID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘标签二维码"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   765
         Width           =   1260
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "标签绑定对照表"
      ForeColor       =   &H00800000&
      Height          =   13215
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   15615
      Begin FPSpreadADO.fpSpread fps 
         Height          =   11775
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   15375
         _Version        =   524288
         _ExtentX        =   27120
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
         SpreadDesigner  =   "FrmCheckLblSys_57.frx":0000
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "FrmCheckLblSys_57"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum E_REEL_PSN_ASN

    E_REEL_ID = 1
    E_PSN
    E_PN
    E_MPN
    E_MFG_CODE
    E_CUST_DEVICE
    E_PRODUCT
    E_M_LOT
    E_QTY
    E_WO_ID
    E_CUST_LOT_ID
    E_CUST_WAFER_ID
    E_DATE_CODE
    E_END

End Enum

Private Type T_REEL_INFO

    T_HT_PN As String
    T_CUST_PN As String
    T_PN As String
    T_MPN As String
    T_M_LOT_ID As String
    T_DATE_CODE As String
    T_LOT_ID As String
    T_WAWFER_ID As String
    T_QTY As Long
    T_REEL_ID As String
    T_WOID As String
    T_GrossDie As Long

End Type

Private Type T_PSN_INFO
    T_PN As String
    T_MFG_CODE As String
    T_MPN As String
    T_M_LOT_ID As String
    T_PSN As String
    T_QTY As Long
    T_DATE_CODE As String
            
End Type

Private Const strSplitFlag = "@$"

Private Const iSplitCnt_ReelLbl = 8

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub Form_Activate()
If txtScan.Enabled Then
    txtScan.SetFocus

End If

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
' Date-Time  :       2019/7/24-9:07:59
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitData()
txtShipID.Text = "S" & Right(Year(Now), 2) & Right$("00" & Month(Now), 2) & Right("00" & Day(Now), 2) & Right$("0000" & Get_OracleStr("select SEQ_57SHIP.NEXTVAL from dual"), 4)
Call PlaySound("请扫描卷盘标签二维码")

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitCtrls
' Description:       初始化控件
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-9:08:16
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitCtrls()
Call InitFps

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       InitFps
' Description:       初始化FPS
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-9:08:46
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub InitFps()

With Fps(0)
    .ReDraw = False
    .MaxCols = E_REEL_PSN_ASN.E_END - 1
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
    .SetText E_REEL_PSN_ASN.E_REEL_ID, 0, "REELID"
    .SetText E_REEL_PSN_ASN.E_PSN, 0, "PSN"
    .SetText E_REEL_PSN_ASN.E_PN, 0, "PN"
    .SetText E_REEL_PSN_ASN.E_MPN, 0, "MPN"
    .SetText E_REEL_PSN_ASN.E_MFG_CODE, 0, "MFG_CODE"
    .SetText E_REEL_PSN_ASN.E_CUST_DEVICE, 0, "客户机种"
    .SetText E_REEL_PSN_ASN.E_PRODUCT, 0, "场内机种"
    .SetText E_REEL_PSN_ASN.E_M_LOT, 0, "M_LOT"
    .SetText E_REEL_PSN_ASN.E_QTY, 0, "卷盘数量"
    .SetText E_REEL_PSN_ASN.E_WO_ID, 0, "工单号"
    .SetText E_REEL_PSN_ASN.E_CUST_LOT_ID, 0, "LOTID"
    .SetText E_REEL_PSN_ASN.E_CUST_WAFER_ID, 0, "WAFERID"
    .SetText E_REEL_PSN_ASN.E_DATE_CODE, 0, "DATECODE"
    .ColWidth(E_REEL_PSN_ASN.E_REEL_ID) = 12
    .ColWidth(E_REEL_PSN_ASN.E_PSN) = 16
    .ColWidth(E_REEL_PSN_ASN.E_PN) = 8
    .ColWidth(E_REEL_PSN_ASN.E_MPN) = 8
    .ColWidth(E_REEL_PSN_ASN.E_MFG_CODE) = 6
    .ColWidth(E_REEL_PSN_ASN.E_CUST_DEVICE) = 8
    .ColWidth(E_REEL_PSN_ASN.E_PRODUCT) = 8
    .ColWidth(E_REEL_PSN_ASN.E_M_LOT) = 8
    .ColWidth(E_REEL_PSN_ASN.E_QTY) = 6
    .ColWidth(E_REEL_PSN_ASN.E_WO_ID) = 10
    .ColWidth(E_REEL_PSN_ASN.E_CUST_LOT_ID) = 8
    .ColWidth(E_REEL_PSN_ASN.E_CUST_WAFER_ID) = 10
    .ColWidth(E_REEL_PSN_ASN.E_DATE_CODE) = 10
    
    .Col = E_REEL_PSN_ASN.E_REEL_ID
    .BackColor = &HFF00&
    .Col = E_REEL_PSN_ASN.E_PSN
    .BackColor = &H80FFFF
    
    .ReDraw = True

End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       txtScan_KeyPress
' Description:       扫描入口
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-11:31:45
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub txtScan_KeyPress(KeyAscii As Integer)
Dim strScan As String

If KeyAscii <> vbKeyReturn Then Exit Sub
txtReelLblVal.BackColor = vbWhite
txtPSNLblVal.BackColor = vbWhite

strScan = UCase$(Trim$(txtScan.Text))
If txtReelLblVal.Text = "" Then
    Call GetReelLblInfo(strScan)
ElseIf txtPSNLblVal.Text = "" Then
    Call GetPSNLblInfo(strScan)

End If

'If txtReelLblVal.Text = "" Then
'    Call PlaySound("请扫描卷盘标签二维码")
'ElseIf txtPSNLblVal.Text = "" Then
'    Call PlaySound("卷盘标签已扫描,请扫描PSN二维码")
'
'End If

If txtReelLblVal.Text <> "" And txtPSNLblVal.Text <> "" Then
    Call BindReel_PSN
    
    txtReelLblVal.Text = ""
    txtPSNLblVal.Text = ""
End If


txtScan.Text = ""

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetReelLblInfo
' Description:       获取卷盘标签信息
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-10:54:06
'
' Parameters :       strScan (String)
'--------------------------------------------------------------------------------
Private Sub GetReelLblInfo(strScan As String)
Dim strArray() As String
Dim tReelInfo  As T_REEL_INFO

If InStr(strScan, strSplitFlag) = 0 Then
    MsgBox "扫描错误,请扫描正确的二维码", vbCritical, "警告"
    Exit Sub

End If

strArray = Split(strScan, strSplitFlag)
If UBound(strArray) <> iSplitCnt_ReelLbl Then
    MsgBox "扫描错误,请扫描正确的卷盘标签二维码", vbCritical, "警告"
    Exit Sub

End If

If InStr(strArray(0), "/") = 0 Then
    MsgBox "二维码格式不正确", vbCritical, "警告"
    Exit Sub

End If

tReelInfo.T_HT_PN = Split(strArray(0), "/")(0)
tReelInfo.T_CUST_PN = Split(strArray(0), "/")(1)
tReelInfo.T_PN = strArray(1)
tReelInfo.T_MPN = strArray(2)
tReelInfo.T_M_LOT_ID = strArray(3)
tReelInfo.T_DATE_CODE = strArray(4)
tReelInfo.T_LOT_ID = strArray(5)
tReelInfo.T_WAWFER_ID = strArray(5) & strArray(6)
tReelInfo.T_QTY = strArray(7)
tReelInfo.T_REEL_ID = strArray(8)

If Not ChkReelLblInfo(tReelInfo) Then Exit Sub
txtReelLblVal.Text = strScan
Call PlaySound("卷盘标签已扫描,请扫描PSN二维码")
End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ChkReelLblInfo
' Description:       检查卷盘标签信息
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-11:35:04
'
' Parameters :       tReelInfo (T_REEL_INFO)
'--------------------------------------------------------------------------------
Private Function ChkReelLblInfo(tReelInfo As T_REEL_INFO) As Boolean
Dim strSql As String

ChkReelLblInfo = False
strSql = "select * from erptemp..TRAY_PSN_LIST where TRAY_ID = '" & tReelInfo.T_REEL_ID & "'"
If Get_SqlserverCnt(strSql) > 0 Then
    Call PlaySound("该卷盘ID已经绑定过PSN,扫描错误")
    txtReelLblVal.BackColor = vbRed
    Exit Function
End If

ChkReelLblInfo = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetPSNLblInfo
' Description:       获取PSN标签信息
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-10:54:49
'
' Parameters :       strScan (String)
'--------------------------------------------------------------------------------
Private Sub GetPSNLblInfo(strScan As String)
Dim tPSNInfo  As T_PSN_INFO

If Left(strScan, 3) <> "[)>" Then
    MsgBox "扫描错误,请扫描正确的PSN标签二维码", vbCritical, "警告"
    Exit Sub

End If

tPSNInfo.T_PSN = Mid(strScan, InStr(strScan, "52S") + 3, InStr(strScan, "18VLEHWT") - InStr(strScan, "52S") - 3)

If Not ChkPSNLblInfo(tPSNInfo) Then Exit Sub
txtPSNLblVal.Text = strScan
Call PlaySound("PSN已扫描")
End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ChkPSNLblInfo
' Description:       检查PSN标签信息
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-16:41:56
'
' Parameters :       tPSNInfo (T_PSN_INFO)
'--------------------------------------------------------------------------------
Private Function ChkPSNLblInfo(tPSNInfo As T_PSN_INFO)
Dim strSql As String

ChkPSNLblInfo = False
strSql = "select * from erptemp..TRAY_PSN_LIST where PSN_ID = '" & tPSNInfo.T_PSN & "'"
If Get_SqlserverCnt(strSql) > 0 Then
    Call PlaySound("该PSN已经绑定过卷盘,扫描错误")
    txtPSNLblVal.BackColor = vbRed
    Exit Function
End If

ChkPSNLblInfo = True

End Function


'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       GetASNLblInfo
' Description:       获取ASN标签信息
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-10:55:16
'
' Parameters :       strScan (String)
'--------------------------------------------------------------------------------
Private Sub GetASNLblInfo(strScan As String)

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       BindReel_PSN
' Description:       绑定REEL_PSN
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-17:24:03
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub BindReel_PSN()
Dim strReelLblVal As String
Dim strPSNLblVal  As String
Dim strArray() As String
Dim tReelInfo  As T_REEL_INFO
Dim tPSNInfo  As T_PSN_INFO

strReelLblVal = Trim(txtReelLblVal.Text)
strPSNLblVal = Trim$(txtPSNLblVal.Text)

'ReelLblInfo
strArray = Split(strReelLblVal, strSplitFlag)
tReelInfo.T_HT_PN = Split(strArray(0), "/")(0)
tReelInfo.T_CUST_PN = Split(strArray(0), "/")(1)
tReelInfo.T_PN = strArray(1)
tReelInfo.T_MPN = strArray(2)
tReelInfo.T_M_LOT_ID = strArray(3)
tReelInfo.T_DATE_CODE = strArray(4)
tReelInfo.T_LOT_ID = strArray(5)
tReelInfo.T_WAWFER_ID = strArray(5) & strArray(6)
tReelInfo.T_QTY = strArray(7)
tReelInfo.T_REEL_ID = strArray(8)
tReelInfo.T_WOID = Get_OracleStr("select ordername from ib_waferlist where waferid = '" & tReelInfo.T_WAWFER_ID & "'")
tReelInfo.T_GrossDie = Get_OracleNo("select passbincount+failbincount from mappingdatatest where substrateid  =  '" & tReelInfo.T_WAWFER_ID & "' ")

'PSNLblInfo
strPSNLblVal = Replace$(strPSNLblVal, "F01001P", "")
tPSNInfo.T_PSN = Mid(strPSNLblVal, InStr(strPSNLblVal, "52S") + 3, InStr(strPSNLblVal, "18VLEHWT") - InStr(strPSNLblVal, "52S") - 3)
tPSNInfo.T_PN = Mid(strPSNLblVal, InStr(strPSNLblVal, "1P") + 2, InStr(strPSNLblVal, "1V") - InStr(strPSNLblVal, "1P") - 2)
tPSNInfo.T_MFG_CODE = Mid(strPSNLblVal, InStr(strPSNLblVal, "1V") + 2, InStr(strPSNLblVal, "10D") - InStr(strPSNLblVal, "1V") - 2)
tPSNInfo.T_M_LOT_ID = Mid(strPSNLblVal, InStr(strPSNLblVal, "1T") + 2, 12)
tPSNInfo.T_QTY = Val(Mid(Right$(strPSNLblVal, 10), InStr(Right$(strPSNLblVal, 10), "Q") + 1))
tPSNInfo.T_DATE_CODE = Mid(strPSNLblVal, InStr(strPSNLblVal, "10D") + 3, 4)

'Check
If Not ChkReel_PSN(tReelInfo, tPSNInfo) Then Exit Sub
'Save
Call SaveBindInfo(tReelInfo, tPSNInfo)
'Show
Call ShowBindInfo

Call PlaySound("PSN已绑定")
End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ChkReel_PSN
' Description:       检查是否符合绑定要求
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-17:38:01
'
' Parameters :       tReelInfo (T_REEL_INFO)
'                    tPSNInfo (T_PSN_INFO)
'--------------------------------------------------------------------------------
Private Function ChkReel_PSN(tReelInfo As T_REEL_INFO, tPSNInfo As T_PSN_INFO) As Boolean
ChkReel_PSN = False
If tReelInfo.T_PN <> tPSNInfo.T_PN Then
    MsgBox "PN不匹配,无法绑定", vbCritical, "警告"
    Exit Function
End If
If tReelInfo.T_DATE_CODE <> tPSNInfo.T_DATE_CODE Then
    MsgBox "DATECODE不匹配,无法绑定", vbCritical, "警告"
    Exit Function
End If
If tReelInfo.T_M_LOT_ID <> tPSNInfo.T_M_LOT_ID Then
    MsgBox "M.LOT不匹配,无法绑定", vbCritical, "警告"
    Exit Function
End If
If tReelInfo.T_QTY <> tPSNInfo.T_QTY Then
    MsgBox "数量不匹配,无法绑定", vbCritical, "警告"
    Exit Function
End If

ChkReel_PSN = True
End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       SaveBindInfo
' Description:       保存绑定对照表
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/25-10:36:27
'
' Parameters :       tReelInfo (T_REEL_INFO)
'                    tPSNInfo (T_PSN_INFO)
'--------------------------------------------------------------------------------
Private Sub SaveBindInfo(tReelInfo As T_REEL_INFO, tPSNInfo As T_PSN_INFO)
Dim strSql As String

strSql = "insert into erptemp..TRAY_PSN_LIST(TRAY_ID,PSN_ID,PN,MFG_CODE,MPN,M_LOT,QTY,PRODUCT,ORDER_NAME,CUST_LOT,WAFER_ID,CUST_DEVICE,WAFER_DIE,PSN_DC,CREATE_BY,CREATE_DATE,FLAG,REMARK1) " & _
" values('" & tReelInfo.T_REEL_ID & "','" & tPSNInfo.T_PSN & "','" & tReelInfo.T_PN & "','" & tPSNInfo.T_MFG_CODE & "','" & tReelInfo.T_MPN & "','" & tReelInfo.T_M_LOT_ID & "','" & tReelInfo.T_QTY & "','" & tReelInfo.T_HT_PN & "','" & tReelInfo.T_WOID & "','" & tReelInfo.T_LOT_ID & "','" & tReelInfo.T_WAWFER_ID & "','" & tReelInfo.T_CUST_PN & "','" & tReelInfo.T_GrossDie & "','" & tPSNInfo.T_DATE_CODE & "','" & gUserName & "',GetDate(),'0','" & Trim(txtShipID.Text) & "')"

AddSql2 (strSql)

End Sub
'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ShowBindInfo
' Description:       显示绑定对照表
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/7/24-17:42:47
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ShowBindInfo()
Dim strSql     As String
Dim rsREEL_PSN As New ADODB.Recordset

strSql = "select TRAY_ID,PSN_ID,PN,MPN,MFG_CODE,CUST_DEVICE,PRODUCT,M_LOT,QTY,ORDER_NAME,CUST_LOT,WAFER_ID,PSN_DC from erptemp..TRAY_PSN_LIST where remark1 = '" & Trim(txtShipID.Text) & "' order by CREATE_DATE desc"
Set rsREEL_PSN = Get_SqlserveRs(strSql)

With Fps(0)
    .MaxRows = 0
    If Not rsREEL_PSN.EOF Then
        Set .DataSource = rsREEL_PSN

    End If

End With

Set rsREEL_PSN = Nothing

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
player1.url = Trim(txtMediaDir.Text) & strSound & ".wav"

End Sub
