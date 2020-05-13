VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm57HW 
   Caption         =   "57出华为"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12795
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
   ScaleHeight     =   10200
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   12615
      Begin VB.TextBox txtQty 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   9840
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtPecs 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   9840
         TabIndex        =   15
         Top             =   315
         Width           =   975
      End
      Begin VB.TextBox txtPkgNO 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   6360
         TabIndex        =   12
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtDN 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   6360
         TabIndex        =   10
         Top             =   315
         Width           =   1935
      End
      Begin VB.CommandButton btnStart 
         BackColor       =   &H00FFC0FF&
         Caption         =   "出货扫描绑定开始"
         Height          =   480
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton btnFinish 
         BackColor       =   &H00E0E0E0&
         Caption         =   "出货扫描绑定完毕"
         Height          =   480
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   8655
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   11175
         _Version        =   524288
         _ExtentX        =   19711
         _ExtentY        =   15266
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
         MaxCols         =   10
         MaxRows         =   0
         SpreadDesigner  =   "Frm57HW.frx":0000
      End
      Begin VB.TextBox txtLabelQRCode_HW 
         BackColor       =   &H00FFC0FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   1545
         Width           =   10215
      End
      Begin VB.TextBox txtLabelQRCode_57 
         BackColor       =   &H00FFC0FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   1185
         Width           =   10215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前出货总片数:"
         Height          =   195
         Left            =   8400
         TabIndex        =   14
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前出货总数量:"
         Height          =   195
         Left            =   8400
         TabIndex        =   13
         Top             =   645
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出货虚拟大箱号"
         Height          =   195
         Left            =   5040
         TabIndex        =   11
         Top             =   645
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "虚拟挑料单据号"
         Height          =   195
         Left            =   5040
         TabIndex        =   9
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "华为标签二维码"
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
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "57 标签二维码"
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
         Top             =   1200
         Width           =   1395
      End
      Begin WMPLibCtl.WindowsMediaPlayer player1 
         Height          =   495
         Left            =   10440
         TabIndex        =   1
         Top             =   6000
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
   End
End
Attribute VB_Name = "Frm57HW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type T_REEL_INFO

    T_MPN As String
    T_PN As String
    T_M_LOT_ID As String
    T_DATE_CODE As String
    T_QTY As Long
    T_REEL_ID As String

End Type

Private Sub btnFinish_Click()
Call FinishScan(Trim(txtDN.text), Trim$(txtPkgNO.text))

End Sub

Private Sub btnStart_Click()
txtDN.text = Year(Now) & Right("0000" & Month(Now), 2) & Right$("0000" & Day(Now), 2) & Right("0000" & Get_OracleStr("select SEQ_57SHIPDN.Nextval from dual"), 4)
txtPkgNO.text = GetQID(Trim(txtDN.text))

Call PlaySound("请依次扫描5 7标签二维码和华为标签二维码")
txtLabelQRCode_57.Enabled = True
txtLabelQRCode_57.SetFocus

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       PlaySound
' Description:       播放声音
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/12/16-10:38:26
'
' Parameters :       strSound (String)
'--------------------------------------------------------------------------------
Private Sub PlaySound(strSound As String)
player1.url = "\\10.160.1.84\public\media_source\" & strSound & ".wav"

End Sub

Private Sub Form_Load()

With fps
    .ColWidth(1) = 20
    .ColWidth(2) = 20
    .ColWidth(3) = 40

End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       txtLabelQRCode_57_KeyPress
' Description:       57标签二维码扫描
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/12/16-10:59:22
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub txtLabelQRCode_57_KeyPress(KeyAscii As Integer)
Dim strLabelQRCode_57 As String

strLabelQRCode_57 = Trim$(txtLabelQRCode_57.text)
If KeyAscii <> vbKeyReturn Then Exit Sub
If Len(Trim$(txtLabelQRCode_57.text)) = 0 Then Exit Sub
'检查数据
If CheckLabelQRCode_57(strLabelQRCode_57) = False Then
    txtLabelQRCode_57.text = ""
    Exit Sub

End If

'状态切换
txtLabelQRCode_HW.Enabled = True
txtLabelQRCode_HW.SetFocus
txtLabelQRCode_57.Enabled = False
Call PlaySound("5 7二维码已扫描,请扫描华为二维码")

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckLabelQRCode_57
' Description:       检查57标签二维码数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/12/16-11:02:19
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function CheckLabelQRCode_57(strLabelQRCode_57 As String) As Boolean
Dim strSql       As String
Dim strReelID_57 As String

CheckLabelQRCode_57 = False
If InStr(strLabelQRCode_57, "@$") = 0 Then
    MsgBox "扫描错误,请扫描正确的二维码", vbCritical, "警告"
    Exit Function

End If

If UBound(Split(strLabelQRCode_57, "@$")) <> 8 Then
    MsgBox "扫描错误,请扫描正确的卷盘标签二维码", vbCritical, "警告"
    Exit Function

End If

strReelID_57 = Split(strLabelQRCode_57, "@$")(8)
If InStr(strReelID_57, "-R") = 0 Then
    MsgBox "57标签二维码格式有误", vbCritical, "警告"
    Exit Function

End If

strSql = "select * from erptemp..TRAY_PSN_LIST where TRAY_ID = '" & strReelID_57 & "'  "
If Get_SqlserverCnt(strSql) > 0 Then
    MsgBox "之前已经扫描绑定过该57卷盘ID : " & strReelID_57 & " 请确认是否有误", vbCritical, "警告"
    Exit Function

End If

CheckLabelQRCode_57 = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       txtLabelQRCode_57_KeyPress
' Description:       华为标签二维码扫描
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/12/16-10:59:22
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub txtLabelQRCode_HW_KeyPress(KeyAscii As Integer)
Dim strLabelQRCode_HW As String

strLabelQRCode_HW = Trim$(txtLabelQRCode_HW.text)
If KeyAscii <> vbKeyReturn Then Exit Sub
If Len(Trim$(txtLabelQRCode_HW.text)) = 0 Then Exit Sub
'检查数据
If CheckLabelQRCode_HW(strLabelQRCode_HW) = False Then
    txtLabelQRCode_HW.text = ""
    Exit Sub

End If

'关联数据
Call RelateReelID
'状态切换
txtLabelQRCode_57.text = ""
txtLabelQRCode_HW.text = ""
txtLabelQRCode_57.Enabled = True
txtLabelQRCode_57.SetFocus
txtLabelQRCode_HW.Enabled = False
Call PlaySound("卷盘已绑定,请继续扫描")

'Call CheckScanningComplate(txtDN.text)
End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       CheckLabelQRCode_57
' Description:       检查华为标签二维码数据
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/12/16-11:02:19
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function CheckLabelQRCode_HW(strLabelQRCode_HW As String) As Boolean
Dim strSql       As String
Dim strReelID_HW As String

CheckLabelQRCode_HW = False
If Left(strLabelQRCode_HW, 3) <> "[)>" Then
    MsgBox "扫描错误,请扫描正确的PSN标签二维码", vbCritical, "警告"
    Exit Function

End If

strReelID_HW = Mid(strLabelQRCode_HW, InStr(strLabelQRCode_HW, "52S") + 3, InStr(strLabelQRCode_HW, "18VLEHWT") - InStr(strLabelQRCode_HW, "52S") - 3)
strSql = "select * from erptemp..TRAY_PSN_LIST where PSN_ID = '" & strReelID_HW & "' " '  "
If Get_SqlserverCnt(strSql) > 0 Then
    MsgBox "之前已经扫描绑定过该华为卷盘ID : " & strReelID_HW & " 请确认是否有误", vbCritical, "警告"
    Exit Function

End If

CheckLabelQRCode_HW = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       RelateReelID
' Description:       关联卷盘ID
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/12/16-11:31:31
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function RelateReelID() As Boolean
Dim strLabelQRCode_HW As String
Dim strLabelQRCode_57 As String
Dim strArray()        As String
Dim strPSN            As String
Dim tReelInfo         As T_REEL_INFO
Dim strSql            As String

RelateReelID = False
strPSN = Mid(txtLabelQRCode_HW.text, InStr(txtLabelQRCode_HW.text, "52S") + 3, InStr(txtLabelQRCode_HW.text, "18VLEHWT") - InStr(txtLabelQRCode_HW.text, "52S") - 3)
strLabelQRCode_57 = txtLabelQRCode_57.text
strLabelQRCode_HW = txtLabelQRCode_HW.text
strArray = Split(strLabelQRCode_57, "@$")
tReelInfo.T_MPN = strArray(1)
tReelInfo.T_PN = strArray(2)
tReelInfo.T_M_LOT_ID = strArray(3)
tReelInfo.T_DATE_CODE = strArray(4)
tReelInfo.T_QTY = strArray(7)
tReelInfo.T_REEL_ID = strArray(8)
If InStr(strLabelQRCode_HW, "P" & tReelInfo.T_PN) = 0 Then
    MsgBox "PN不匹配", vbCritical, "警告"
    Exit Function

End If

If InStr(strLabelQRCode_HW, "1T" & tReelInfo.T_M_LOT_ID) = 0 Then
    MsgBox "MLOTID不匹配", vbCritical, "警告"
    Exit Function

End If

If InStr(strLabelQRCode_HW, "10D" & tReelInfo.T_DATE_CODE) = 0 Then
    MsgBox "DATECODE不匹配", vbCritical, "警告"
    Exit Function

End If

If InStr(strLabelQRCode_HW, "Q" & tReelInfo.T_QTY) = 0 Then
    MsgBox "数量不匹配", vbCritical, "警告"
    Exit Function

End If

'插入DN-Reel-绑定关系
strSql = "insert into erptemp..TRAY_PSN_LIST(TRAY_ID,PSN_ID,PN,MPN,M_LOT,QTY,PSN_DC,CREATE_BY,CREATE_DATE,FLAG,REMARK1) " & " values('" & tReelInfo.T_REEL_ID & "','" & strPSN & "','" & tReelInfo.T_PN & "','" & tReelInfo.T_MPN & "','" & tReelInfo.T_M_LOT_ID & "','" & tReelInfo.T_QTY & "','" & tReelInfo.T_DATE_CODE & "','" & gUserName & "',GetDate(),'0','" & Trim$(txtDN.text) & "')"
AddSql2 (strSql)
Call ShowData
RelateReelID = True

End Function

Private Sub ShowData()
Dim strSql As String
Dim rs     As New ADODB.Recordset

strSql = "select REMARK1 as 出货虚拟DN,TRAY_ID as 卷盘ID,PSN_ID as 华为PSN,QTY,CREATE_DATE as 绑定日期 from erptemp..TRAY_PSN_LIST where  Remark1 = '" & Trim(txtDN.text) & "'  order by create_date desc"
Set rs = Get_SqlserveRs(strSql)
fps.MaxRows = 0
If rs.RecordCount > 0 Then

    With fps
        Set .DataSource = rs

    End With

End If

txtPecs.text = Get_SqlStr("select count(1) from erptemp..TRAY_PSN_LIST where remark1 = '" & Trim(txtDN.text) & "'")
txtQty.text = Get_SqlStr("select sum(QTY) from erptemp..TRAY_PSN_LIST where remark1 = '" & Trim(txtDN.text) & "'")

End Sub

Private Function FinishScan(strDN As String, strPkgNo As String)

AddSql2 ("update erptemp..TRAY_PSN_LIST set REMARK2 = '" & strPkgNo & "' where REMARK1 = '" & strDN & "' ")
Call UpdateERP_CARTON_NO(strDN)

MsgBox "出货及箱号数据绑定完成", vbInformation, "提示"
txtDN.text = ""
txtPkgNO.text = ""
txtPecs.text = ""
txtQty.text = ""
fps.MaxRows = 0
txtLabelQRCode_57.Enabled = False
txtLabelQRCode_HW.Enabled = False

End Function

Private Sub UpdateERP_CARTON_NO(strDN As String)
Dim strSql      As String
Dim rs          As ADODB.Recordset
Dim strCartonID As String, strCartonQty As String
Dim id          As String
Dim strReelID   As String

strSql = "select REMARK2, SUM(QTY) from erptemp..TRAY_PSN_LIST where REMARK1 = '" & strDN & "' group by REMARK2"
Set rs = Get_SqlserveRs(strSql)
If rs.EOF Then
    MsgBox "查询不到该DN", vbInformation, "提示"
    Exit Sub

End If

rs.MoveFirst

Do While Not rs.EOF
    strCartonID = Trim$("" & rs(0))
    strCartonQty = Trim$("" & rs(1))
    ' ---------------------------------------------------删除
    '0
    strSql = "delete from [erpdata].[dbo].[tblPackTreeInf] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
    strSql = "delete from [erpdata].[dbo].[tblPackMainInf] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
    strSql = "update [erpdata].[dbo].[tblPackTreeInf] set 上级序号 = 0, Memo = '' where 箱号 in (select TRAY_ID from erptemp..TRAY_PSN_LIST where REMARK1 = '" & strDN & "' )  "
    'AddSql2 (strSql)
    strSql = "delete from [erpdata].[dbo].[tblStockNumTree] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strSql)
    strSql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号=0,Memo='', dn='' where 箱号 in (select TRAY_ID from erptemp..TRAY_PSN_LIST where REMARK1 = '" & strDN & "' ) "
    'AddSql2 (strSql)
    ' --------------------------------------------------更新
    '1 insert [erpdata].[dbo].[tblPackMainInf]
    strSql = "insert into [erpdata].[dbo].[tblPackMainInf](箱号,客户代码,数量,合格标记,装箱标记,产线标记) values('" & strCartonID & "','57'," & strCartonQty & ",'0','1','1')"
    If AddSql2(strSql) = 0 Then
        MsgBox "1 insert [erpdata].[dbo].[tblPackMainInf]:failed!!! ", vbCritical, "提示"
        Exit Sub

    End If

    '2 insert - update [erpdata].[dbo].[tblPackTreeInf]
    strSql = "insert into [erpdata].[dbo].[tblPackTreeInf](箱号,上级序号,基层标记 ,Memo) values('" & strCartonID & "',0,1,'57')"
    If AddSql2(strSql) = 0 Then
        MsgBox "2 insert [erpdata].[dbo].[tblPackTreeInf]:failed!!!", vbCritical, "提示"
        Exit Sub

    End If

    id = Get_SqlserverNo("select 序号 as ID from [erpdata].[dbo].[tblPackTreeInf] a where a.箱号='" & strCartonID & "' and Memo='57' ")
    strSql = "Update [erpdata].[dbo].[tblPackTreeInf] set 上级序号='" & id & "',Memo='57' " & " where 箱号 in (select TRAY_ID from erptemp..TRAY_PSN_LIST where REMARK1 = '" & strDN & "' ) "
    'If AddSql2(strSql) = 0 Then
    '    MsgBox "2 update [erpdata].[dbo].[tblPackTreeInf]:failed!!!", vbCritical, "提示"
    '    Exit Sub
    'End If
    '3 insert - update [erpdata].[dbo].[tblStockNumTree]
    strSql = "insert into [erpdata].[dbo].[tblStockNumTree](序号,箱号,上级序号,基层标记 ,尺寸,重量,Memo,DN) values( " & id & ",'" & strCartonID & "',0,1,'','','57','" & strDN & "')"
    If AddSql2(strSql) = 0 Then
        MsgBox "3 insert [erpdata].[dbo].[tblStockNumTree]", vbCritical, "提示"
        Exit Sub

    End If

    strSql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & id & "',Memo='57', dn='" & strDN & "' " & " where 箱号 in (select TRAY_ID from erptemp..TRAY_PSN_LIST where REMARK1 = '" & strDN & "' ) "
    'If AddSql2(strSql) = 0 Then
    '    MsgBox "3 update [erpdata].[dbo].[tblStockNumTree]", vbCritical, "提示"
    '    Exit Sub
    'End If
    rs.MoveNext
Loop

End Sub

Private Function GetQID(strReelID As String)
Dim strSql As String

strSql = "select trglabelseq.QTSeq_NotMesQbox('" & strReelID & "')  from dual"
GetQID = Get_OracleStr(strSql)

End Function
