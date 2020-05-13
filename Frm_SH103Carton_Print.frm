VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_SH103Carton_Print 
   Caption         =   "SH103外包"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13380
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
   ScaleHeight     =   11055
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "菜单"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13695
      Begin VB.CommandButton cmd 
         Caption         =   "补打标签"
         Height          =   360
         Left            =   10080
         TabIndex        =   14
         Top             =   1200
         Width           =   990
      End
      Begin VB.TextBox txtCartonNO 
         Height          =   285
         Left            =   8280
         TabIndex        =   12
         Top             =   1223
         Width           =   1695
      End
      Begin VB.CommandButton cmdERP 
         Caption         =   "更新ERP箱号"
         Height          =   480
         Left            =   11760
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtScanCode 
         Height          =   300
         Left            =   1680
         TabIndex        =   7
         Top             =   1215
         Visible         =   0   'False
         Width           =   5295
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   600
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   12945
         _ExtentX        =   22834
         _ExtentY        =   1058
         ButtonWidth     =   2408
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "开始扫描"
               Key             =   "SCAN"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "卷盘合箱"
               Key             =   "BIND"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印标签"
               Key             =   "PRINT"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "导出记录"
               Key             =   "EXPORT"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "EXIT"
               ImageIndex      =   5
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   12120
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_SH103Carton_Print.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_SH103Carton_Print.frx":0C52
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_SH103Carton_Print.frx":18A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_SH103Carton_Print.frx":24F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_SH103Carton_Print.frx":3148
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl222 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补打外箱箱号"
         Height          =   195
         Left            =   7200
         TabIndex        =   13
         Top             =   1268
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描卷盘ID:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1245
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "外箱明细"
      ForeColor       =   &H00FF0000&
      Height          =   11535
      Left            =   9360
      TabIndex        =   2
      Top             =   1680
      Width           =   4335
      Begin VB.TextBox txtWeight 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   10215
         Index           =   1
         Left            =   480
         TabIndex        =   10
         Top             =   600
         Width           =   3375
         _Version        =   524288
         _ExtentX        =   5953
         _ExtentY        =   18018
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
         SpreadDesigner  =   "Frm_SH103Carton_Print.frx":3D9A
         TextTip         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱重量(Kg)"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   285
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "卷盘明细"
      ForeColor       =   &H00FF0000&
      Height          =   11535
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   9135
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   10455
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   8775
         _Version        =   524288
         _ExtentX        =   15478
         _ExtentY        =   18441
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
         SpreadDesigner  =   "Frm_SH103Carton_Print.frx":4200
         TextTip         =   2
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   10680
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
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
         _cx             =   1720
         _cy             =   873
      End
   End
End
Attribute VB_Name = "Frm_SH103Carton_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim strShipID As String

Private Sub cmd_Click()
Call printOuterLbl2
End Sub

Private Sub cmdERP_Click()
Dim strsql      As String
Dim rs          As ADODB.Recordset
Dim strCartonID As String, strCartonQty As String
Dim id          As String

On Error GoTo ERRON

INIadoCon.BeginTrans
strsql = "select CARTON_NO, SUM(QTY) from packing_detailed_sh103 where print_flag = '1' group by CARTON_NO"
Set rs = Get_OracleRs(strsql)
If rs.EOF Then
    MsgBox "查询不到数据", vbInformation, "提示"
    INIadoCon.RollbackTrans
    Exit Sub

End If

rs.MoveFirst

Do While Not rs.EOF
    strCartonID = Trim$("" & rs(0))
    strCartonQty = Trim$("" & rs(1))
    ' ---------------------------------------------------删除
    '0
    strsql = "delete from [erpdata].[dbo].[tblPackTreeInf] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strsql)
    strsql = "delete from [erpdata].[dbo].[tblPackMainInf] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strsql)
    strsql = "delete from [erpdata].[dbo].[tblStockNumTree] where 箱号 = '" & strCartonID & "'"
    AddSql2 (strsql)
    ' --------------------------------------------------更新
    '1 insert [erpdata].[dbo].[tblPackMainInf]
    strsql = "insert into [erpdata].[dbo].[tblPackMainInf](箱号,客户代码,数量,合格标记,装箱标记,产线标记) values('" & strCartonID & "','SH103'," & strCartonQty & ",'0','1','1')"
    AddSql2 (strsql)
    '2 insert - update [erpdata].[dbo].[tblPackTreeInf]
    strsql = "insert into [erpdata].[dbo].[tblPackTreeInf](箱号,上级序号,基层标记 ,Memo) values('" & strCartonID & "',0,1,'SH103')"
    AddSql2 (strsql)
    id = Get_SqlserverNo("select 序号 as ID from [erpdata].[dbo].[tblPackTreeInf] a where a.箱号='" & strCartonID & "' and Memo='SH103' ")
    strsql = "Update [erpdata].[dbo].[tblPackTreeInf] set 上级序号='" & id & "',Memo='SH103' " & " where 箱号 in (select REEL_CODE from OPENQUERY(ORACLEDB, 'SELECT * from  packing_detailed_sh103' ) X where X.CARTON_NO = '" & strCartonID & "') "
    AddSql2 (strsql)
    '3 insert - update [erpdata].[dbo].[tblStockNumTree]
    strsql = "insert into [erpdata].[dbo].[tblStockNumTree](序号,箱号,上级序号,基层标记 ,尺寸,重量,Memo,DN) values( " & id & ",'" & strCartonID & "',0,1,'','','SH103','')"
    AddSql2 (strsql)
    strsql = "Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & id & "',Memo='SH103' where 箱号 in (select REEL_CODE from OPENQUERY(ORACLEDB, 'SELECT * from  packing_detailed_sh103' ) X where X.CARTON_NO = '" & strCartonID & "') "
    AddSql2 (strsql)
    rs.MoveNext
Loop
INIadoCon.CommitTrans
MsgBox "箱号已更新", vbInformation, "提示"
Exit Sub
ERRON:
INIadoCon.RollbackTrans
MsgBox "错误:" & Err.DESCRIPTION, vbCritical, "警告"

End Sub

Private Sub Form_Load()
InitCtrls

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "SCAN"
        txtScanCode.Visible = True
        txtScanCode.SetFocus
        Fps(0).MaxRows = 0
        Fps(1).MaxRows = 0
        Play ("请扫描卷盘条码")
        strShipID = Get_OracleStr("select SH103_SHIP_SEQ.NEXTVAL from dual")
    Case "BIND"
        bindReels

    Case "PRINT"
        printOuterLbl
        
    Case "EXPORT"
        exportHistory

    Case "EXIT"
        Unload Me

End Select

End Sub

Private Sub exportHistory()

Dim strsql As String

strsql = "select * from PACKING_DETAILED_SH103 where PRINT_FLAG = '1' order by DN,CARTON_NO,REEL_CODE"
ExporToExcel (strsql)

End Sub

Private Sub InitCtrls()

With Fps(0)
    .Col = -1
    .Row = -1
    .Lock = True
    .SetText 0, 0, "序号"
    .SetText 1, 0, "ReelCode"
    .SetText 2, 0, "ProdName"
    .SetText 3, 0, "LotNo"
    .SetText 4, 0, "Qty"
    .SetText 5, 0, "H"
    .SetText 6, 0, "D/C"
    .SetText 7, 0, "外箱箱号"
    .SetText 8, 0, "厂内机种"
    .ColWidth(1) = 15
    .ColWidth(2) = 15
    .ColWidth(3) = 10
    .ColWidth(4) = 6
    .ColWidth(5) = 2
    .ColWidth(6) = 5
    .ColWidth(7) = 10
    .ColWidth(8) = 8

End With

With Fps(1)
    .Col = -1
    .Row = -1
    .Lock = True
    .SetText 0, 0, "序号"
    .SetText 1, 0, "外箱箱号"
    .SetText 2, 0, "外箱重量(Kg)"
    .ColWidth(1) = 10
    .ColWidth(2) = 10

End With

End Sub

Private Sub txtScanCode_KeyPress(KeyAscii As Integer)

Dim strScanCode As String

If KeyAscii <> vbKeyReturn Or Len(Trim(txtScanCode.text)) = 0 Then Exit Sub
strScanCode = UCase(Trim$(txtScanCode.text))

If strScanCode = "000000" Then
    Call bindReels
Else
    Call getReels(strScanCode)

End If

txtScanCode.text = ""

End Sub

Private Sub getReels(strReelCode As String)
Dim strsql     As String
Dim strContent As String
Dim strSp
Dim strData   As SH103_REEL_INFO
Dim i         As Integer
Dim iMaxReels As Integer

iMaxReels = 0
strsql = "select * from PACKING_DETAILED_SH103 where REEL_CODE = '" & strReelCode & "' and PRINT_FLAG = '1' "
If Get_OracleCnt(strsql) > 0 Then
    MsgBox "该卷盘：" & strReelCode & "有打印历史记录,本次扫描出错", vbCritical, "警告"
    Exit Sub

End If

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 7
        If .text = "" Then
            iMaxReels = iMaxReels + 1

        End If

    Next
    If iMaxReels = 18 Then
        MsgBox "已满足外箱最大卷盘数:18" & vbCrLf & "请维护称重并点击合箱", vbInformation, "提示"
        Exit Sub

    End If

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .text = strReelCode Then
            Play ("请勿重复扫描卷盘号")
            Exit Sub


        End If

    Next

End With

strsql = "select top 1 Content from erpdata..tblME_PrintInfo aa , erpdata..tblErpInStockDetailInfo bb where bb.KEY_VALUE = '" & strReelCode & "' and bb.keyid = aa.EVENT_ID and bb.KEY_NAME = 'CONTAINER_NAME'  and aa.BartenderName = 'SH103IN1.btw' and aa.PrinterNameID = 'W_IN_2B5F_2' order by ID desc"
strContent = Get_SqlStr(strsql)
If strContent = "" Then
    MsgBox "查询不到该ReelCode, 请确认是否为实物标签,或信息有误", vbInformation, "提示"
    Exit Sub

End If

strSp = Split(strContent, ";")
If Split(Replace(strSp(3), """", ""), ",")(1) = strReelCode Then
    strData.SH103_REEL_CODE = Split(Replace(strSp(3), """", ""), ",")(1)
    strData.Customer_Device = Split(Replace(strSp(0), """", ""), ",")(1)
    strData.IN_GOOD_DIE = Split(Replace(strSp(1), """", ""), ",")(1)
    strData.PACKING_DATE_10 = Split(Replace(strSp(2), """", ""), ",")(1)
    strData.SH103_H = Split(Replace(strSp(4), """", ""), ",")(1)
    strData.SH103_LOT_NO = Split(Replace(strSp(5), """", ""), ",")(1)
    strData.HT_PN = Split(Replace(strSp(6), """", ""), ",")(1)
Else
    strData.SH103_REEL_CODE = Split(Replace(strSp(0), """", ""), ",")(1)
    strData.Customer_Device = Split(Replace(strSp(1), """", ""), ",")(1)
    strData.IN_GOOD_DIE = Split(Replace(strSp(2), """", ""), ",")(1)
    strData.PACKING_DATE_10 = Split(Replace(strSp(3), """", ""), ",")(1)
    strData.SH103_H = Split(Replace(strSp(4), """", ""), ",")(1)
    strData.SH103_LOT_NO = Split(Replace(strSp(5), """", ""), ",")(1)
    strData.HT_PN = Split(Replace(strSp(6), """", ""), ",")(1)

End If

Call listReelsFps(strData)
If iMaxReels = 7 Then
    Play ("已满足最大包装量, 请维护称重并点击合箱")
    txtWeight.text = "1.82"
    Call bindReels

End If

End Sub

Private Sub listReelsFps(strData As SH103_REEL_INFO)

Dim i As Integer

With Fps(0)
    .MaxRows = .MaxRows + 1
    i = .MaxRows
    .SetText 1, i, strData.SH103_REEL_CODE
    .SetText 2, i, strData.Customer_Device
    .SetText 3, i, strData.SH103_LOT_NO
    .SetText 4, i, strData.IN_GOOD_DIE
    .SetText 5, i, strData.SH103_H
    .SetText 6, i, strData.PACKING_DATE_10
    .SetText 8, i, strData.HT_PN

End With

Play ("卷盘已扫描")

End Sub

Private Sub bindReels()

Dim strCartonNO     As String
Dim strCartonSeq    As String
Dim strReelCode     As String
Dim strCartonWeight As String
Dim strsql          As String
Dim i               As Integer
Dim strData         As SH103_QBOX_INFO
Dim bBindReady      As Boolean

bBindReady = False

If Fps(0).MaxRows = 0 Then
    MsgBox "没有卷盘扫描,不可以合箱", vbInformation, "提示"
    Exit Sub

End If

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 7

        If .text = "" Then
            bBindReady = True

        End If

    Next

End With

If Not bBindReady Then
    MsgBox "没有待合箱的卷盘,请扫描需要合箱的卷盘", vbInformation, "提示"
    Exit Sub

End If

If txtWeight.text = "" Then
    MsgBox "请输入该外箱的重量", vbInformation, "提示"
    Exit Sub

End If

With Fps(0)
    .Row = .MaxRows
    .Col = 1
    strReelCode = Trim(.text)

End With

' 产生外箱号
strData.SH103_QBOXID = Get_OracleStr("select trglabelseq.QTSeq_NotMesQbox('" & strReelCode & "')  from dual")
strData.SH103_QBOXWEIGHT = Trim(txtWeight.text)
strData.SH103_QBOXSEQ = Fps(1).MaxRows + 1
Call listQBoxsFps(strData)
Call updateReelsFps(strData)

txtWeight.text = ""

End Sub

Private Sub saveSql()
    Dim strsql  As String
    Dim i       As Integer
    Dim j       As Integer
    Dim strData As SH103_REEL_INFO

    With Fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            strData.SH103_REEL_CODE = .text
            .Col = 2
            strData.Customer_Device = .text
            .Col = 3
            strData.SH103_LOT_NO = .text
            .Col = 4
            strData.IN_GOOD_DIE = .text
            .Col = 5
            strData.SH103_H = .text
            .Col = 6
            strData.PACKING_DATE_10 = .text
            .Col = 7
            strData.SH103_QBOXID = .text
            .Col = 8
            strData.HT_PN = .text
            
            With Fps(1)

                For j = 1 To .MaxRows
                    .Row = j
                    .Col = 1

                    If .text = strData.SH103_QBOXID Then
                        .Col = 0
                        strData.SH103_QBOXSEQ = .text
                    
                        .Col = 2
                        strData.SH103_QBOXWEIGHT = .text

                    End If

                Next j

            End With
            
            strsql = "insert into PACKING_DETAILED_SH103(DN,REEL_CODE,LOT_NO,PROD_NAME,H,QTY,DC,CARTON_NO,CARTON_SEQ,CARTON_WEIGHT,PRINT_FLAG,CREATE_DATE,CREATE_BY,HTPN) " & " values('" & strShipID & "', '" & strData.SH103_REEL_CODE & "', '" & strData.SH103_LOT_NO & "','" & strData.Customer_Device & "','" & strData.SH103_H & "','" & strData.IN_GOOD_DIE & "', '" & strData.PACKING_DATE_10 & "','" & strData.SH103_QBOXID & "' ,'" & strData.SH103_QBOXSEQ & "','" & strData.SH103_QBOXWEIGHT & "' ,'0',sysdate,'" & gUserName & "','" & strData.HT_PN & "')"
            AddSql (strsql)

        Next i

    End With

  

End Sub

Private Sub listQBoxsFps(strData As SH103_QBOX_INFO)

Dim i As Integer

With Fps(1)
    .MaxRows = .MaxRows + 1
    i = .MaxRows
    .SetText 1, i, strData.SH103_QBOXID
    .SetText 2, i, strData.SH103_QBOXWEIGHT

End With

Play ("外箱已生成")

End Sub

Private Sub updateReelsFps(strData As SH103_QBOX_INFO)

Dim i As Integer

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 7

        If .text = "" Then
            .text = strData.SH103_QBOXID

        End If

    Next

End With

End Sub

Private Sub printOuterLbl()

Dim strCartonNO     As String
Dim strCartonNo2    As String
Dim strDate         As String
Dim strCartonQty    As String
Dim strCartonSeq    As String
Dim strCartonTotal  As String
Dim strCartonWeight As String
Dim strH            As String
Dim i               As Integer, j As Integer
Dim strCartonCode   As String
Dim strReelCode     As String
Dim strQrCode       As String
Dim strLblCode      As String
Dim rs              As New ADODB.Recordset
Dim sFileName       As String
Dim sFilePath       As String
Dim lReelQty        As Long
Dim strData         As SH103_REEL_INFO
Dim strsql          As String
Dim strHTPN As String


strLblCode = ""
strReelCode = ""
strQrCode = ""
strCartonCode = ""
lReelQty = 0

If MsgBox("请确认此批出货是否已全部扫描完成,是否开始打印标签", vbYesNo, "提示") = vbNo Then
    Exit Sub

End If

If Fps(1).MaxRows = 0 Then
    MsgBox "外箱没有数据,不可打印", vbInformation, "提示"
    Exit Sub
Else

    With Fps(0)
        .Row = .MaxRows
        .Col = 7

        If .text = "" Then
            MsgBox "最后一个外箱没有合箱,请点击合箱", vbInformation, "警告"
            Exit Sub

        End If

    End With

End If

Call saveSql

strCartonTotal = Right("000" & Fps(1).MaxRows, 3)

With Fps(1)

    For i = 1 To .MaxRows
        .Row = i
        strCartonSeq = Right("000" & .Row, 3)
        .Col = 1
        strCartonNO = Trim$(.text)
        .Col = 2
        strCartonWeight = Trim$(.text)
        
        strCartonQty = getCartonQty(strCartonNO)
        strH = getCartonH(strCartonNO)
        strCartonNo2 = Right(Year(Now), 2) & Right(("0" & Month(Now)), 2) & Right(("0" & Day(Now)), 2) & "03" & Right$(strCartonNO, 4)
        strDate = Right(("0" & Month(Now)), 2) & "/" & Right(("0" & Day(Now)), 2) & "/" & Year(Now)
        ' 外箱信息
        strCartonCode = strCartonNo2 & ";" & strDate & ";" & strCartonNO & ";" & strCartonQty & ";" & strCartonSeq & ";" & strCartonTotal & ";" & strCartonWeight & ";" & strH & ";"
  
        strsql = "select REEL_CODE,PROD_NAME,LOT_NO,QTY,H,DC,HTPN from PACKING_DETAILED_SH103 where dn = '" & strShipID & "' and CARTON_NO = '" & strCartonNO & "' order by REEL_CODE"
        Set rs = Get_OracleRs(strsql)
        
        If Not rs.EOF Then

            Do While Not rs.EOF
                lReelQty = lReelQty + 1
                strData.SH103_REEL_CODE = rs(0).Value
                strData.Customer_Device = rs(1).Value
                strData.SH103_LOT_NO = rs(2).Value
                strData.IN_GOOD_DIE = rs(3).Value
                strData.SH103_H = rs(4).Value
                strData.PACKING_DATE_10 = rs(5).Value
                strHTPN = rs(6).Value
                
                ' 卷盘信息
                strReelCode = strReelCode & strData.Customer_Device & ";" & strData.SH103_REEL_CODE & ";" & strData.PACKING_DATE_10 & ";" & strData.IN_GOOD_DIE & ";" & strData.SH103_LOT_NO & ";"
                ' 二维码信息
                strQrCode = strQrCode & strData.Customer_Device & "," & strData.SH103_REEL_CODE & ","
                
                rs.MoveNext
            Loop
            
        End If
        
        If lReelQty < 18 Then

            For k = 1 To (18 - lReelQty)
                strReelCode = strReelCode & ";" & ";" & ";" & ";" & ";"
            Next

        End If
        
        strQrCode = Left(strQrCode, Len(strQrCode) - 1)
        
        ' 标签总信息
        strLblCode = strCartonCode & strReelCode & strQrCode & ";" & strHTPN
'        strLblCode = strCartonCode & strReelCode & strQrCode
        ' 打印该标签
        sFileName = "SH103_LBL" & "_" & strCartonNO & "_" & Format(Now(), "YYYYMMDDHHmmSS")

        If gUserName = "07885" Then
            sFilePath = "C:\test\"
        Else
            sFilePath = "\\10.160.1.84\public\BarCode\SH103\OUTPKG\"

        End If

        Call CreateTxt(sFileName, strLblCode, sFilePath)
        
        ' 更新打印状态
        strsql = "update PACKING_DETAILED_SH103 set PRINT_FLAG = '1',CARTON_SEQ = '" & strCartonSeq & "',CARTON_NO2 = '" & strCartonNo2 & "',LBLDATE = '" & strDate & "',CARTON_TOTAL = '" & strCartonTotal & "'    where dn = '" & strShipID & "' and CARTON_NO = '" & strCartonNO & "' "
        AddSql (strsql)
        
        strCartonCode = ""
        strReelCode = ""
        strQrCode = ""
        lReelQty = 0
        Sleep (2000)
    Next

End With

MsgBox "外箱标签已全部打印完成,本次出货虚拟ID为：" & strShipID, vbInformation, "提示"

End Sub

'补打外箱标签
Private Sub printOuterLbl2()
Dim strCartonNO   As String
Dim strsql        As String
Dim rs            As New ADODB.Recordset
Dim strCartonCode As String
Dim strReelCode   As String
Dim strQrCode     As String
Dim strLblCode    As String
Dim sFileName     As String
Dim sFilePath     As String
Dim lReelQty      As Long
Dim strData       As SH103_REEL_INFO
Dim strHTPN As String
strCartonNO = UCase(Trim$(txtCartonNO.text))

If strCartonNO = "" Then
    MsgBox "请输入需要补打的外箱箱号Q", vbCritical, "提示"
    Exit Sub

End If

strsql = "select CARTON_NO2 || ';'|| LBLDATE || ';' || CARTON_NO || ';' || sum(QTY) || ';' || CARTON_SEQ || ';' || CARTON_TOTAL || ';' || to_char(CARTON_WEIGHT,'fm999999990.999999999') || ';' || H || ';' from PACKING_DETAILED_SH103 where carton_no = '" & strCartonNO & "' and print_flag = '1' " & " group by CARTON_NO2,LBLDATE,CARTON_NO,CARTON_SEQ,CARTON_WEIGHT,CARTON_TOTAL ,H "

'外箱信息
strCartonCode = Get_OracleStr(strsql)

If strCartonCode = "" Then
    MsgBox "找不到该外箱箱号, 请确定是否有该外箱号的补打记录", vbInformation, "提示"
    Exit Sub

End If

strsql = "select REEL_CODE,PROD_NAME,LOT_NO,QTY,H,DC,HTPN from PACKING_DETAILED_SH103 where CARTON_NO = '" & strCartonNO & "' and print_flag = '1' order by REEL_CODE"
Set rs = Get_OracleRs(strsql)
        
If Not rs.EOF Then

    Do While Not rs.EOF
        lReelQty = lReelQty + 1
        strData.SH103_REEL_CODE = rs(0).Value
        strData.Customer_Device = rs(1).Value
        strData.SH103_LOT_NO = rs(2).Value
        strData.IN_GOOD_DIE = rs(3).Value
        strData.SH103_H = rs(4).Value
        strData.PACKING_DATE_10 = rs(5).Value
        strHTPN = rs(6).Value
                
        ' 卷盘信息
        strReelCode = strReelCode & strData.Customer_Device & ";" & strData.SH103_REEL_CODE & ";" & strData.PACKING_DATE_10 & ";" & strData.IN_GOOD_DIE & ";" & strData.SH103_LOT_NO & ";"
        ' 二维码信息
        strQrCode = strQrCode & strData.Customer_Device & "," & strData.SH103_REEL_CODE & ","
                
        rs.MoveNext
    Loop
            
End If
        
If lReelQty < 18 Then

    For k = 1 To (18 - lReelQty)
        strReelCode = strReelCode & ";" & ";" & ";" & ";" & ";"
    Next

End If
        
strQrCode = Left(strQrCode, Len(strQrCode) - 1)
        
' 标签总信息
strLblCode = strCartonCode & strReelCode & strQrCode & ";" & strHTPN

' 打印该标签
sFileName = "SH103_LBL" & "_" & strCartonNO & "_" & Format(Now(), "YYYYMMDDHHmmSS")

If gUserName = "07885" Then
    sFilePath = "C:\test\"
Else
    sFilePath = "\\10.160.1.84\public\BarCode\SH103\OUTPKG\"

End If

Call CreateTxt(sFileName, strLblCode, sFilePath)
        
MsgBox "箱号:" & strCartonNO & " 补打完成", vbInformation, "提示"

End Sub

Private Function getCartonQty(strCartonID As String) As Long

Dim lQty As Long
Dim i    As Integer

lQty = 0

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 7

        If .text = strCartonID Then
            .Col = 4
            lQty = lQty + CLng(.text)

        End If

    Next

End With

getCartonQty = lQty

End Function

Private Function getCartonH(strCartonID As String) As String

Dim strH As String
Dim i    As Integer

lQty = 0

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 7

        If .text = strCartonID Then
            .Col = 5
            strH = .text

        End If

    Next

End With

getCartonH = strH

End Function

Private Sub Play(sFileName As String)

Dim sPath   As String
Dim sSuffix As String

sPath = "\\10.160.1.84\public\media_source\"
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix

End Sub

Private Sub CreateTxt(filename As String, msgTxt As String, dirtemp As String)

Dim fileNameTemp As String
Dim dirNameTemp  As String
Dim fileTemp     As String

dirNameTemp = dirtemp
fileNameTemp = Replace(filename, "'", "") & ".txt"
fileTemp = dirNameTemp & fileNameTemp
Open fileTemp For Output As #1
Print #1, msgTxt
Close #1

'Sleep (1000)
End Sub

