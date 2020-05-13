VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmSemtech_Report 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Semtech报表查询"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14445
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
   ScaleHeight     =   9105
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHBDC 
      BackColor       =   &H0000FF00&
      Caption         =   "合并导出报表"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5160
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdDXDC 
      BackColor       =   &H000080FF&
      Caption         =   "单项导出报表"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3360
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Fra 
      ForeColor       =   &H000000FF&
      Height          =   7335
      Index           =   1
      Left            =   3600
      TabIndex        =   17
      Top             =   960
      Width           =   9615
      Begin VB.CheckBox chooseALL 
         Caption         =   "全选/反选"
         Height          =   315
         Left            =   120
         MaskColor       =   &H0000FF00&
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   3375
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   6615
         _Version        =   524288
         _ExtentX        =   11668
         _ExtentY        =   5953
         _StockProps     =   64
         EditEnterAction =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   5
         SpreadDesigner  =   "FrmSemtech_Report.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "查询条件"
      ForeColor       =   &H00FF0000&
      Height          =   13815
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "重置"
         Height          =   360
         Left            =   2640
         TabIndex        =   27
         Top             =   3720
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "导出"
         Height          =   360
         Left            =   2640
         TabIndex        =   26
         Top             =   3240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtLotID 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "检索"
         Height          =   285
         Left            =   2640
         TabIndex        =   24
         Top             =   2880
         Width           =   735
      End
      Begin VB.ListBox lstLotID 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   10320
         Left            =   840
         Style           =   1  'Checkbox
         TabIndex        =   23
         Top             =   3240
         Width           =   1695
      End
      Begin VB.ComboBox cmbCombo1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         ItemData        =   "FrmSemtech_Report.frx":04E1
         Left            =   1080
         List            =   "FrmSemtech_Report.frx":04E8
         TabIndex        =   19
         Top             =   2400
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   0
         Left            =   1080
         TabIndex        =   15
         Top             =   1080
         Width           =   2355
      End
      Begin VB.ComboBox cmbCombo1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         ItemData        =   "FrmSemtech_Report.frx":04F7
         Left            =   1080
         List            =   "FrmSemtech_Report.frx":04F9
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   109576195
         CurrentDate     =   41387
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   109576195
         CurrentDate     =   41387
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D N"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   22
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   2460
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发货单号"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据类型"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期起"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期末"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job      No"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   720
      End
   End
   Begin VB.Frame Fra 
      Height          =   855
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.CommandButton cmdInvRpt 
         Caption         =   "一键导出库存报表"
         Height          =   360
         Left            =   9240
         MaskColor       =   &H8000000F&
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin MSComDlg.CommonDialog Com 
         Left            =   11760
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "上传文件"
         Height          =   480
         Left            =   11160
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "导出报表"
         Enabled         =   0   'False
         Height          =   360
         Left            =   6720
         MaskColor       =   &H0000FFFF&
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退 出"
         Height          =   360
         Left            =   8040
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdExprot 
         Caption         =   "导出当前数据"
         Enabled         =   0   'False
         Height          =   360
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "查  询"
         Height          =   360
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   990
      End
   End
End
Attribute VB_Name = "FrmSemtech_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim strdjbh         As String
Dim strdjbh1         As String
Dim DirShare        As String
Dim DirFileShare    As String
Dim DirQrShare As String

Dim DirInvRpt       As String
Dim order           As String
Dim RsClone         As New ADODB.Recordset
Const C_Left = 60
Const C_Top = 120

Private Enum fpSDetail
    E_CHOOSE = 1
    e_DJBH = 2
    E_cust = 3
    e_YDH = 7
End Enum
'更新种子值
Private Function GetExcelName(ByVal strTitle As String) As String
Dim strSql          As String
Dim rs              As New ADODB.Recordset
Dim strExFileName   As String
Dim strCurDate      As String
    
    If strTitle = "output Invoice" Then
        strCurDate = Format(DATE, "YY-MMDD")
    Else
        strCurDate = Format(DATE, "YYYYMMDD")
    End If
    strSql = "select nvl(max(para_3),0)+1 para from tblsys_parameter where sysname='TSVSYS' and kind='Semtech报表' and para_1='" & strTitle & "' and para_2='" & strCurDate & "'"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        strExFileName = strCurDate + "_" + Trim$("" & rs!Para)
        strSql = "Update tblsys_parameter set para_2='" & strCurDate & "',para_3='" & Trim$("" & rs!Para) & "' where sysname='TSVSYS' and kind='Semtech报表' and para_1='" & strTitle & "'"
        Cnn.Execute strSql
    Else
        strExFileName = strCurDate + "_1"
        strSql = "Update tblsys_parameter set para_2='" & strCurDate & "',para_3='" & Trim$("" & rs!Para) & "' where sysname='TSVSYS' and kind='Semtech报表' and para_1='" & strTitle & "'"
        Cnn.Execute strSql
    End If
    rs.Close
    
    If strTitle = "output Invoice" Then
        GetExcelName = "HTKS_SEDC" & strExFileName
    Else
        GetExcelName = strTitle & "_" & strExFileName
    End If
    
End Function
'全选和反向全选
Private Sub chooseALL_Click()

Dim i As Integer

If chooseALL.Value = 1 Then

    For i = 1 To Fps(0).MaxRows

        With Fps(0)
            .Row = i
            .Col = 1
            .text = 1

        End With

    Next i

ElseIf chooseALL.Value = 0 Then

    For i = 1 To Fps(0).MaxRows

        With Fps(0)
            .Row = i
            .Col = 1
            .text = 0

        End With

    Next i

End If
End Sub

Private Sub cmbCombo1_Click(Index As Integer)
'Dim strSql              As String
'Dim Rs                  As New ADODB.Recordset


Fps(0).MaxRows = 0

    If Index = 0 Then
        If cmbCombo1(0).text = "库存报表" Then
            '加载仓库
            cmbCombo1(1).Clear
            cmbCombo1(1).AddItem "所有"
            cmbCombo1(1).AddItem "1000"
            cmbCombo1(1).AddItem "6000"
            cmbCombo1(1).AddItem "7000"
            cmbCombo1(1).AddItem "8000"
            cmbCombo1(1).AddItem "9000"
            cmbCombo1(1).AddItem "Scrap"
            cmbCombo1(1).ListIndex = 0
        Else
            cmbCombo1(1).Clear
        End If
    End If
End Sub


'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cmdDXDC_Click
' Description:       DN地址合并导出功能
' Created by :       Project Administrator
' Machine    :       DESKTOP-F6L8S2V
' Date-Time  :       2019/10/30-14:38:24
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdHBDC_Click()
    Dim strExportName           As String
    
    cmdReport.Enabled = False
    If Fps(0).MaxRows <= 0 Then
        MsgBox "没有可导出的数据！", vbInformation, "提示"
        Exit Sub
    End If
    '导出报表
    strExportName = GetExcelName(Trim(Fra(1).Caption))
    
    If cmbCombo1(0).ListIndex = 0 Then
'        Call SEDCExportPrintExcel(RsClone, strExportName)                     'SEDC报表
    ElseIf cmbCombo1(0).ListIndex = 1 Then
'        Call InputPackinglistExportPrintExcel(RsClone, strExportName)         'outputPackinglist报表
    ElseIf cmbCombo1(0).ListIndex = 2 Then
'        Call InputInvoiceExportPrintExcel(RsClone, strExportName)             'outputInvoice报表
    ElseIf cmbCombo1(0).ListIndex = 3 Then
'        Call Daily_InvExportPrintExcel(RsClone, strExportName)                'Daily_inventory_report
    ElseIf cmbCombo1(0).ListIndex = 4 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel2(strExportName, 0)      'Shipping Packinglist
    ElseIf cmbCombo1(0).ListIndex = 5 Then
        If Not CheckData Then Exit Sub
        Call ShippingInvoiceExportPrintExcel2            'Shipping invoice
    ElseIf cmbCombo1(0).ListIndex = 9 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel2(strExportName, 1)
    End If
    
    cmdReport.Enabled = True
    
End Sub


'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       cmdZXDC_Click
' Description:       单项导出功能
' Created by :       祝祎凡
' Machine    :       DESKTOP-F6L8S2V
' Date-Time  :       2019/10/30-14:37:45
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdDXDC_Click()
    Dim strExportName           As String

    cmdReport.Enabled = False
    If Fps(0).MaxRows <= 0 Then
        MsgBox "没有可导出的数据！", vbInformation, "提示"
        Exit Sub
    End If
    '导出报表
    strExportName = GetExcelName(Trim(Fra(1).Caption))
    
    If cmbCombo1(0).ListIndex = 0 Then
'        Call SEDCExportPrintExcel(RsClone, strExportName)                     'SEDC报表
    ElseIf cmbCombo1(0).ListIndex = 1 Then
'        Call InputPackinglistExportPrintExcel(RsClone, strExportName)         'outputPackinglist报表
    ElseIf cmbCombo1(0).ListIndex = 2 Then
'        Call InputInvoiceExportPrintExcel(RsClone, strExportName)             'outputInvoice报表
    ElseIf cmbCombo1(0).ListIndex = 3 Then
'        Call Daily_InvExportPrintExcel(RsClone, strExportName)                'Daily_inventory_report
    ElseIf cmbCombo1(0).ListIndex = 4 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel1(strExportName, 0)    'Shipping Packinglist
    ElseIf cmbCombo1(0).ListIndex = 5 Then
        If Not CheckData Then Exit Sub
'        Call ShippingInvoiceExportPrintExcel(order, strExportName)          'Shipping invoice
    ElseIf cmbCombo1(0).ListIndex = 9 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel1(strExportName, 1)
    End If
    
    cmdReport.Enabled = True
    
'Shell (App.Path & "\install.bat")
End Sub

Private Sub cmdExit_Click() '退出
    Unload Me
End Sub

Private Sub cmdExprot_Click()
Dim strExportName           As String

cmdExprot.Enabled = False
    '校验数据
    If Fps(0).MaxRows <= 0 Then
        MsgBox "没有可导出的数据！", vbInformation, "提示"
        Exit Sub
    End If
    '导出报表
    If cmbCombo1(0).ListIndex = 6 Or cmbCombo1(0).ListIndex = 7 Or cmbCombo1(0).ListIndex = 8 Then '库存报表,SMTCList,Shipped
        strExportName = Trim(Fra(1).Caption)
        If cmbCombo1(1).text <> "" Then
            strExportName = Trim(Fra(1).Caption) + "-" + Trim(cmbCombo1(1).text)
        End If
    Else
        strExportName = GetExcelName(Trim(Fra(1).Caption))
    End If

    If Not ExportFpspreadToExcel(Fps(0), strExportName, strExportName) Then Exit Sub
    
End Sub

Private Sub cmdInvRpt_Click()
'一键导出库存报表到指定文件夹中
Dim strSql                  As String
Dim strSqlDetail            As String
Dim rs                      As New ADODB.Recordset
Dim i                       As Integer
Dim strFileName             As String
Dim strmsg                  As String
    
    If MsgBox("确定要导出吗?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
        Exit Sub
    End If
    '导出的是库存
    strSql = "SELECT RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Reel_Size,Remark" & _
             ",Move_in_Date FROM Vw_InvStockRptFor37 Where 2>1 "
    For i = 0 To cmbCombo1(1).ListCount
        If cmbCombo1(1).List(i) <> "所有" And cmbCombo1(1).List(i) <> "" Then
            If cmbCombo1(1).List(i) = "9000" Then
               
'            strSql = "SELECT b.fab_conv_id as [Wafer Type],b.mpn_desc as [Assy Part#],a.工单号 as [Fab Lot], " & _
'            " right(replace(rtrim(a.流程卡编号),'+',''),2) ID,b.date_code as [D/C],b.die as [QTY(称重后的数量)]," & _
'            " b.test_mtrl_desc as [Job#],b.bag as Bag#,a.箱号 as Comment,b.alternatename as [HT Part#]" & _
'            " FROM erpdata..tblStockNumsub a," & _
'            " (SELECT  substring(cast(datepart(year,ora.create_date)as nvarchar(20))+substring(CAST ('100'+datepart(week,ora.create_date) as nvarchar(20)),2,2),3,4) as date_code," & _
'            " ora.waferid,ora.firstname ,ora.test_mtrl_desc,ora.mpn_desc,ora.bag,ora.fab_conv_id,ora.die,ora.alternatename FROM " & _
'            " OPENQUERY(ORACLEDB, 'SELECT d.fab_conv_id,a.waferid,a.die," & _
'            " a.create_date+6 as create_date,b.firstname ,d.test_mtrl_desc,d.mpn_desc,get_37bagid(b.containername) bag,e.alternatename" & _
'            " FROM weight37 a,container b ,mappingdatatest c,customeroitbl_test d,product e" & _
'            " where a.waferid||''-A'' = b.containername  and a.waferid = c.substrateid " & _
'            " and e.productid=b.productid and c.filename = d.id ' ) ora ) b where a.库房编号 IN('44','45')" & _
'            " and a.流程卡编号=b.waferid "
            strSql = " SELECT * FROM DBO.Vw_InvStockRptFor37By9000 "
            strSqlDetail = ""
            Else
              strSql = "SELECT RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Reel_Size,Remark" & _
              ",Move_in_Date FROM Vw_InvStockRptFor37 Where 2>1 "
              strSqlDetail = " And 仓库名称='" & cmbCombo1(1).List(i) & "'"
            End If
            If rs.State = adStateOpen Then rs.Close
            rs.Open strSql + strSqlDetail, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs.EOF Then '表示有数据才导出报表
                strFileName = Format(Now(), "YYMMDD_HHMM") + "_" + Replace(cmbCombo1(1).List(i), "Scrap", "SCR")
                strmsg = strmsg + DirInvRpt + "\" + strFileName + vbCrLf                '提示消息
                RsExporToExcel rs, cmbCombo1(1).List(i), strFileName                    '导出到Excel
            End If
            rs.Close
        End If
    Next
    '导出Shipped
    strSql = "SELECT RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Remark,SO as [DN#],SHIPPED_DATE,Cust_Name " & _
             " FROM Vw_InvShippedRptFor37 "
'             " WHERE SHIPPED_DATE>='" & DateAdd("m", -1, Format(Now(), "YYYY-MM-DD")) & "' and SHIPPED_DATE<'" & DateAdd("d", 1, Format(Now(), "YYYY-MM-DD")) & "' "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then '表示有数据才导出报表
        strFileName = Format(Now(), "YYMMDD_HHMM") + "_SHIPPED"
        strmsg = strmsg + DirInvRpt + "\" + strFileName + vbCrLf                '提示消息
        RsExporToExcel rs, "SHIPPED", strFileName                               '导出到Excel
    End If
    rs.Close
    
    MsgBox "导出成功，导出文件路径为：" + vbCrLf + strmsg
    
End Sub

Private Sub cmdquery_Click()
Dim strKey As String
Dim i      As Integer
Dim bRet   As Boolean

bRet = False
strKey = Trim$(txtLotID.text)
If strKey = "" Then
    MsgBox "请输入DN", vbInformation, "提示:"
    Exit Sub

End If

With lstLotID

    For i = 0 To .ListCount - 1
        If strKey = .List(i) Then
            .Selected(i) = True
            bRet = True

        End If

    Next

End With

If bRet = False Then
    MsgBox "查询不到该DN", vbInformation, "提示"

End If




End Sub

Private Sub cmdReport_Click() '导出报表
    Dim strExportName           As String

    cmdReport.Enabled = False
    If Fps(0).MaxRows <= 0 Then
        MsgBox "没有可导出的数据！", vbInformation, "提示"
        Exit Sub
    End If
    '导出报表
    strExportName = GetExcelName(Trim(Fra(1).Caption))
    
    If cmbCombo1(0).ListIndex = 0 Then
        Call SEDCExportPrintExcel(RsClone, strExportName)                     'SEDC报表
    ElseIf cmbCombo1(0).ListIndex = 1 Then
        Call InputPackinglistExportPrintExcel(RsClone, strExportName)         'outputPackinglist报表
    ElseIf cmbCombo1(0).ListIndex = 2 Then
        Call InputInvoiceExportPrintExcel(RsClone, strExportName)             'outputInvoice报表
    ElseIf cmbCombo1(0).ListIndex = 3 Then
        Call Daily_InvExportPrintExcel(RsClone, strExportName)                'Daily_inventory_report
    ElseIf cmbCombo1(0).ListIndex = 4 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel(order, strExportName, 0)    'Shipping Packinglist
    ElseIf cmbCombo1(0).ListIndex = 5 Then
        If Not CheckData Then Exit Sub
        Call ShippingInvoiceExportPrintExcel(order, strExportName)          'Shipping invoice
    ElseIf cmbCombo1(0).ListIndex = 9 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel(order, strExportName, 1)
    End If
   
    
cmdReport.Enabled = True
End Sub

Private Sub cmdSearch_Click() '查询报表
Dim i                   As Long
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
   
        Dim strDNList As String
    '初始化FPS
    
                With lstLotID

            For i = 0 To .ListCount - 1
        
                If .Selected(i) = True Then
                    strDNList = strDNList & Trim$("" & .List(i)) & "','"
    

                End If

            Next

        End With

        
          strDNList = Mid(strDNList, 1, Len(strDNList) - 3)
    
    order = ""
    InitFps
    cmdReport.Enabled = True
    cmdExprot.Enabled = True
    cmdDXDC.Enabled = True
    cmdHBDC.Enabled = True
    
    '---------------------------------------------
    If cmbCombo1(0).ListIndex = 0 Then  'SEDC
    
              strSql = "select distinct 0 as 选择,a.containername,a.creationtimestamp as receive_date,to_char(ibwo.erpcreatedate,'yyyyww') testdc,'' as invlocation " & _
                ",wo.mpn_desc as device_name,wo.SOURCE_BATCH_ID as job_no,conn.firstname as lot_no,a.moveinqty as qty,wo.date_code,'' as invcomment " & _
                ",'new packing' as remark,'' as so,'' as invoice,Get_37Inv_MergeDetails(a.containername,wo.mtrl_num) as merge " & _
                "from a_wiplothistory a ,a_wiplotdetailshistory b ,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf " & _
                ",customeroitbl_test wo " & _
                "where a.specname='5272' and b.wiplothistoryid=a.wiplothistoryid " & _
                "and conn.containername=a.containername " & _
                "and waf.waferscribenumber=substr(conn.containername,1,instr(conn.containername,'-A')-1) " & _
                "and mfg.mfgordername=waf.workordername and ibwo.ordername=mfg.mfgordername " & _
                "and wo.source_batch_id=waf.wafernumber " & _
                " and wo.customershortname='37' and a.containername not like '%-F%' " & _
                "And a.creationtimestamp>='" & DTP(0).Value & "' And a.creationtimestamp<'" & DTP(1).Value + 1 & "'"
                
        If txt(1).text <> "" Then
            strSql = strSql & " And wo.SOURCE_BATCH_ID='" & Trim(txt(1).text) & "'"
        End If
        
        ElseIf cmbCombo1(0).ListIndex = 1 Then  'output packing list
     
        strSql = " select 选择,creationtimestamp outdate,po_num,po_item,PKG,DEVICE,lot_no,Job_No,Sublot_No,date_code,sum( qty) as qty,sum(Price) as Price,sum(USD) as USD,Test_Reject,Merge_in_Job,RMA_Number,mtrl_num from (" & _
                " select distinct 0 as 选择,a.containername, wo.po_num, wo.po_item,'' PKG,wo.mpn_desc DEVICE,wo.MTRL_NUM lot_no,waf.wafernumber Job_No," & _
                " conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No, wo.date_code,a.moveinqty qty, 0 as Price,a.moveinqty * 0 as USD," & _
                "  '' as Test_Reject, Get_37Inv_MergeDetails(a.containername,wo.MTRL_NUM) as Merge_in_Job,'' as RMA_Number, wo.mtrl_num,a.creationtimestamp " & _
                "  from a_wiplothistory a, a_wiplotdetailshistory b,container  conn,mfgorder  mfg,ib_wohistory  ibwo,a_lotwafers waf, customeroitbl_test wo " & _
                " where a.specname = '5272' and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
                "  and waf.waferscribenumber = substr(conn.containername, 1, instr(conn.containername, '-A') - 1) " & _
                "   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername and wo.source_batch_id = waf.wafernumber and status<>2 " & _
                "   and wo.customershortname = '37' and a.containername not like '%-F%' And a.creationtimestamp >= '" & DTP(0).Value & "'" & _
                "   And a.creationtimestamp < '" & DTP(1).Value & "' and   a.containername<>'10001-A-01' and not  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername=a.containername)  " & _
                " union select distinct 0 as 选择, a.containername, wo.po_num,wo.po_item,'' PKG, wo.mpn_desc DEVICE, wo.MTRL_NUM lot_no, waf.wafernumber as job_no, " & _
                " conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No,wo.date_code, waf.ndpw qty,0 as Price, a.moveinqty * 0 as USD," & _
                " '' as Test_Reject, Get_37Inv_MergeDetails(a.containername,wo.MTRL_NUM) as Merge_in_Job,'' as RMA_Number , wo.mtrl_num,a.creationtimestamp" & _
                "  from a_wiplothistory  a, a_wiplotdetailshistory b, container conn,mfgorder mfg, ib_wohistory ibwo, a_lotwafers  waf, customeroitbl_test  wo " & _
                " where a.specname = '5272' and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
                "   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername and wo.source_batch_id = waf.wafernumber " & _
                "   and wo.customershortname = '37' and a.containername not like '%-F%' And a.creationtimestamp >= '" & DTP(0).Value & "' " & _
                "   And a.creationtimestamp < '" & DTP(1).Value & "' and waf.containerid=conn.containerid  and  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername = a.containername)  " & _
                " )X Where 2>1 "
                
        If txt(1).text <> "" Then
            strSql = strSql & " And Job_No='" & Trim(txt(1).text) & "'"
        End If
        strSql = strSql & " group by 选择,po_num,po_item,PKG,DEVICE,lot_no,Job_No,Sublot_No,date_code,Test_Reject,Merge_in_Job,RMA_Number,mtrl_num,creationtimestamp order by Sublot_No,Merge_in_Job desc"
    
    ElseIf cmbCombo1(0).ListIndex = 2 Then  'output invoice
      
      strSql = " select 选择,creationtimestamp outdate,po_num,po_item, PKG, DEVICE,lot_no,job_no,Sublot_No,date_code, sum(qty), Price,sum(USD)  from (" & _
                "select distinct 0 as 选择,a.containername,wo.po_num,wo.po_item,'' PKG,wo.mpn_desc DEVICE,wo.MTRL_NUM lot_no,wo.SOURCE_BATCH_ID job_no " & _
                ",conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No,wo.date_code,a.moveinqty qty,wo.t_price as Price,a.moveinqty*wo.t_price as USD,a.creationtimestamp " & _
                "from a_wiplothistory a ,a_wiplotdetailshistory b ,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf " & _
                ",customeroitbl_test wo " & _
                "where a.specname='5272' and b.wiplothistoryid=a.wiplothistoryid " & _
                "and conn.containername=a.containername " & _
                "and waf.waferscribenumber=substr(conn.containername,1,instr(conn.containername,'-A')-1) " & _
                "and mfg.mfgordername=waf.workordername and ibwo.ordername=mfg.mfgordername " & _
                "and wo.source_batch_id=waf.wafernumber " & _
                " and wo.customershortname='37' and a.containername not like '%-F%' " & _
                "And a.creationtimestamp>='" & DTP(0).Value & "' And a.creationtimestamp<'" & DTP(1).Value & "'" & _
                "and a.containername<>'10001-A-01'and not  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername=a.containername)   union " & _
                " select  distinct 0 as 选择,a.containername,wo.po_num,wo.po_item,'' PKG,wo.mpn_desc DEVICE,wo.MTRL_NUM lot_no,waf.wafernumber job_no, " & _
                " conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No,wo.date_code,waf.ndpw qty,wo.t_price as Price,waf.ndpw* wo.t_price as USD,a.creationtimestamp " & _
                " from a_wiplothistory a,a_wiplotdetailshistory b,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf,customeroitbl_test wo " & _
                " where a.specname = '5272' and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername" & _
                " and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername " & _
                " and wo.source_batch_id = waf.wafernumber and wo.customershortname = '37' and a.containername not like '%-F%' " & _
                " And a.creationtimestamp >= '" & DTP(0).Value & "' And a.creationtimestamp < '" & DTP(1).Value & "' and waf.containerid=conn.containerid   and  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername = a.containername) " & _
                " )X Where 2>1 "
                
        If txt(1).text <> "" Then
            strSql = strSql & " And job_no='" & Trim(txt(1).text) & "'"
        End If
        strSql = strSql & " group by 选择,po_num,po_item, PKG, DEVICE, lot_no, job_no ,Sublot_No,date_code, Price,creationtimestamp order by Sublot_No"
    
    ElseIf cmbCombo1(0).ListIndex = 3 Then  ' Daily Inv
             
        strSql = " select 选择,max(RECEIVE_DATE) as RECEIVE_DATE,TESTDC,LOCATION,DEVICE_NAME,job_no,HTLOT_NO,sum(qty) as qty,date_code,CCOMMENT,Reel_Size,Remark,Move_in_Date from (" & _
                "select distinct 0 as 选择,a.containername,a.creationtimestamp as RECEIVE_DATE,to_char(ibwo.erpcreatedate,'yyww') TESTDC,'' as LOCATION " & _
                ",wo.mpn_desc DEVICE_NAME,wo.SOURCE_BATCH_ID||Get_37Inv_MergeStatus(a.containername) job_no " & _
                ",conn.firstname||Get_37Inv_MergeStatus(a.containername) HTLOT_NO,a.moveinqty qty,wo.date_code,'' CCOMMENT,'NEW PACKING' Reel_Size,'' Remark,'' Move_in_Date " & _
                "from a_wiplothistory a ,a_wiplotdetailshistory b ,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf " & _
                ",customeroitbl_test wo " & _
                "where a.specname='5272' and b.wiplothistoryid=a.wiplothistoryid " & _
                "and conn.containername=a.containername " & _
                "and waf.waferscribenumber=substr(conn.containername,1,instr(conn.containername,'-A')-1) " & _
                "and mfg.mfgordername=waf.workordername and ibwo.ordername=mfg.mfgordername " & _
                "and wo.source_batch_id=waf.wafernumber " & _
                " and wo.customershortname='37' and a.containername not like '%-F%'  and status<>'2' " & _
                "And a.creationtimestamp>='" & DTP(0).Value & "' And a.creationtimestamp<'" & DTP(1).Value + 1 & "'" & _
                " ) X Where 2>1 "
                
        If txt(1).text <> "" Then
            strSql = strSql & " And job_no='" & Trim(txt(1).text) & "'"
        End If
        strSql = strSql & " group by 选择,TESTDC,LOCATION,DEVICE_NAME,job_no,HTLOT_NO,date_code,CCOMMENT,Reel_Size,Remark,Move_in_Date"
        
  ElseIf cmbCombo1(0).ListIndex = 4 Then  'shipping packing list (INNER JOIN tblCustomerOI d ON CASE WHEN charindex('M',RTRIM(c.batchnumber))>0 THEN LEFT(RTRIM(c.batchnumber),LEN(RTRIM(c.batchnumber))-1) ELSE RTRIM(c.batchnumber) END=d.SOURCE_BATCH_ID)
        strSql = "SELECT X.* FROM ( SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",箱号,料号,replace(mpn_desc,'.P2','') as mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,净重,毛重,尺寸 " & _
                 " FROM Vw_InvShippedPLFor37_NEW a " & _
                 " WHERE 发货日期>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and 发货日期<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "' " & _
                 " union all " & _
                 "SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",箱号,料号,replace(mpn_desc,'.P2','') as mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,净重,毛重,尺寸 " & _
                 " FROM Vw_InvShippedPLFor37 a " & _
                 " WHERE 发货日期>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and 发货日期<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "'  ) X "
                 
        If txt(0).text <> "" Then
            strSql = strSql & " And 单据编号='" & Trim(txt(0).text) & "'"
        End If
        If txt(1).text <> "" Then
            strSql = strSql & " And batchnumber='" & Trim(txt(1).text) & "'"
        End If
 
     
'            strSql = "SELECT 0 选择,h.代码简称 + a.单据编号 单据编号,c.delivery,dbo.usp_date(a.单据日期) 发货日期,ISNULL(dn_address_new, c.shiptoname) AS shiptoname,ISNULL(x.ship_to_street1_new, c.shiptostreet1) AS shiptostreet1, " & _
'"ISNULL(x.ship_to_street2_new, c.shiptostreet2) AS shiptostreet2,ISNULL(x.ship_to_street3_new, c.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, c.city) AS city, ISNULL(x.dn_st_new, c.State) AS state, " & _
'"ISNULL(x.postal_code_new, c.postalcode) AS postalcode,ISNULL(x.country_new, c.countrykey) AS countrykey,ISNULL(x.contact_new, c.contactname) AS contactname,ISNULL(x.phone_new, c.phone) AS phone, " & _
'"c.SalesDocument,'''' + c.PurchasingDocNo AS PurchasingDocNo,RTRIM(erpdata.dbo.f_getparent(b.箱号)) 箱号,b.料号, " & _
'"CASE WHEN RTRIM(gg.MPN_DESC) = 'UCLAMP0541Z.TFT' THEN RTRIM(gg.MPN_DESC) + '.P2' ELSE REPLACE(REPLACE(gg.MPN_DESC, '.P2', ''), '.P3', '') END AS mpn_desc,SUM(b.数量) 数量,c.batchnumber, " & _
'"hh.CREATE_DATE DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no,c.customerPartNumber,ROUND(CAST(f.重量 AS FLOAT) * 0.4, 2) 净重,f.重量 毛重,f.尺寸 FROM erpdata .. tblStockSQfh a " & _
'"INNER JOIN erpdata .. tblStocksqfhsub b ON a.单据编号 = b.单据编号 AND a.序号 = b.单据项次 INNER JOIN erpdata .. tblStockNumTree g ON g.箱号 = b.箱号 INNER JOIN erpdata .. tblStockNumTree f " & _
'"ON f.序号 = g.上级序号 INNER JOIN (SELECT a.BOX_ID,SUBSTRING(a.KEY_VALUE,1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox,SUBSTRING(a.KEY_VALUE,CHARINDEX('|', a.KEY_VALUE) + 1,10) AS job " & _
'"FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON g.箱号 = aa.qbox INNER JOIN (SELECT dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, " & _
'"dn.shiptostreet3, dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber " & _
'"FROM tblCustomerShippingUp dn GROUP BY dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname, " & _
'"dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,dn.customerPartNumber) c ON c.Delivery = g.DN AND c.BatchNumber = aa.job INNER JOIN dbo.tblstock h " & _
'"ON CONVERT(NVARCHAR(4), h.库房代码) = CONVERT(NVARCHAR(4), a.仓库编号) INNER JOIN ERPBASE .. tblmappingData ff ON ff.SUBSTRATEID = b.流程卡编号 " & _
'"INNER JOIN ERPBASE .. tblCustomerOI gg ON CONVERT(VARCHAR(100), gg.ID) = ff.FILENAME AND gg.SOURCE_BATCH_ID = ff.LOTID AND gg.CUSTOMERSHORTNAME = '37' INNER JOIN erpbase .. weight37 hh " & _
'"ON hh.WAFERID = REPLACE(b.流程卡编号, '+', '') INNER JOIN erpdata .. tblErpInStockMainData jj ON jj.BOX_ID = aa.BOX_ID LEFT JOIN erptemp .. dn_address x ON dn_address = c.ShipToName " & _
'"WHERE a.客户代码 = '37' AND a.单据编号 LIKE 'F%' AND a.单据日期 >= CONVERT(VARCHAR(100), GETDATE() - 5, 23) AND c.Delivery = g.DN " & _
'"and c.Delivery in ('" & strDNList & "') GROUP BY h.代码简称, a.单据编号, c.delivery,dbo.usp_date(a.单据日期), c.shiptoname,c.shiptostreet1,c.shiptostreet2,c.shiptostreet3,c.city,c.State,c.postalcode, " & _
'"c.countrykey,c.contactname,c.phone,c.SalesDocument,c.PurchasingDocNo,erpdata.dbo.f_getparent(b.箱号),b.料号,gg.MPN_DESC,c.batchnumber,hh.CREATE_DATE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2), " & _
'"c.customerPartNumber,f.重量,f.尺寸,dn_address_new,x.ship_to_street1_new,x.ship_to_street2_new,x.ship_to_street3_new,x.city_new,x.dn_st_new,x.postal_code_new,x.country_new,x.contact_new,x.phone_new"

            
    strSql = "SELECT 0 AS  选择,y.代码简称 + b.单据编号 AS 单据编号,d.Delivery, CONVERT(VARCHAR(100), c.单据日期,23) AS 发货日期, ISNULL(dn_address_new, d.shiptoname) AS shiptoname, ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1 " & _
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
 " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber, f.重量 , f.尺寸,y.代码简称,d.Delivery order by shiptoname, Delivery,箱号"



        
    ElseIf cmbCombo1(0).ListIndex = 5 Then  'shipping invoice (INNER JOIN tblCustomerOI d ON CASE WHEN charindex('M',RTRIM(c.batchnumber))>0 THEN LEFT(RTRIM(c.batchnumber),LEN(RTRIM(c.batchnumber))-1) ELSE RTRIM(c.batchnumber) END=d.SOURCE_BATCH_ID)
        strSql = "select x.* from ( SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
                 ",city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",箱号,料号,replace(mpn_desc,'.P2','') as mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,单价,AMount,customerPartNumber " & _
                 " FROM Vw_InvShippedInvoiceFor37_NEW a " & _
                 " WHERE 发货日期>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and 发货日期<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "' " & _
                 "union all " & _
                 " SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
                 ",city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",箱号,料号,replace(mpn_desc,'.P2','') as mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,单价,AMount,customerPartNumber " & _
                 " FROM Vw_InvShippedInvoiceFor37 a " & _
                 " WHERE 发货日期>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and 发货日期<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "') x "
                 
        
        
                If txt(0).text <> "" Then
            strSql = strSql & " And 单据编号='" & Trim(txt(0).text) & "'"
        End If
        If txt(1).text <> "" Then
            strSql = strSql & " And batchnumber='" & Trim(txt(1).text) & "'"
        End If
        
        
        
        
        
  
        
        
'
'        strSql = "SELECT 0 选择, h.代码简称+a.单据编号 单据编号,c.delivery,dbo.usp_date(a.单据日期) 发货日期, ISNULL(dn_address_new, c.shiptoname) AS  shiptoname,ISNULL(x.ship_to_street1_new ,c.shiptostreet1)  AS shiptostreet1  " & _
'",ISNULL(x.ship_to_street2_new,c.shiptostreet2) AS shiptostreet2 ,ISNULL(x.ship_to_street3_new,c.shiptostreet3) AS shiptostreet3 ,ISNULL(x.city_new ,c.city) AS city,ISNULL(x.dn_st_new, c.State) AS state,ISNULL(x.postal_code_new,c.postalcode  ) AS  postalcode " & _
'",ISNULL(x.country_new,c.countrykey) AS countrykey,ISNULL(x.contact_new,c.contactname) AS contactname ,ISNULL(x.phone_new,c.phone ) AS phone      " & _
'",c.SalesDocument,'''' + c.PurchasingDocNo AS PurchasingDocNo,RTRIM(erpdata.dbo.f_getparent(b.箱号)) 箱号,b.料号,CASE WHEN RTRIM(gg.MPN_DESC)='UCLAMP0541Z.TFT' THEN RTRIM(gg.MPN_DESC)+'.P2' " & _
'"ELSE REPLACE(REPLACE(gg.MPN_DESC,'.P2',''),'.P3','') END AS mpn_desc,SUM(b.数量) 数量,c.batchnumber,hh.CREATE_DATE DATE_CODE " & _
'",SUBSTRING(aa.qbox ,2,CHARINDEX('-R',aa.qbox)-2)  HTlot_no ,ISNULL(ISNULL( BB.含税单价 / AB.良品数,0)  + ( cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0) AS 单价 " & _
'",ROUND( SUM(b.数量) * ISNULL(ISNULL( BB.含税单价 / AB.良品数,0)  +  (cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0),2) AS AMount,c.customerPartNumber  ,e.销售单编号  " & _
'",ROUND( SUM(b.数量) * ISNULL(ISNULL( BB.含税单价 / AB.良品数,0)  +  (cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0),2) -  CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( BB.含税单价 / AB.良品数,0)) AS 加工费金额 " & _
'", CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( BB.含税单价 / AB.良品数,0)) AS 客供料金额 FROM erpdata..tblStockSQfh a           " & _
'"INNER JOIN erpdata..tblStocksqfhsub b ON a.单据编号 = b.单据编号 AND a.序号=b.单据项次 INNER JOIN erpdata..tblStockNumTree g ON g.箱号=b.箱号 " & _
'"INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE,1,CHARINDEX('|',a.KEY_VALUE)-1) AS qbox , SUBSTRING(a.KEY_VALUE,CHARINDEX('|',a.KEY_VALUE)+1,10) AS job " & _
'" FROM erpdata..tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND  a.KEY_VALUE LIKE '%SS%|%')  aa ON g.箱号 = aa.qbox  " & _
'"INNER JOIN (SELECT dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode "
'
'        strSql = strSql & ",dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber FROM tblCustomerShippingUp dn GROUP BY dn.Delivery,dn.shiptoname,dn.shiptostreet1 " & _
'",dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,dn.customerPartNumber)  c ON c.Delivery = g.DN  AND c.BatchNumber =  aa.job " & _
'"INNER JOIN dbo.tblstock h ON CONVERT(NVARCHAR(4),h.库房代码) = CONVERT(NVARCHAR(4),a.仓库编号)    INNER JOIN ERPBASE..tblmappingData ff ON ff.SUBSTRATEID = b.流程卡编号 " & _
'"INNER JOIN ERPBASE..tblCustomerOI gg ON CONVERT(VARCHAR(30), gg.ID) = ff.FILENAME AND gg.SOURCE_BATCH_ID = ff.LOTID AND gg.CUSTOMERSHORTNAME = '37' " & _
'"INNER JOIN erpbase..weight37 hh ON hh.WAFERID = REPLACE(b.流程卡编号,'+','') INNER JOIN erpdata..tblErpInStockMainData jj ON jj.BOX_ID = aa.BOX_ID " & _
'"LEFT JOIN erpbase..tbltoinrec_wafer  AB ON ab.批号 = ff.LOTID AND AB.晶圆ID = REPLACE(B.流程卡编号,'+','') LEFT JOIN erpbase..tbltorec_wafer  ww ON  ww.批号 = ab.批号 AND  ww.晶圆ID = ab.晶圆ID  " & _
'"LEFT JOIN ERPBASE..TblToInsub BB ON BB.入库单编号 = AB.入库单编号 AND BB.到货批号 = AB.批号 AND ww.到货单编号 = bb.到货单编号 AND bb.含税单价 IS NOT NULL " & _
'"LEFT JOIN erptemp..tblBB_CSRPO cb ON cb.PO_NUM = gg.PO_NUM AND cb.FAB_DEVICE = gg.MPN_DESC LEFT JOIN ERPBASE..tblmappingData db ON db.SUBSTRATEID = REPLACE(B.流程卡编号,'+','')  " & _
'"LEFT JOIN  erpdata..tblSalerec e ON  e.单据编号 = a.单据编号       AND a.序号 = e.单据项次  AND e.小箱号 = b.箱号    " & _
'"LEFT JOIN erptemp..dn_address x ON dn_address = c.ShipToName WHERE a.客户代码='37' and c.Delivery in ('" & strDNList & "') AND a.单据日期 >= CONVERT(VARCHAR(100),GETDATE()- 8,23) AND  a.单据编号 LIKE 'F%' AND a.良品数量 >0  " & _
'"GROUP BY gg.PO_NUM,h.代码简称,a.单据编号,c.delivery,dbo.usp_date(a.单据日期),c.shiptoname,c.shiptostreet1,c.shiptostreet2         " & _
'",c.shiptostreet3,c.city,c.State,c.postalcode,c.countrykey,c.contactname,c.phone,c.SalesDocument,c.PurchasingDocNo ,erpdata.dbo.f_getparent(b.箱号),b.料号,gg.MPN_DESC,c.batchnumber,hh.CREATE_DATE      " & _
'",SUBSTRING(aa.qbox ,2,CHARINDEX('-R',aa.qbox)-2),e.销售单编号,c.customerPartNumber ,ISNULL( BB.含税单价 / AB.良品数,0) ,  cb.WAFER_PRICE,db.PASSBINCOUNT , cb.DIE_PRICE,dn_address_new " & _
'",x.ship_to_street1_new,x.ship_to_street2_new,x.ship_to_street3_new,x.city_new,x.dn_st_new,x.postal_code_new,x.country_new,x.contact_new ,x.phone_new "

'      strsql = " SELECT 0 AS 选择,y.代码简称 + b.单据编号 AS 单据编号,a.DN, CONVERT(VARCHAR(100), c.单据日期,23) AS 发货日期, ISNULL(dn_address_new, d.shiptoname) AS shiptoname " & _
' ",ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1,ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2 , ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, d.city) AS city, " & _
' " ISNULL(x.dn_st_new, d.State) AS state, ISNULL(x.postal_code_new, d.postalcode) AS postalcode,ISNULL(x.country_new, d.countrykey) AS countrykey, ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone,d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo, f.箱号, " & _
' " b.料号,d.MarketingPN,SUM(b.数量), d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no ,ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  + ( dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0) AS 单价 " & _
' " ,ROUND( SUM(b.数量) * ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  +  (dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0),2) AS AMount ,d.customerPartNumber ,ROUND( SUM(b.数量) * ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  +  (dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0),2) " & _
' "  -  CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( cd.含税单价 / cc.良品数,0)) AS 加工费金额 , CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( cd.含税单价 / cc.良品数,0)) AS 客供料金额 " & _
'"   FROM erpdata .. tblStockNumTree a INNER JOIN erpdata .. tblStocksqfhsub b ON b.箱号 = a.箱号 INNER JOIN erpdata .. tblStockSQfh c  ON c.单据编号 = b.单据编号 AND c.序号 = b.单据项次 INNER JOIN erpdata..tblstock y ON y.库房代码 = c.仓库编号 " & _
'"  INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE, 1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job  FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.箱号 = aa.qbox  INNER JOIN (SELECT dn.Delivery, " & _
' "  dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State,dn.postalcode, dn.countrykey, dn.contactname,dn.phone, dn.SalesDocument,  dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE  FROM ERPBASE .. tblCustomerShippingUp dn " & _
' "   WHERE dn.Delivery IN ('" & strDNList & "')  GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State, dn.postalcode,  dn.countrykey,dn.contactname,dn.phone, dn.SalesDocument, dn.PurchasingDocNo, dn.BatchNumber, dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE) d " & _
' "   ON d.Delivery = a.DN  AND d.BatchNumber = aa.job INNER JOIN erpdata .. tblStockNumTree e  ON e.箱号 = b.箱号 INNER JOIN erpdata .. tblstocknumtree f  ON f.序号 = e.上级序号  INNER JOIN ERPBASE..tblmappingData qq ON qq.SUBSTRATEID = b.流程卡编号 AND qq.LOTID = b.工单号 LEFT JOIN ERPBASE..tblCustomerOI bb " & _
' " ON CONVERT(VARCHAR(100), bb.ID) = qq.FILENAME AND bb.SOURCE_BATCH_ID = qq.LOTID  LEFT JOIN ERPBASE..tblToRec_Wafer cc ON cc.晶圆ID = REPLACE(qq.SUBSTRATEID,'+','')   AND cc.批号 = qq.LOTID  LEFT JOIN ERPBASE..tblToRecEntry cd ON cd.到货单编号 = cc.到货单编号 AND cd.到货批号 = cc.批号 LEFT JOIN erptemp..tblBB_CSRPO dd " & _
'"  ON dd.PO_NUM = bb.PO_NUM AND dd.FAB_DEVICE = bb.MPN_DESC LEFT JOIN  erpdata..tblSalerec j ON j.单据编号 = c.单据编号 AND j.单据项次 = b.单据项次 AND j.小箱号 = b.箱号 LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName  WHERE a.DN IN ('" & strDNList & "')  GROUP BY b.单据编号, c.单据日期, ISNULL(dn_address_new, d.shiptoname), " & _
'"  ISNULL(x.ship_to_street1_new, d.shiptostreet1), ISNULL(x.ship_to_street2_new, d.shiptostreet2),ISNULL(x.ship_to_street3_new, d.shiptostreet3), ISNULL(x.city_new, d.city),  ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode),ISNULL(x.country_new, d.countrykey),ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone), " & _
'" d.SalesDocument, d.PurchasingDocNo,f.箱号, b.料号,d.MarketingPN,d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber,   a.DN , J.销售单编号, bb.PO_NUM, cd.含税单价, cc.良品数, dd.WAFER_PRICE, qq.PASSBINCOUNT, dd.die_price, y.代码简称  order by shiptoname,DN,箱号 "
'
'
          
   
strSql = "SELECT 0 AS 选择,y.代码简称 + b.单据编号 AS 单据编号,a.DN, CONVERT(VARCHAR(100), c.单据日期,23) AS 发货日期, ISNULL(dn_address_new, d.shiptoname) AS shiptoname ,ISNULL(x.ship_to_street1_new " & _
         " , d.shiptostreet1) AS shiptostreet1,ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2 , ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, d.city) AS city " & _
         " ,  ISNULL(x.dn_st_new, d.State) AS state, ISNULL(x.postal_code_new, d.postalcode) AS postalcode,ISNULL(x.country_new, d.countrykey) AS countrykey, ISNULL(x.contact_new, d.contactname) AS contactname " & _
         " , ISNULL(x.phone_new, d.phone) AS phone,d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo, f.箱号,  b.料号,d.MarketingPN,SUM(b.数量), d.BatchNumber, d.DATE_CODE " & _
         " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no ,ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  + ( dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0) AS 单价 ,ROUND( SUM(b.数量) * ISNULL(ISNULL( cd.含税单价 / cc.良品数,0) " & _
         "  +  (dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0),2) AS AMount ,d.customerPartNumber,ROUND( SUM(b.数量) * ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  +  (dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0),2)   -  CONVERT(DECIMAL(18,2) " & _
         " ,SUM(b.数量) * ISNULL( cd.含税单价 / cc.良品数,0)) AS 加工费金额 , CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( cd.含税单价 / cc.良品数,0)) AS 客供料金额 " & _
         "  FROM erpdata .. tblStockNumTree a INNER JOIN erpdata .. tblStocksqfhsub b ON b.箱号 = a.箱号 INNER JOIN erpdata .. tblStockSQfh c  ON c.单据编号 = b.单据编号 AND c.序号 = b.单据项次 " & _
         " INNER JOIN erpdata..tblstock y ON y.库房代码 = c.仓库编号   INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE, 1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox " & _
         " , SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job  FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.箱号 = aa.qbox " & _
         " INNER JOIN (SELECT dn.Delivery,   dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State,dn.postalcode, dn.countrykey, dn.contactname,dn.phone, dn.SalesDocument " & _
         " ,  dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE  FROM ERPBASE .. tblCustomerShippingUp dn " & _
         " WHERE dn.Delivery IN ('" & strDNList & "')  GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State, dn.postalcode,  dn.countrykey,dn.contactname,dn.phone " & _
         " , dn.SalesDocument, dn.PurchasingDocNo, dn.BatchNumber, dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE) d    ON d.Delivery = a.DN  AND d.BatchNumber = aa.job " & _
         " INNER JOIN erpdata .. tblstocknumtree f  ON f.序号 = a.上级序号  INNER JOIN ERPBASE..tblmappingData qq ON qq.SUBSTRATEID = b.流程卡编号 AND qq.LOTID = b.工单号 " & _
         " LEFT JOIN ERPBASE..tblCustomerOI bb  ON CONVERT(VARCHAR(100), bb.ID) = qq.FILENAME AND bb.SOURCE_BATCH_ID = qq.LOTID  LEFT JOIN ERPBASE..tblToInRec_Wafer cc " & _
         " ON cc.晶圆ID = REPLACE(qq.SUBSTRATEID,'+','')   AND cc.批号 = qq.LOTID  LEFT JOIN ERPBASE..TblToInSub cd ON cd.入库单编号 = cc.入库单编号 AND cd.到货批号 = cc.批号 " & _
         " LEFT JOIN erptemp..tblBB_CSRPO dd   ON dd.PO_NUM = bb.PO_NUM AND dd.FAB_DEVICE = bb.MPN_DESC LEFT JOIN  erpdata..tblSalerec j ON j.单据编号 = c.单据编号 AND j.单据项次 = b.单据项次 " & _
         " AND j.小箱号 = b.箱号 LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName  WHERE a.DN IN ('" & strDNList & "')  GROUP BY b.单据编号, c.单据日期, ISNULL(dn_address_new, d.shiptoname) " & _
         "  ,   ISNULL(x.ship_to_street1_new, d.shiptostreet1), ISNULL(x.ship_to_street2_new, d.shiptostreet2),ISNULL(x.ship_to_street3_new, d.shiptostreet3), ISNULL(x.city_new, d.city) " & _
         "  ,  ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode),ISNULL(x.country_new, d.countrykey),ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) " & _
         " ,  d.SalesDocument, d.PurchasingDocNo,f.箱号, b.料号,d.MarketingPN,d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber,   a.DN " & _
         " , J.销售单编号, bb.PO_NUM, cd.含税单价, cc.良品数, dd.WAFER_PRICE, qq.PASSBINCOUNT, dd.die_price, y.代码简称  order by shiptoname,DN,箱号 "

          
          
          
          
          
          
    ElseIf cmbCombo1(0).ListIndex = 6 Then  '库存报表
        If cmbCombo1(1).text = "9000" Then
            strSql = " SELECT * FROM DBO.Vw_InvStockRptFor37By9000 "
            strSql = "select * from [dbo].[Vw_InvStockRptFor37By9000_new_temp] "
        
            strSql = "select RECEIVE_DATE,[Wafer Type],[Assy Part#],[Fab Lot],ID,[D/C],[QTY(称重后的数量)],Job#,Bag#,Comment,NCMR,[HT Part#],[保税/非保税], [库房编号] from erpbase..Vw_InvStockRptFor37By9000_temp a where a.库房编号 in ('44','45') " & _
"Union select RECEIVE_DATE,[Wafer Type],[Assy Part#],[Fab Lot],ID,[D/C],sum([QTY(称重后的数量)]),Job#,Bag#,Comment,NCMR,[HT Part#],[保税/非保税],[库房编号] from erpbase..Vw_InvStockRptFor37By9000_new_temp a where a.库房编号 in ('44','45') " & _
"group by RECEIVE_DATE,[Wafer Type],[Assy Part#],[Fab Lot],ID,[D/C],Job#,Bag#,Comment,NCMR,[HT Part#],[保税/非保税],[库房编号] " & _
"order by [Assy Part#]"
            
            
        Else
            strSql = "SELECT 0 选择,RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Reel_Size,Remark" & _
                     ",Move_in_Date FROM Vw_InvStockRptFor37 Where 2>1 "
            If txt(1).text <> "" Then
                strSql = strSql & " And JOB_NO='" & Trim(txt(1).text) & "'"
            End If
            If cmbCombo1(1).text <> "所有" Then
                strSql = strSql & " And 仓库名称='" & Trim(cmbCombo1(1).text) & "'"
            End If
        End If
    ElseIf cmbCombo1(0).ListIndex = 7 Then  'SMTCList
        strSql = "SELECT 0 选择,Invoice_No,Carton_No,PartName,LotID,QTY,Job_No,DATE_CODE " & _
                 " FROM Vw_InvShippedSMTCListFor37 " & _
                 " WHERE 发货日期>='" & DTP(0).Value & "' and 发货日期<'" & DTP(1).Value + 1 & "' "
        If txt(0).text <> "" Then
            strSql = strSql & " And 单据编号='" & Trim(txt(0).text) & "'"
        End If
        If txt(1).text <> "" Then
            strSql = strSql & " And Job_No='" & Trim(txt(1).text) & "'"
        End If
    ElseIf cmbCombo1(0).ListIndex = 8 Then  'Shipped
        strSql = "SELECT 0 选择,RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Remark,SO as [DN#],SHIPPED_DATE,Cust_Name " & _
                 " FROM Vw_InvShippedRptFor37 " & _
                 " WHERE SHIPPED_DATE>='" & DTP(0).Value & "' and SHIPPED_DATE<'" & DTP(1).Value + 1 & "' "
        If txt(0).text <> "" Then
            strSql = strSql & " And 单据编号='" & Trim(txt(0).text) & "'"
        End If
        If txt(1).text <> "" Then
            strSql = strSql & " And JOB_NO='" & Trim(txt(1).text) & "'"
        End If
    ElseIf cmbCombo1(0).ListIndex = 9 Then  'Shipped
          strSql = "SELECT 0 AS  选择,y.代码简称 + b.单据编号 AS 单据编号,d.Delivery, CONVERT(VARCHAR(100), c.单据日期,23) AS 发货日期, ISNULL(dn_address_new, d.shiptoname) AS shiptoname, ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1 " & _
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
         " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber, f.重量 , f.尺寸,y.代码简称,d.Delivery order by shiptoname, Delivery,箱号"
   
    End If
    '赋值到FRA(1)中 INIadoCon
    Fra(1).Caption = cmbCombo1(0).text
    If rs.State = adStateOpen Then rs.Close
    If cmbCombo1(0).ListIndex = 0 Or cmbCombo1(0).ListIndex = 1 Or cmbCombo1(0).ListIndex = 2 Or cmbCombo1(0).ListIndex = 3 Then
        rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    Else
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    End If
    Fps(0).MaxRows = 0
    Set RsClone = Nothing
    
    If rs.EOF Then
        MsgBox "无数据！"
    End If
    
    If Not rs.EOF Then
        Set RsClone = rs.Clone '克隆一份数据到另一个数据集中，为后面使用
        With Fps(0)
            .MaxRows = 0
            Set .DataSource = rs
            .MaxRows = rs.RecordCount
        End With
    End If
    rs.Close
    '特殊几个报表增加汇总栏位
    CalcTotal
    
'      With lstLotID
'
'        For i = 0 To .ListCount - 1
'            .Selected(i) = False
'        Next
'
'    End With
End Sub
Private Sub CalcTotal()
'计算汇总
Dim i                   As Long
Dim dblTotal            As Double
Dim colTotal            As Integer
    
    If cmbCombo1(0).ListIndex <> 6 And cmbCombo1(0).ListIndex <> 7 And cmbCombo1(0).ListIndex <> 8 Then Exit Sub
    
    dblTotal = 0
    colTotal = 0
    With Fps(0)
        If .MaxRows <= 0 Then Exit Sub
        For i = 1 To .MaxRows
            .Row = i
            If cmbCombo1(0).ListIndex = 6 Or cmbCombo1(0).ListIndex = 8 Then
                colTotal = 8
                .Col = colTotal
            Else
                colTotal = 6
                .Col = colTotal
            End If
            dblTotal = dblTotal + Val(Trim$(.text))
        Next
        If dblTotal > 0 Then '表示有数量
            .MaxRows = .MaxRows + 1
            .SetText colTotal, .MaxRows, dblTotal
        End If
        .DeleteCols fpSDetail.E_CHOOSE, 1
        .MaxCols = .MaxCols - 1
    End With
    
End Sub

Private Sub cmdUpload_Click()
Dim strFilePath         As String
Dim strFileName         As String
Dim strSql              As String
Dim image_Data()        As Byte         '图片二进制
Dim rs                  As New ADODB.Recordset
    '打开图片
    Com.Filter = "上传文件(*.xls,*.xlsx)|*.xls;*.xlsx"
    Com.ShowOpen '打开对话框
    strFilePath = Trim(Com.filename)  '保存路径
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1) '文件名
    '开始保存到资料库
    '数据转换为流
    Open strFilePath For Binary As #1
    ReDim image_Data(LOF(1) - 1)
    Get #1, , image_Data()
    Close #1
    '查询是否保存过此图片
    
    
    
    'strsql = "Select * From TblPMC_PicInfo Where FileName='" & Trim$(strFilename) & "' For Update"
    
    strSql = " SELECT * FROM erpdata..tblSystemTemplet WHERE SYS_NAME='关务' AND UPPER(TEMPLETNAME)='Invoice.xls'"
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenKeyset, adLockOptimistic
    If Not rs.EOF Then
        rs("SYS_NAME") = "关务"
        rs("TEMPLETNAME") = "Invoice.xls"
        rs("FILECONTENT") = image_Data()
        rs.Update
    Else
        rs.AddNew
        rs("FileName") = strFileName
        rs("FilePath") = strFilePath
        rs("FileComent") = image_Data()
        rs("Flag") = 1
        '记得添加数据库中txt存放的路径和txt框
        rs.Update
    End If
    rs.Close
    
    MsgBox "上传成功", vbInformation, "提示"
    
End Sub

Private Sub Command1_Click()
Dim strExportName As String
Dim i             As Integer

cmdReport.Enabled = False
strExportName = GetExcelName(Trim(Fra(1).Caption))
Call ShippingPackinglistExportPrintExcel(order, strExportName, 0)    'Shipping Packinglist

MsgBox "装箱单已经导出完成", vbInformation, "提示"
Call ShippingInvoiceExportPrintExcel(order, strExportName)          'Shipping invoice
cmdReport.Enabled = True
MsgBox "Invoice已经导出完成", vbInformation, "提示"

End Sub

Private Sub Command2_Click()
Dim i As Integer
With lstLotID

        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next

End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Fra(2).Move C_Left, Fra(2).Top, Me.ScaleWidth - C_Left, Fra(2).Height
    Fra(0).Move C_Left, Fra(0).Top, Fra(0).Width, Me.ScaleHeight - Fra(2).Height
    Fra(1).Move Fra(1).Left, Fra(1).Top, Me.ScaleWidth - Fra(0).Width - C_Top, Me.ScaleHeight - Fra(2).Height
    Fps(0).Move C_Left, Fps(0).Top, Fra(1).Width - C_Top, Me.ScaleHeight - Fra(2).Height - 3 * C_Top
End Sub
Private Sub Form_Load()
    If gUserName = "07885" Then
        cmdUpload.Visible = True
    End If
    
    DirQrShare = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\jpg"
    DirShare = App.Path & "\NewSemtechReport"               '系统路径
    DirFileShare = App.Path & "\SemtechExcelReport"         '系统Excel文件路径
    DirInvRpt = "C:\37-InventoryRpt"                        '库存报表存放路径
    '判定文件夹是否存在,不存在就创建
    If Dir(DirShare, vbDirectory) = "" Then
        MkDir DirShare                                      '创建文件夹
    End If
    If Dir(DirFileShare, vbDirectory) = "" Then
        MkDir DirFileShare                                  '创建文件夹
    End If
    If Dir(DirInvRpt, vbDirectory) = "" Then
        MkDir DirInvRpt                                     '创建文件夹
    End If
    '初始化控件
    InitCtrl
    
    
    
End Sub
Private Sub CheckXls()
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
Dim image_filename      As String
Dim temp_image()        As Byte
Dim i                   As Integer
    '设定鼠标状态
    Screen.MousePointer = 0
    strSql = "Select * From TblPMC_PicInfo Where Flag=1 Order by Create_Date"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        For i = 1 To rs.RecordCount
            '加载图片
            temp_image = rs("FileComent")
            image_filename = DirShare & "\" & rs("FileName")
            Open image_filename For Binary As #1
            Close #1
          '  Put #1, , temp_image()
          '  Close #1
            rs.MoveNext
        Next
    End If
    rs.Close

End Sub
'初始化控件
Private Sub InitCtrl()
Dim i                   As Integer
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
 Dim Rs2                  As New ADODB.Recordset
    
    strSql = "select  distinct dn_num from packing_detailed  where create_date > sysdate - 30  order by dn_num desc"
     
    Set Rs2 = Get_OracleRs(strSql)
'Show
lstLotID.Clear
If Not Rs2.EOF Then

    Do While Not Rs2.EOF
        lstLotID.AddItem Trim("" & Rs2!DN_NUM)
        Rs2.MoveNext
    Loop
Else
End If
    
    
    strdjbh = ""
    '加载单据类型
    strSql = "select para_1 from tblsys_parameter where sysname='TSVSYS' and kind='Semtech报表' order by id "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    cmbCombo1(0).Clear
    If Not rs.EOF Then
        Do While Not rs.EOF
            cmbCombo1(0).AddItem Trim$("" & rs!para_1)
            rs.MoveNext
        Loop
        cmbCombo1(0).ListIndex = 0
    End If
    rs.Close
'    '加载仓库
'    cmbCombo1(1).Clear
'    cmbCombo1(1).AddItem "所有"
'    cmbCombo1(1).AddItem "1000"
'    cmbCombo1(1).AddItem "6000"
'    cmbCombo1(1).AddItem "7000"
'    cmbCombo1(1).AddItem "8000"
'    cmbCombo1(1).AddItem "Scrap"
'    cmbCombo1(1).ListIndex = 0
    '初始化FPS
    InitFps
    
   DTP(0).Value = Format(Now() - 1, "YYYY-MM-DD")
   DTP(1).Value = Format(Now(), "YYYY-MM-DD")
   '检查模版
   CheckXls
   
End Sub
'初始化FPS控件
Public Sub InitFps()
Dim i                   As Integer
    'Fps初始化
    With Fps(0)
    .TypeMaxEditLen = 500
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsBest
        .MaxRows = 0
        .ColsFrozen = 2
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        '设定列类型
        .Col = fpSDetail.E_CHOOSE   '选择
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '设定列宽
        .ColWidth(-1) = 10
        .ColWidth(fpSDetail.E_CHOOSE) = 4
        .RowHeight(-1) = 10
        '设定是否排序
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ZOrder
        .ReDraw = True
    End With
End Sub
Private Sub fps_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i           As Long
Dim j           As Integer
Dim strTmp      As String
    
    '几个报表特殊处理
    
    cmdReport.Enabled = True
    cmdExprot.Enabled = True
    
    
    If cmbCombo1(0).ListIndex = 6 Or cmbCombo1(0).ListIndex = 7 Or cmbCombo1(0).ListIndex = 8 Then Exit Sub
    '点击把选择的单号都选上
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    With Fps(0)
'        .Col = FpsDetail.e_Choose
'        For i = 1 To .MaxRows
'            .Row = i
'            If i <> Row Then
'                .Col = FpsDetail.e_Choose
'                If Val(.Value) = 1 Then
''                    .Value = 0
'                    .Col = -1
'                    .ForeColor = vbBlack
'                End If
'            End If
'        Next

        .Col = fpSDetail.E_CHOOSE
        .Row = Row
        .Value = Abs(Val(.Value) - 1)
'        strDJBH = ""
        If Val(.Value) = 1 Then
            '将所有一样的单号选择上
            .Col = fpSDetail.e_DJBH
            .Row = Row
            strTmp = Trim$(.text)
'            strDJBH = Trim$(.Text) '共用的单据编号，在导出打印时会用到
            For i = 1 To .MaxRows
                .Row = i
                .Col = fpSDetail.e_DJBH
                If Trim$(.text) = strTmp Then
                    .Col = fpSDetail.E_CHOOSE
                    .Value = 1
                    .Col = -1
                    .ForeColor = &HFF8080
                End If
            Next
            
            order = strTmp & "'" & "," & "'" & order
            
        Else
            '将所有一样的单号选择上
            .Col = fpSDetail.e_DJBH
            .Row = Row
            strTmp = Trim$(.text)
'            strDJBH = Trim$(.Text) '共用的单据编号，在导出打印时会用到
            For i = 1 To .MaxRows
                .Row = i
                .Col = fpSDetail.e_DJBH
                If Trim$(.text) = strTmp Then
                    .Col = fpSDetail.E_CHOOSE
                    .Value = 0
                    .Col = -1
                    .ForeColor = vbBlack
                End If
            Next
            
            order = Replace(order, strTmp & "'" & "," & "'", "")
        End If
        
    End With
    
End Sub

'校验数据
Private Function CheckData() As Boolean
Dim i               As Integer
Dim intCount        As Integer
Dim strCust         As String

    CheckData = False
    
    strdjbh = ""     '--单据编号记录
    strCust = ""
    intCount = 0
    
    With Fps(0)
        If .MaxRows <= 0 Then
            MsgBox "没有任何资料,请先查询！", vbInformation, "提示"
            Exit Function
        End If
        '看是否有选择
        For i = 1 To .MaxRows
            .Row = i
            .Col = fpSDetail.E_CHOOSE  '选择
            If .Value = 1 Then
                intCount = intCount + 1
                .Col = fpSDetail.e_DJBH '单据编号
                If InStr(strdjbh, Trim$(.text)) <= 0 Then
                    strdjbh = strdjbh + Trim$(.text) + ","
                    strdjbh1 = Mid(strdjbh, 2, Len(strdjbh)) + Trim$(.text) + ","
                End If
            End If
        Next
    End With
    '去除单据编号最后一个逗号

    '--------------------------
    If intCount <= 0 Then
        MsgBox "没有选择任何资料！", vbInformation, "提示"
        Exit Function
    End If
    strdjbh = Left$(strdjbh, Len(strdjbh) - 1)
    strdjbh1 = Left$(strdjbh1, Len(strdjbh1) - 1)
    CheckData = True
End Function

'SEDC
Public Sub SEDCExportPrintExcel(ByVal rs As ADODB.Recordset, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double '总金额
Dim strExtsion          As String '后缀名
Dim strNewFullPath      As String '新Excel文件

    
    If rs.RecordCount <= 0 Then
        MsgBox "没有要导出的资料！", vbInformation, "提示！"
        Exit Sub
    End If
    
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "初始化数据..."
    
    strFileName = DirShare & "\SEDC.xls" '要打开的文件
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
    rs.MoveFirst    '数据集移动到第一个
    
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '是否显示
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        
        lngRows = 3
        IntInertRow = rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = rs.RecordCount
        
'        ClsP.ShowProgress 50, "正在导出..."
        
        Do While Not rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & rs.Fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & rs.Fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & rs.Fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & rs.Fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & rs.Fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & rs.Fields(10).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(11).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(12).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(13).Value)
            wkst.Cells(lngRows, 13) = Trim$("" & rs.Fields(14).Value)
            
            lngRows = lngRows + 1
            rs.MoveNext
        Loop
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
'        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
'        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
'        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
'        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.DESCRIPTION, vbInformation, "提示！"
    Exit Sub
End Sub

'Packing list
Public Sub InputPackinglistExportPrintExcel(ByVal rs As ADODB.Recordset, ByVal strExName As String)
    Dim strSql              As String
    Dim lngRows             As Long
    Dim rsQuery             As Excel.QueryTable
    Dim ExApp               As Excel.Application
    Dim wkbk                As New Workbook
    Dim wkst                As New Worksheet
    Dim i                   As Long
    Dim j                   As Long
    Dim IntCols             As Integer
    Dim strCols             As String
    Dim strFileName         As String
    Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
    Dim DblNum              As Double
    Dim DblAmt              As Double '总金额
    Dim strExtsion          As String '后缀名
    Dim strNewFullPath      As String '新Excel文件
    Dim strXH               As String '打印序号

    If rs.RecordCount <= 0 Then
        MsgBox "没有要导出的资料！", vbInformation, "提示！"
        Exit Sub
    End If
'   ClsP.Init 100, True
'   ClsP.ShowProgress 10, "初始化数据..."
    
    strFileName = DirShare & "\output_packing_list.xls" '要打开的文件
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
    strXH = Mid$(strExName, InStrRev(strExName, "_"))    '获取序号
    rs.MoveFirst    '数据集移动到第一个
        
    
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '是否显示
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        '赋值到Excel中，表头
        
        wkst.Cells(8, 12) = Format(DATE, "YYYY/mm/DD")
        wkst.Cells(9, 12) = "HTKS-SEDC" & Format(DATE, "YY-MMDD") & strXH
        
        lngRows = 17
        IntInertRow = rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = rs.RecordCount
        
'        ClsP.ShowProgress 50, "正在导出..."
        
        Do While Not rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & rs.Fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & rs.Fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & rs.Fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & rs.Fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & rs.Fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & rs.Fields(10).Value)
            
            'jiayun 修改 ，把0改为空值
            
'            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(11).Value)
'            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(12).Value)
            wkst.Cells(lngRows, 10) = ""
            wkst.Cells(lngRows, 11) = ""
            
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(13).Value)
            wkst.Cells(lngRows, 13) = Trim$("" & rs.Fields(14).Value)
            wkst.Cells(lngRows, 14) = Trim$("" & rs.Fields(15).Value)
            
            'jiayun add Bag#
            wkst.Cells(lngRows, 15) = Trim$("" & rs.Fields(16).Value)
            
            DblNum = DblNum + Val(Trim$("" & rs.Fields(10).Value))
            DblAmt = DblAmt + Val(Trim$("" & rs.Fields(12).Value))
            
            lngRows = lngRows + 1
            rs.MoveNext
        Loop
        
        wkst.Cells(lngRows, 9) = DblNum
        'wkst.Cells(lngRows, 11) = DblAmt
        
        wkst.Cells(lngRows, 11) = ""
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
'        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
'        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
'        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
'        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With
  '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.DESCRIPTION, vbInformation, "提示！"
    Exit Sub
End Sub
'Invoice
Public Sub InputInvoiceExportPrintExcel(ByVal rs As ADODB.Recordset, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double '总金额
Dim strExtsion          As String '后缀名
Dim strNewFullPath      As String '新Excel文件
Dim strXH               As String '序号
    
    If rs.RecordCount <= 0 Then
        MsgBox "没有要导出的资料！", vbInformation, "提示！"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "初始化数据..."
    
    strFileName = DirShare & "\output_invoice.xls" '要打开的文件
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
    strXH = Mid$(strExName, InStrRev(strExName, "_"))    '获取序号
    rs.MoveFirst    '数据集移动到第一个
    
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '是否显示
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        '赋值到Excel中，表头
        wkst.Cells(8, 11) = Format(DATE, "YYYY/mm/DD")
        wkst.Cells(9, 11) = "HTKS-SEDC" & Format(DATE, "YY-MMDD") & strXH
        
        lngRows = 17
        IntInertRow = rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = rs.RecordCount
        
'        ClsP.ShowProgress 50, "正在导出..."
        
        Do While Not rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & rs.Fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & rs.Fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & rs.Fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & rs.Fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & rs.Fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & rs.Fields(10).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(11).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(12).Value)
            
            DblNum = DblNum + Val(Trim$("" & rs.Fields(10).Value))
            DblAmt = DblAmt + Val(Trim$("" & rs.Fields(12).Value))
            
            
            lngRows = lngRows + 1
            rs.MoveNext
        Loop
        
        wkst.Cells(lngRows, 9) = DblNum
        wkst.Cells(lngRows, 11) = DblAmt
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
'        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
'        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
'        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
'        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.DESCRIPTION, vbInformation, "提示！"
    Exit Sub
End Sub

'Daily_inventory
Public Sub Daily_InvExportPrintExcel(ByVal rs As ADODB.Recordset, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double '总金额
Dim strExtsion          As String '后缀名
Dim strNewFullPath      As String '新Excel文件

    
    If rs.RecordCount <= 0 Then
        MsgBox "没有要导出的资料！", vbInformation, "提示！"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "初始化数据..."
    
    strFileName = DirShare & "\Daily_inventory_report.xls" '要打开的文件
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
    rs.MoveFirst    '数据集移动到第一个
    
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '是否显示
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        
        lngRows = 3
        IntInertRow = rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = rs.RecordCount
        
'        ClsP.ShowProgress 50, "正在导出..."
        
        Do While Not rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & rs.Fields(1).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & rs.Fields(2).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(3).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & rs.Fields(6).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & rs.Fields(7).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & rs.Fields(8).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & rs.Fields(9).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(10).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(11).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(12).Value)
            
            lngRows = lngRows + 1
            rs.MoveNext
        Loop
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
'        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
'        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
'        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
'        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.DESCRIPTION, vbInformation, "提示！"
    Exit Sub
End Sub

'shipping Packing list
Public Sub ShippingPackinglistExportPrintExcel(ByVal ordertemp As String, _
                                               ByVal strExName As String, lxflag As Integer)

    Dim strSql         As String

    Dim lngRows        As Long

    Dim rsQuery        As Excel.QueryTable

    Dim ExApp          As Excel.Application

    Dim wkbk           As New Workbook

    Dim wkst           As New Worksheet

    Dim i              As Long

    Dim j              As Long

    Dim IntCols        As Integer

    Dim strCols        As String

    Dim strFileName    As String

    Dim IntInertRow    As Integer, IntMaxDetailRow As Integer

    Dim DblNum         As Double

    Dim DblAmt         As Double  '总金额

    Dim intBoxNum      As Integer '箱数

    Dim strPBigBox     As String  '前箱号

    Dim strNBigBox     As String  '新箱号

    Dim IntBMegerRow   As Integer

    Dim IntEMegerRow   As Integer

    Dim DblJZ          As Double   '净重

    Dim DblMZ          As Double   '毛重

    Dim DblJZ1         As Double   '净重

    Dim DblMZ1         As Double   '毛重

    Dim DblJZ2         As Double   '净重

    Dim DblMZ2         As Double   '毛重

    Dim intBegin       As Integer

    Dim strdjTmp       As String

    Dim SD             As String

    Dim SD1            As String

    Dim strTmp()       As String

    Dim strExtsion     As String '后缀名

    Dim strNewFullPath As String '新Excel文件
    
    Dim strNewFullPathNew As String
    
    Dim DirFileShare1 As String

    Dim RsNew          As New ADODB.Recordset  '记录大箱的个数，方便后面计算体积重

    Dim rs             As New ADODB.Recordset

    Dim dnnum          As String

    Dim dnnum1         As String

    strNewFullPathNew = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\excel"
    
    dnnum = ""
    dnnum1 = ""
    
    strPBigBox = ""
    strNBigBox = ""
    strdjTmp = ""
    intBoxNum = 1
    '    If Rs.RecordCount <= 0 Then
    '        MsgBox "没有要导出的资料！", vbInformation, "提示！"
    '        Exit Sub
    '    End If
    '    ClsP.Init 100, True
    '    ClsP.ShowProgress 10, "初始化数据..."
    
    If lxflag = 0 Then
         strFileName = DirShare & "\shipping_packing_list.xlsx" '要打开的文件
         strExtsion = ".xls"
         strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
         strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径

    Else
         strFileName = DirShare & "\shipping_packing_list_2.xlsx" '要打开的文件
         strExtsion = ".xls"
         DirFileShare1 = "C:\老版_2"
         strNewFullPath = DirFileShare1 & "\" & strExName & strExtsion    '获取新文件要保存的路径
         If Dir("C:\老版_2", vbDirectory) = "" Then '判断文件夹是否存在
            MkDir ("C:\老版_2") '创建文件夹 msgbox ("创建完毕")
            MsgBox ("文件夹已创建！路径为 C:\老版_2")
         End If
    End If
    
    '    strSql = "SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
    '                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
    '                 ",箱号,料号,mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,净重,毛重,尺寸 " & _
    '                 " FROM Vw_InvShippedPLFor37 a  where 单据编号 in ('" & Ordertemp & "')  order by 箱号"

    strSql = " select * from ( SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & ",箱号,料号,replace(mpn_desc,'.P2','') AS mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,净重,毛重,尺寸 " & " FROM Vw_InvShippedPLFor37_NEW a  where 单据编号 in ('" & ordertemp & "')  " & " union all " & "SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & ",箱号,料号,replace(mpn_desc,'.P2','') AS mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,净重,毛重,尺寸 " & " FROM Vw_InvShippedPLFor37 a  where 单据编号 in ('" & ordertemp & "') ) x order by x.箱号 "
    



        Dim strDNList As String
        With lstLotID

            For i = 0 To .ListCount - 1
        
                If .Selected(i) = True Then
                    strDNList = strDNList & Trim$("" & .List(i)) & "','"
                     

                End If

            Next

        End With

        
strDNList = Mid(strDNList, 1, Len(strDNList) - 3)

'    strSql = "SELECT 0 选择,h.代码简称 + a.单据编号 单据编号,c.delivery,dbo.usp_date(a.单据日期) 发货日期,ISNULL(dn_address_new, c.shiptoname) AS shiptoname,ISNULL(x.ship_to_street1_new, c.shiptostreet1) AS shiptostreet1, " & _
'"ISNULL(x.ship_to_street2_new, c.shiptostreet2) AS shiptostreet2,ISNULL(x.ship_to_street3_new, c.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, c.city) AS city, ISNULL(x.dn_st_new, c.State) AS state, " & _
'"ISNULL(x.postal_code_new, c.postalcode) AS postalcode,ISNULL(x.country_new, c.countrykey) AS countrykey,ISNULL(x.contact_new, c.contactname) AS contactname,ISNULL(x.phone_new, c.phone) AS phone, " & _
'"c.SalesDocument,'''' + c.PurchasingDocNo AS PurchasingDocNo,RTRIM(erpdata.dbo.f_getparent(b.箱号)) 箱号,b.料号, " & _
'"CASE WHEN RTRIM(gg.MPN_DESC) = 'UCLAMP0541Z.TFT' THEN RTRIM(gg.MPN_DESC) + '.P2' ELSE REPLACE(REPLACE(gg.MPN_DESC, '.P2', ''), '.P3', '') END AS mpn_desc,SUM(b.数量) 数量,c.batchnumber, " & _
'"hh.CREATE_DATE DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no,c.customerPartNumber,ROUND(CAST(f.重量 AS FLOAT) * 0.4, 2) 净重,f.重量 毛重,f.尺寸 FROM erpdata .. tblStockSQfh a " & _
'"INNER JOIN erpdata .. tblStocksqfhsub b ON a.单据编号 = b.单据编号 AND a.序号 = b.单据项次 INNER JOIN erpdata .. tblStockNumTree g ON g.箱号 = b.箱号 INNER JOIN erpdata .. tblStockNumTree f " & _
'"ON f.序号 = g.上级序号 INNER JOIN (SELECT a.BOX_ID,SUBSTRING(a.KEY_VALUE,1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox,SUBSTRING(a.KEY_VALUE,CHARINDEX('|', a.KEY_VALUE) + 1,10) AS job " & _
'"FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON g.箱号 = aa.qbox INNER JOIN (SELECT dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, " & _
'"dn.shiptostreet3, dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber " & _
'"FROM tblCustomerShippingUp dn GROUP BY dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname, " & _
'"dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,dn.customerPartNumber) c ON c.Delivery = g.DN AND c.BatchNumber = aa.job INNER JOIN dbo.tblstock h " & _
'"ON CONVERT(NVARCHAR(4), h.库房代码) = CONVERT(NVARCHAR(4), a.仓库编号) INNER JOIN ERPBASE .. tblmappingData ff ON ff.SUBSTRATEID = b.流程卡编号 " & _
'"INNER JOIN ERPBASE .. tblCustomerOI gg ON CONVERT(VARCHAR(100), gg.ID) = ff.FILENAME AND gg.SOURCE_BATCH_ID = ff.LOTID AND gg.CUSTOMERSHORTNAME = '37' INNER JOIN erpbase .. weight37 hh " & _
'"ON hh.WAFERID = REPLACE(b.流程卡编号, '+', '') INNER JOIN erpdata .. tblErpInStockMainData jj ON jj.BOX_ID = aa.BOX_ID LEFT JOIN erptemp .. dn_address x ON dn_address = c.ShipToName " & _
'"WHERE a.客户代码 = '37' AND a.单据编号 LIKE 'F%' AND a.单据日期 >= CONVERT(VARCHAR(100), GETDATE() - 5, 23) AND c.Delivery = g.DN " & _
'"and c.Delivery in ('" & strDNList & "') GROUP BY h.代码简称, a.单据编号, c.delivery,dbo.usp_date(a.单据日期), c.shiptoname,c.shiptostreet1,c.shiptostreet2,c.shiptostreet3,c.city,c.State,c.postalcode, " & _
'"c.countrykey,c.contactname,c.phone,c.SalesDocument,c.PurchasingDocNo,erpdata.dbo.f_getparent(b.箱号),b.料号,gg.MPN_DESC,c.batchnumber,hh.CREATE_DATE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2), " & _
'"c.customerPartNumber,f.重量,f.尺寸,dn_address_new,x.ship_to_street1_new,x.ship_to_street2_new,x.ship_to_street3_new,x.city_new,x.dn_st_new,x.postal_code_new,x.country_new,x.contact_new,x.phone_new order by Delivery"

 strSql = "SELECT 0 AS  选择,y.代码简称 + b.单据编号 AS 单据编号,d.Delivery, CONVERT(VARCHAR(100), c.单据日期,23) AS 发货日期, ISNULL(dn_address_new, d.shiptoname) AS shiptoname, ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1 " & _
 " ,  ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2, ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,  ISNULL(x.city_new, d.city) AS city,  ISNULL(x.dn_st_new, d.State) AS state,  ISNULL(x.postal_code_new, d.postalcode) AS postalcode, " & _
 "  ISNULL(x.country_new, d.countrykey) AS countrykey,  ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone, d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo " & _
 " ,f.箱号 ,b.料号 ,d.MarketingPN,SUM(b.数量),d.BatchNumber,d.DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no, d.customerPartNumber, ROUND(CAST(f.重量 AS FLOAT) * 0.4, 2) 净重, f.重量 毛重, " & _
 " f.尺寸 FROM erpdata..tblStockNumTree a INNER JOIN erpdata..tblStocksqfhsub b ON b.箱号 = a.箱号 INNER JOIN erpdata..tblStockSQfh c ON c.单据编号 = b.单据编号 AND c.序号 = b.单据项次 INNER JOIN erpdata..tblstock y " & _
 " ON y.库房代码 = c.仓库编号  INNER JOIN (SELECT a.BOX_ID,  SUBSTRING(a.KEY_VALUE,  1, CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job " & _
 " FROM erpdata .. tblErpInStockDetailInfo a  WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.箱号 = aa.qbox  INNER JOIN (SELECT dn.Delivery, dn.shiptoname, dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3 " & _
 " , dn.city, dn.State, dn.postalcode, dn.countrykey, dn.contactname,  dn.phone, dn.SalesDocument,  dn.PurchasingDocNo, dn.BatchNumber, SUM(dn.Quantity) AS Quantity, dn.customerPartNumber, dn.MarketingPN, dn.DATE_CODE " & _
 " FROM ERPBASE..tblCustomerShippingUp dn WHERE dn.Delivery in ('" & strDNList & "') GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1, dn.shiptostreet2, dn.shiptostreet3,dn.city, dn.State, dn.postalcode, dn.countrykey " & _
 " , dn.contactname,dn.phone, dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber, dn.customerPartNumber, dn.MarketingPN,dn.DATE_CODE) d  ON d.Delivery = a.DN AND d.BatchNumber = aa.job " & _
 " INNER JOIN erpdata..tblStockNumTree e ON e.箱号 = b.箱号 INNER JOIN erpdata..tblstocknumtree f  ON f.序号 = e.上级序号 LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName " & _
"  WHERE a.DN IN ('" & strDNList & "')GROUP BY  b.单据编号,c.单据日期,ISNULL(dn_address_new, d.shiptoname) , ISNULL(x.ship_to_street1_new, d.shiptostreet1) ,ISNULL(x.ship_to_street2_new, d.shiptostreet2) " & _
 " ,ISNULL(x.ship_to_street3_new, d.shiptostreet3),ISNULL(x.city_new, d.city),ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode), ISNULL(x.country_new, d.countrykey) " & _
 " ,ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) , d.SalesDocument,d.PurchasingDocNo ,f.箱号,b.料号 ,d.MarketingPN,d.BatchNumber ,d.DATE_CODE " & _
 " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber, f.重量 , f.尺寸,y.代码简称,d.Delivery order by Delivery,箱号"

    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
    '    '-------------RS判定和筛选--------------------------------------------
    '    If InStr(strdjbh, ",") > 0 Then
    '        strTmp = Split(strdjbh, ",")
    '        For i = 0 To UBound(strTmp)
    '            strdjTmp = strdjTmp + "单据编号='" + strTmp(i) + "' OR "
    '        Next
    '        strdjTmp = Mid$(strdjTmp, 1, Len(strdjTmp) - 5)
    '    Else
    '        strdjTmp = "单据编号='" & strdjbh & "'"
    '    End If
    '    Rs.Filter = Trim(strdjTmp)          '数据集筛选
    '    Rs.Sort = "单据编号,箱号 ASC"       '数据集排序

    '    Rs.MoveFirst    '数据集移动到第一个
    '    '---------------------------------------------------------------------
    If rs.RecordCount > 0 Then
        '        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '是否显示
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
        '        ExApp.ActiveWindow.DisplayGridlines = False
        
        ' wkbk.ActiveSheet.Range(A3).Select
        
        DblNum = 0
        DblJZ = 0
        DblMZ = 0
        '

        '赋值到Excel中，表头
        'wkst.Cells(8, 2) = Trim$("" & rs.Fields(2).Value)
        wkst.Cells(8, 17) = Trim$("" & rs.Fields(3).Value)
        'shipto  销售装箱单时才执行 需求Sold to = Ship to 2019-12-5
        If lxflag = 1 Then
            wkst.Cells(10, 2) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(11, 2) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(12, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
            wkst.Cells(13, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
            wkst.Cells(14, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
            wkst.Cells(15, 2) = ""
        End If
        'sold
        wkst.Cells(17, 2) = Trim$("" & rs.Fields(4).Value)
        wkst.Cells(18, 2) = Trim$("" & rs.Fields(5).Value)
        wkst.Cells(19, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
        wkst.Cells(20, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
        wkst.Cells(22, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
        wkst.Cells(23, 2) = ""
        wkst.Cells(17, 17) = Trim$("" & rs.Fields(11).Value) 'To
        wkst.Cells(25, 3) = Trim$("" & rs.Fields(14).Value)
        wkst.Cells(25, 6) = Trim$("'" & rs.Fields(15).Value)
        
        lngRows = 28
        
        IntInertRow = rs.RecordCount * 2

        For i = 1 To IntInertRow - 1
            wkst.Rows(lngRows & ":" & lngRows).Select
            ExApp.Selection.Copy
            ExApp.Selection.Insert Shift:=xlDown
            wkst.Rows(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '边框无
        Next i

        IntMaxDetailRow = rs.RecordCount
        
        '        ClsP.ShowProgress 50, "正在导出..."
        
        IntBMegerRow = 27
        IntEMegerRow = 30
        intBegin = 1

        Dim QBX As String

        For i = 0 To rs.RecordCount - 1

            '            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '箱号
            If dnnum1 <> Trim$("" & rs.Fields(2).Value) Then
                dnnum = Trim$("" & rs.Fields(2).Value) + "/" + dnnum
                dnnum1 = Trim$("" & rs.Fields(2).Value)

            End If
             
            strPBigBox = Trim$("" & rs.Fields(16).Value) '箱号

            'QBX = strPBigBox
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & rs.Fields(16).Value) '箱号
                '箱数
                intBoxNum = intBoxNum + 1
                wkst.Cells(lngRows, 1) = "K" & Trim(intBoxNum - 1) '箱号进行转换为客户所要资料
                
                IntBMegerRow = IntBMegerRow + intBegin
                intBegin = 1
            Else
                '                '合并
                '                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).Merge
                '                '设定水平和竖直居中
                '                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).HorizontalAlignment = xlCenter
                '                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).VerticalAlignment = xlCenter
                '--------------------------
                intBegin = intBegin + 1

            End If
            
            If SD <> Trim$("" & rs.Fields(14).Value) Then
                SD = Trim$("" & rs.Fields(14).Value)
                SD1 = SD1 & SD & " "

            End If

            wkst.Cells(25, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & rs.Fields(19).Value)) / 1000 '数量改为已千为单位
            DblNum = DblNum + Val(Trim$("" & rs.Fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(21).Value) 'datacode
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(22).Value) 'lotno

            If strPBigBox <> QBX Then
                wkst.Cells(lngRows, 14) = Trim$("" & rs.Fields(24).Value) '净重
                wkst.Cells(lngRows, 15) = "KG"   '净重单位
                wkst.Cells(lngRows, 18) = "KG"   '毛重单位
                wkst.Cells(lngRows, 19) = Trim$("" & rs.Fields(26).Value)   '尺寸
                wkst.Cells(lngRows, 17) = Trim$("" & rs.Fields(25).Value)   '毛重
            
            End If
           
            DblJZ1 = Val(Trim$("" & rs.Fields(24).Value))
            
            If strPBigBox <> QBX Then
                DblJZ = DblJZ1 + DblJZ

            End If

            DblMZ1 = Val(Trim$("" & rs.Fields(25).Value))

            If strPBigBox <> QBX Then
                DblMZ = DblMZ + DblMZ1

            End If

            '
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(23).Value)
            
            QBX = strPBigBox
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            rs.MoveNext
        Next
        '计算汇总
        wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '总数量改为已千为单位
        wkst.Cells(lngRows + 1, 9) = "KPCS" '单位
        wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)    '箱数
        wkst.Cells(lngRows + 1, 14) = Format(DblJZ, "0.00") '净重
        wkst.Cells(lngRows + 1, 17) = DblMZ '毛重，记录它到后面进行对比
    Else
        '        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub

    End If

    '查询箱号尺寸，计算体积重
    Dim strXHCC As String       '箱数和尺寸

    Dim DblTJZ  As String       '体积重

    Dim order   As String
    
    order = Replace(ordertemp, "A", "")
    order = Replace$(order, "B", "")
    
    strXHCC = ""
    DblTJZ = 0
    'strdjbh1 = Mid(strdjbh, 2, Len(strdjbh) - 1)
    '    strSql = "SELECT COUNT(DISTINCT erpdata.dbo.f_getparent(b.箱号)) 箱数,c.尺寸 " & _
    '             " FROM erpdata..tblStockMove a " & _
    '             " INNER JOIN erpdata..tblStockMovesub b ON a.单据编号 = b.单据编号 AND a.序号=b.单据项次 " & _
    '             " INNER JOIN erpdata..tblStockNumTree c On c.箱号=erpdata.dbo.f_getparent(b.箱号) " & _
    '             " WHERE a.单据编号 IN ('" & order & "')" & _
    '             " GROUP BY c.尺寸"
             
    strSql = "SELECT  COUNT(DISTINCT d.箱号) 箱数,d.尺寸  " & " FROM erpdata..tblStockSQfh  a  " & "  INNER JOIN erpdata..tblStockSQfhsub b ON a.单据编号 = b.单据编号 AND a.序号=b.单据项次 " & "   INNER JOIN erpdata..tblStockNumTree c On c.箱号=b.箱号 AND c.基层标记 = 0 " & "   INNER JOIN erpdata..tblStockNumTree d On d.序号 = c.上级序号 AND d.基层标记 = 1 " & " WHERE a.单据编号 IN ('" & order & "','') GROUP BY d.尺寸 "
             
    
    If RsNew.State = adStateOpen Then RsNew.Close
    RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If Not RsNew.EOF Then

        Do While Not RsNew.EOF
            '循环得到箱号的箱数和尺寸，进行拼接
            strXHCC = strXHCC & Trim$("" & RsNew!箱数) & "@" & Trim$("" & RsNew!尺寸) & "cm;"

            '对尺寸进行分割，计算体积重
            If Trim$("" & RsNew!尺寸) <> "" And InStr(Trim$("" & RsNew!尺寸), "*") > 0 Then
                strTmp = Split(Trim$("" & RsNew!尺寸), "*") '分割字符
                '计算体积重并汇总
                DblTJZ = DblTJZ + Val(Trim$("" & RsNew!箱数)) * strTmp(0) * strTmp(1) * strTmp(2) / 5000

            End If

            RsNew.MoveNext
        Loop

    End If

    RsNew.Close
    
    wkst.Cells(8, 2) = Mid(dnnum, 1, Len(dnnum) - 1)
    
    wkst.Cells(8, 5) = ""
    
   ' wkst.Cells(7, 10).Select
    
    ' 生成二维码
    Dim strQrCodePath As String
    
    strNewFullPathNew = strNewFullPathNew & "\" & strExName & strExtsion
    strNewFullPathNew = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\excel" & "\" & strExName & strExtsion
    strQrCodePath = DirQrShare & "\" & strExName & ".JPG"
    strQrCodePath = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\jpg" & "\" & strExName & ".JPG"
    test.Visible = False

    test.QRmaker1.InputData = wkst.Cells(8, 2)
    test.QRmaker1.Refresh
    test.QRmaker1.CreateQrMetaFile hDC, strQrCodePath, 2
    Unload test

    'wkst.Pictures.Insert (App.Path + "\dn.bmp")
    wkst.Shapes.AddPicture _
    strQrCodePath, _
    True, True, 1100, 200, 400, 400
    
    '赋值体积重
    wkst.Cells(lngRows + 3, 4) = Format(DblTJZ, "0.00")

    '比较体积重和毛重看哪个大,就取哪个
    If DblMZ > DblTJZ Then
        wkst.Cells(lngRows + 3, 11) = Format(DblMZ, "0.00")
    Else
        wkst.Cells(lngRows + 3, 11) = Format(DblTJZ, "0.00")

    End If

    '赋值到EXCEL箱数和尺寸
    wkst.Cells(lngRows + 4, 3) = strXHCC

    '    ExApp.Columns.AutoFit '自适应列宽
    With wkst.PageSetup

        '        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
        '        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
        '        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
        '        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
        '        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        ' .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With

    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else

            On Error Resume Next

            Kill strNewFullPath

            If Err.number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub

            End If

        End If

    End If

    ' wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    ' wkbk.Saved = True
    
    'wkbk.SaveAs strNewFullPathNew, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPathNew
    wkbk.Saved = True
    
    
    
    '---------------------------------------------------------------------------------------------------------------
    '    ClsP.ShowProgress 100, "导出成功！"
    '    If Not ClsP Is Nothing Then
    '        Set ClsP = Nothing
    '    End If
    ExApp.Visible = True
    
    '    If intFlag = 1 Then
    '        wkst.PrintPreview
    '        wkbk.Close (False)
    '        ExApp.Quit
    '    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing

    End If

    Exit Sub
ErrHandle:

    On Error Resume Next

    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing

    End If

    '    If Not ClsP Is Nothing Then
    '        Set ClsP = Nothing
    '    End If
    MsgBox Err.DESCRIPTION, vbInformation, "提示！"
    Exit Sub


End Sub

'DN地址相同合并导出
Public Sub ShippingPackinglistExportPrintExcel2(ByVal strExName As String, lxflag As Integer)

   Dim strDNList As String  '确定DN范围
   Dim i As Integer
   Dim ShiptoName As String
   

        With lstLotID
            For i = 0 To .ListCount - 1
                If .Selected(i) = True Then
                    strDNList = strDNList & Trim$("" & .List(i)) & "','"
                    
                End If

            Next

        End With

 
 strDNList = Mid(strDNList, 1, Len(strDNList) - 3)
        
        
   With Fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 5
                If Trim(.text) <> ShiptoName Then
                    ShiptoName = .text
                    strExName = strDNList
                   Call SPPE2(strDNList, ShiptoName, strExName, lxflag)
                End If
            End If
            Next
        
    End With
    MsgBox "导出完成！"
End Sub

'DN地址相同合并导出子函数
Public Sub SPPE2(strDNList As String, ShiptoName As String, strExName As String, lxflag As Integer)
    Dim strSql         As String

    Dim lngRows        As Long

    Dim rsQuery        As Excel.QueryTable

    Dim ExApp          As Excel.Application

    Dim wkbk           As New Workbook

    Dim wkst           As New Worksheet

    Dim i              As Long

    Dim j              As Long

    Dim IntCols        As Integer

    Dim strCols        As String

    Dim strFileName    As String

    Dim IntInertRow    As Integer, IntMaxDetailRow As Integer

    Dim DblNum         As Double

    Dim DblAmt         As Double  '总金额

    Dim intBoxNum      As Integer '箱数

    Dim strPBigBox     As String  '前箱号

    Dim strNBigBox     As String  '新箱号

    Dim IntBMegerRow   As Integer

    Dim IntEMegerRow   As Integer

    Dim DblJZ          As Double   '净重

    Dim DblMZ          As Double   '毛重

    Dim DblJZ1         As Double   '净重

    Dim DblMZ1         As Double   '毛重

    Dim DblJZ2         As Double   '净重

    Dim DblMZ2         As Double   '毛重

    Dim intBegin       As Integer

    Dim strdjTmp       As String

    Dim SD             As String

    Dim SD1            As String

    Dim strTmp()       As String

    Dim strExtsion     As String '后缀名

    Dim strNewFullPath As String '新Excel文件
    
    Dim strNewFullPathNew As String
    
    Dim order1 As String
    
    ' Dim strExName As String
    
    Dim ordertemp  As String
    Dim ordertemp1  As String

    Dim RsNew          As New ADODB.Recordset  '记录大箱的个数，方便后面计算体积重

    Dim rs             As New ADODB.Recordset

    Dim dnnum          As String

    Dim dnnum1         As String
    Dim DirFileShare1 As String
    
    order1 = ""
    dnnum = ""
    dnnum1 = ""
    ordertemp = ""
    ordertemp1 = ""
    strPBigBox = ""
    strNBigBox = ""
    strdjTmp = ""
    intBoxNum = 1
    '    If Rs.RecordCount <= 0 Then
    '        MsgBox "没有要导出的资料！", vbInformation, "提示！"
    '        Exit Sub
    '    End If
    '    ClsP.Init 100, True
    '    ClsP.ShowProgress 10, "初始化数据..."
    

 
     If lxflag = 0 Then
        strNewFullPathNew = "C:\合并_spl"
        strFileName = DirShare & "\shipping_packing_list.xlsx" '要打开的文件
        strExtsion = ".xlsx"
        DirFileShare1 = "C:\合并_spl"
        strNewFullPath = DirFileShare1 & "\" & strExName & strExtsion    '获取新文件要保存的路径
        If Dir("C:\合并_spl", vbDirectory) = "" Then '判断文件夹是否存在
            MkDir ("C:\合并_spl") '创建文件夹 msgbox ("创建完毕")
            MsgBox ("文件夹已创建！路径为 C:\合并_spl")
        Else
            'MsgBox ("文件夹已在")
        End If
  
    Else
        strNewFullPathNew = "C:\合并_spl2"
        strFileName = DirShare & "\shipping_packing_list_2.xlsx" '要打开的文件
        strExtsion = ".xlsx"
        DirFileShare1 = "C:\合并_spl2"
        strNewFullPath = DirFileShare1 & "\" & strExName & strExtsion    '获取新文件要保存的路径
        If Dir("C:\合并_spl2", vbDirectory) = "" Then '判断文件夹是否存在
            MkDir ("C:\合并_spl2") '创建文件夹 msgbox ("创建完毕")
            MsgBox ("文件夹已创建！路径为 C:\合并_spl2")
        Else
            'MsgBox ("文件夹已在")
        End If
    End If
    
    
     
    'strExName = GetExcelName(Trim(Fra(1).Caption))

    '    strSql = "SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
    '                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
    '                 ",箱号,料号,mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,净重,毛重,尺寸 " & _
    '                 " FROM Vw_InvShippedPLFor37 a  where 单据编号 in ('" & Ordertemp & "')  order by 箱号"

 strSql = "select aaa.* from (SELECT 0 AS  选择,y.代码简称 + b.单据编号 AS 单据编号,d.Delivery, CONVERT(VARCHAR(100), c.单据日期,23) AS 发货日期, ISNULL(dn_address_new, d.shiptoname) AS shiptoname, ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1 " & _
 " ,  ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2, ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,  ISNULL(x.city_new, d.city) AS city,  ISNULL(x.dn_st_new, d.State) AS state,  ISNULL(x.postal_code_new, d.postalcode) AS postalcode, " & _
 "  ISNULL(x.country_new, d.countrykey) AS countrykey,  ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone, d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo " & _
 " ,f.箱号 ,b.料号 ,d.MarketingPN,SUM(b.数量) as sum,d.BatchNumber,d.DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no, d.customerPartNumber, ROUND(CAST(f.重量 AS FLOAT) * 0.4, 2) 净重, f.重量 毛重, " & _
 " f.尺寸 FROM erpdata..tblStockNumTree a INNER JOIN erpdata..tblStocksqfhsub b ON b.箱号 = a.箱号 INNER JOIN erpdata..tblStockSQfh c ON c.单据编号 = b.单据编号 AND c.序号 = b.单据项次 INNER JOIN erpdata..tblstock y " & _
 " ON y.库房代码 = c.仓库编号  INNER JOIN (SELECT a.BOX_ID,  SUBSTRING(a.KEY_VALUE,  1, CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job " & _
 " FROM erpdata .. tblErpInStockDetailInfo a  WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.箱号 = aa.qbox  INNER JOIN (SELECT dn.Delivery, dn.shiptoname, dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3 " & _
 " , dn.city, dn.State, dn.postalcode, dn.countrykey, dn.contactname,  dn.phone, dn.SalesDocument,  dn.PurchasingDocNo, dn.BatchNumber, SUM(dn.Quantity) AS Quantity, dn.customerPartNumber, dn.MarketingPN, dn.DATE_CODE " & _
 " FROM ERPBASE..tblCustomerShippingUp dn WHERE dn.Delivery in ('" & strDNList & "') GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1, dn.shiptostreet2, dn.shiptostreet3,dn.city, dn.State, dn.postalcode, dn.countrykey " & _
 " , dn.contactname,dn.phone, dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber, dn.customerPartNumber, dn.MarketingPN,dn.DATE_CODE) d  ON d.Delivery = a.DN AND d.BatchNumber = aa.job " & _
 " INNER JOIN erpdata..tblStockNumTree e ON e.箱号 = b.箱号 INNER JOIN erpdata..tblstocknumtree f  ON f.序号 = e.上级序号 LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName " & _
"  WHERE a.DN IN ('" & strDNList & "') GROUP BY  b.单据编号,c.单据日期,ISNULL(dn_address_new, d.shiptoname) , ISNULL(x.ship_to_street1_new, d.shiptostreet1) ,ISNULL(x.ship_to_street2_new, d.shiptostreet2) " & _
 " ,ISNULL(x.ship_to_street3_new, d.shiptostreet3),ISNULL(x.city_new, d.city),ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode), ISNULL(x.country_new, d.countrykey) " & _
 " ,ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) , d.SalesDocument,d.PurchasingDocNo ,f.箱号,b.料号 ,d.MarketingPN,d.BatchNumber ,d.DATE_CODE " & _
 " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber, f.重量 , f.尺寸,y.代码简称,d.Delivery ) aaa where aaa.shiptoname = '" & ShiptoName & "' order by aaa.shiptoname, aaa.Delivery,aaa.箱号 "


    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
    '    '-------------RS判定和筛选--------------------------------------------
    '    If InStr(strdjbh, ",") > 0 Then
    '        strTmp = Split(strdjbh, ",")
    '        For i = 0 To UBound(strTmp)
    '            strdjTmp = strdjTmp + "单据编号='" + strTmp(i) + "' OR "
    '        Next
    '        strdjTmp = Mid$(strdjTmp, 1, Len(strdjTmp) - 5)
    '    Else
    '        strdjTmp = "单据编号='" & strdjbh & "'"
    '    End If
    '    Rs.Filter = Trim(strdjTmp)          '数据集筛选
    '    Rs.Sort = "单据编号,箱号 ASC"       '数据集排序
   
   If Not rs.EOF Then

        Do While Not rs.EOF
            If ordertemp <> Trim$(rs.Fields("单据编号")) Then
                ordertemp = Trim$(rs.Fields("单据编号"))
                ordertemp1 = ordertemp1 & ordertemp & "','"
            End If
            rs.MoveNext
        Loop

   End If
        'MsgBox "" & ordertemp1
        
     rs.MoveFirst    '数据集移动到第一个
    '    '---------------------------------------------------------------------
    If rs.RecordCount > 0 Then
            
        '        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '是否显示
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
        '        ExApp.ActiveWindow.DisplayGridlines = False
        
        ' wkbk.ActiveSheet.Range(A3).Select
        
        DblNum = 0
        DblJZ = 0
        DblMZ = 0
        '

        '赋值到Excel中，表头
        'wkst.Cells(8, 2) = Trim$("" & rs.Fields(2).Value)
        wkst.Cells(8, 17) = Trim$("" & rs.Fields(3).Value)
        'shipto  销售装箱单时才执行 需求Sold to = Ship to 2019-12-5
        If lxflag = 1 Then
            wkst.Cells(10, 2) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(11, 2) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(12, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
            wkst.Cells(13, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
            wkst.Cells(14, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
            wkst.Cells(15, 2) = ""
        End If
        
        wkst.Cells(17, 2) = Trim$("" & rs.Fields(4).Value)
        wkst.Cells(18, 2) = Trim$("" & rs.Fields(5).Value)
        wkst.Cells(19, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
        wkst.Cells(20, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
        wkst.Cells(22, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
        If UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION (CAMARILLO)" Or UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION" Then
        
           wkst.Cells(23, 2) = "TAX ID: 95-2119684"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INTERCONNECT" Then
        
           wkst.Cells(23, 2) = "TAX ID: 82-5035949"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INCORPORATED(FEDERAL）" Then
        
            wkst.Cells(23, 2) = "TAX ID: 82-5035949"
        
        Else
           
            wkst.Cells(23, 2) = ""
        
        End If
        
     '   wkst.Cells(23, 2) = ""
        wkst.Cells(17, 17) = Trim$("" & rs.Fields(11).Value) 'To
        wkst.Cells(25, 3) = Trim$("" & rs.Fields(14).Value)
        wkst.Cells(25, 6) = Trim$("'" & rs.Fields(15).Value)
        
        lngRows = 28
        
        IntInertRow = rs.RecordCount * 2

        For i = 1 To IntInertRow - 1
            
            wkst.Rows(lngRows & ":" & lngRows).Select
            ExApp.Selection.Copy
            ExApp.Selection.Insert Shift:=xlDown
            wkst.Rows(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '边框无
        Next i

        IntMaxDetailRow = rs.RecordCount
        
        '        ClsP.ShowProgress 50, "正在导出..."
        
        IntBMegerRow = 27
        IntEMegerRow = 30
        intBegin = 1

        Dim QBX As String

        For i = 0 To rs.RecordCount - 1

            '            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '箱号
            If dnnum1 <> Trim$("" & rs.Fields(2).Value) Then
                dnnum = Trim$("" & rs.Fields(2).Value) + "/" + dnnum
                dnnum1 = Trim$("" & rs.Fields(2).Value)

            End If
             
            strPBigBox = Trim$("" & rs.Fields(16).Value) '箱号

            'QBX = strPBigBox
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & rs.Fields(16).Value) '箱号
                '箱数
                intBoxNum = intBoxNum + 1
                wkst.Cells(lngRows, 1) = "K" & Trim(intBoxNum - 1) '箱号进行转换为客户所要资料
                
                IntBMegerRow = IntBMegerRow + intBegin
                intBegin = 1
            Else
                '                '合并
                '                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).Merge
                '                '设定水平和竖直居中
                '                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).HorizontalAlignment = xlCenter
                '                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).VerticalAlignment = xlCenter
                '--------------------------
                intBegin = intBegin + 1

            End If
            
            If SD <> Trim$("" & rs.Fields(14).Value) Then
                SD = Trim$("" & rs.Fields(14).Value)
                SD1 = SD1 & SD & " "

            End If

            wkst.Cells(25, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & rs.Fields(19).Value)) / 1000 '数量改为已千为单位
            DblNum = DblNum + Val(Trim$("" & rs.Fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(21).Value) 'datacode
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(22).Value) 'lotno

            If strPBigBox <> QBX Then
                
                If Trim$("" & rs.Fields(24).Value) = "" Or Trim$("" & rs.Fields(25).Value) = "" Or Trim$("" & rs.Fields(26).Value) = "" Then
                    
                    MsgBox "没有维护重量尺寸", vbInformation, "提示"
                    Exit Sub
                    
                End If
                
                wkst.Cells(lngRows, 14) = Trim$("" & rs.Fields(24).Value) '净重
                wkst.Cells(lngRows, 15) = "KG"   '净重单位
                wkst.Cells(lngRows, 18) = "KG"   '毛重单位
                wkst.Cells(lngRows, 19) = Trim$("" & rs.Fields(26).Value)   '尺寸
                wkst.Cells(lngRows, 17) = Trim$("" & rs.Fields(25).Value)   '毛重
            
            End If
           
            DblJZ1 = Val(Trim$("" & rs.Fields(24).Value))
            
            If strPBigBox <> QBX Then
                DblJZ = DblJZ1 + DblJZ

            End If

            DblMZ1 = Val(Trim$("" & rs.Fields(25).Value))

            If strPBigBox <> QBX Then
                DblMZ = DblMZ + DblMZ1

            End If

            '
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(23).Value)
            
            QBX = strPBigBox
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            rs.MoveNext
        Next
        '计算汇总
        wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '总数量改为已千为单位
        wkst.Cells(lngRows + 1, 9) = "KPCS" '单位
        wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)    '箱数
        wkst.Cells(lngRows + 1, 14) = Format(DblJZ, "0.00") '净重
        wkst.Cells(lngRows + 1, 17) = DblMZ '毛重，记录它到后面进行对比
    Else
        '        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub

    End If

    '查询箱号尺寸，计算体积重
    Dim strXHCC As String       '箱数和尺寸

    Dim DblTJZ  As String       '体积重

   ' Dim order   As String
    ordertemp1 = Mid(ordertemp1, 1, Len(ordertemp1) - 3)
    order1 = Replace(ordertemp1, "A", "")
    order1 = Replace$(order1, "B", "")
    
    strXHCC = ""
    DblTJZ = 0
    'strdjbh1 = Mid(strdjbh, 2, Len(strdjbh) - 1)
    '    strSql = "SELECT COUNT(DISTINCT erpdata.dbo.f_getparent(b.箱号)) 箱数,c.尺寸 " & _
    '             " FROM erpdata..tblStockMove a " & _
    '             " INNER JOIN erpdata..tblStockMovesub b ON a.单据编号 = b.单据编号 AND a.序号=b.单据项次 " & _
    '             " INNER JOIN erpdata..tblStockNumTree c On c.箱号=erpdata.dbo.f_getparent(b.箱号) " & _
    '             " WHERE a.单据编号 IN ('" & order & "')" & _
    '             " GROUP BY c.尺寸"
             
    strSql = "SELECT  COUNT(DISTINCT d.箱号) 箱数,d.尺寸  " & " FROM erpdata..tblStockSQfh  a  " & "  INNER JOIN erpdata..tblStockSQfhsub b ON a.单据编号 = b.单据编号 AND a.序号=b.单据项次 " & "   INNER JOIN erpdata..tblStockNumTree c On c.箱号=b.箱号 AND c.基层标记 = 0 " & "   INNER JOIN erpdata..tblStockNumTree d On d.序号 = c.上级序号 AND d.基层标记 = 1 " & " WHERE a.单据编号 IN ('" & order1 & "') GROUP BY d.尺寸 "
             
    
    If RsNew.State = adStateOpen Then RsNew.Close
    RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If Not RsNew.EOF Then

        Do While Not RsNew.EOF
            '循环得到箱号的箱数和尺寸，进行拼接
            strXHCC = strXHCC & Trim$("" & RsNew!箱数) & "@" & Trim$("" & RsNew!尺寸) & "cm;"

            '对尺寸进行分割，计算体积重
            If Trim$("" & RsNew!尺寸) <> "" And InStr(Trim$("" & RsNew!尺寸), "*") > 0 Then
                strTmp = Split(Trim$("" & RsNew!尺寸), "*") '分割字符
                '计算体积重并汇总
                DblTJZ = DblTJZ + Val(Trim$("" & RsNew!箱数)) * strTmp(0) * strTmp(1) * strTmp(2) / 5000

            End If

            RsNew.MoveNext
        Loop

    End If

    RsNew.Close
    
    wkst.Cells(8, 2) = Mid(dnnum, 1, Len(dnnum) - 1)
    
    wkst.Cells(8, 5) = ""
    
   ' wkst.Cells(7, 10).Select
    
    ' 生成二维码
    Dim strQrCodePath As String
    
    strNewFullPathNew = strNewFullPathNew & "\" & strExName & strExtsion
    'strNewFullPathNew = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\excel" & "\" & strExName & strExtsion
    strQrCodePath = DirQrShare & "\" & strExName & ".JPG"
    strQrCodePath = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\jpg" & "\" & strExName & ".JPG"
    test.Visible = False

    test.QRmaker1.InputData = wkst.Cells(8, 2)
    test.QRmaker1.Refresh
    test.QRmaker1.CreateQrMetaFile hDC, strQrCodePath, 2
    Unload test

    'wkst.Pictures.Insert (App.Path + "\dn.bmp")
    wkst.Shapes.AddPicture _
    strQrCodePath, _
    True, True, 1100, 200, 400, 400
    
    '赋值体积重
    wkst.Cells(lngRows + 3, 4) = Format(DblTJZ, "0.00")

    '比较体积重和毛重看哪个大,就取哪个
    If DblMZ > DblTJZ Then
        wkst.Cells(lngRows + 3, 11) = Format(DblMZ, "0.00")
    Else
        wkst.Cells(lngRows + 3, 11) = Format(DblTJZ, "0.00")

    End If

    '赋值到EXCEL箱数和尺寸
    wkst.Cells(lngRows + 4, 3) = strXHCC

    '    ExApp.Columns.AutoFit '自适应列宽
    With wkst.PageSetup

        '        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
        '        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
        '        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
        '        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
        '        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        ' .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With

    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else

            On Error Resume Next

            Kill strNewFullPath

            If Err.number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub

            End If

        End If

    End If

    ' wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    ' wkbk.Saved = True
    
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
    wkbk.Saved = True
    
    
    
    '---------------------------------------------------------------------------------------------------------------
    '    ClsP.ShowProgress 100, "导出成功！"
    '    If Not ClsP Is Nothing Then
    '        Set ClsP = Nothing
    '    End If
    ExApp.Visible = True
    
    '    If intFlag = 1 Then
    '        wkst.PrintPreview
    '        wkbk.Close (False)
    '        ExApp.Quit
    '    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing

    End If

    Exit Sub
ErrHandle:

    On Error Resume Next

    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing

    End If

    '    If Not ClsP Is Nothing Then
    '        Set ClsP = Nothing
    '    End If
    MsgBox Err.DESCRIPTION, vbInformation, "提示！执行完毕"
    Exit Sub
End Sub

'单项DN逐项导出
Public Sub ShippingPackinglistExportPrintExcel1(ByVal strExName As String, lxflag As Integer)

    Dim i              As Long

    Dim j              As Long
    Dim strDNList      As String
    

        

    
        With lstLotID

            For i = 0 To .ListCount - 1
        
                If .Selected(i) = True Then
                    'strDNList = strDNList & Trim$("" & .List(i)) & "','"
                    strDNList = Trim(.List(i))
                    strExName = strDNList
                    Call SPEPrintExcel1(strDNList, strExName, lxflag)
                    
                End If

            Next

        End With
    MsgBox "导出完成！"
        
         'strDNList = Mid(strDNList, 1, Len(strDNList) - 3)
End Sub


'单项DN逐项导出循环遍历子函数
Private Function SPEPrintExcel1(strDNList As String, strExName As String, lxflag As Integer)
    Dim strSql         As String

    Dim lngRows        As Long

    Dim rsQuery        As Excel.QueryTable

    Dim ExApp          As Excel.Application

    Dim wkbk           As New Workbook

    Dim wkst           As New Worksheet

    Dim i              As Long

    Dim j              As Long

    Dim IntCols        As Integer

    Dim strCols        As String

    Dim strFileName    As String

    Dim IntInertRow    As Integer, IntMaxDetailRow As Integer

    Dim DblNum         As Double

    Dim DblAmt         As Double  '总金额

    Dim intBoxNum      As Integer '箱数

    Dim strPBigBox     As String  '前箱号

    Dim strNBigBox     As String  '新箱号

    Dim IntBMegerRow   As Integer

    Dim IntEMegerRow   As Integer

    Dim DblJZ          As Double   '净重

    Dim DblMZ          As Double   '毛重

    Dim DblJZ1         As Double   '净重

    Dim DblMZ1         As Double   '毛重

    Dim DblJZ2         As Double   '净重

    Dim DblMZ2         As Double   '毛重

    Dim intBegin       As Integer

    Dim strdjTmp       As String

    Dim SD             As String

    Dim SD1            As String

    Dim strTmp()       As String

    Dim strExtsion     As String '后缀名

    Dim strNewFullPath As String '新Excel文件
    
    Dim strNewFullPathNew As String
    
    'Dim strExName As String
    
    Dim ordertemp As String
    Dim ordertemp1 As String
    
    Dim order1 As String

    Dim RsNew          As New ADODB.Recordset  '记录大箱的个数，方便后面计算体积重

    Dim rs             As New ADODB.Recordset

    Dim dnnum          As String

    Dim dnnum1         As String
    Dim DirFileShare1 As String

    dnnum = ""
    dnnum1 = ""
    ordertemp = ""
    ordertemp1 = ""
    order1 = ""
    
    strPBigBox = ""
    strNBigBox = ""
    strdjTmp = ""
    intBoxNum = 1
    '    If Rs.RecordCount <= 0 Then
    '        MsgBox "没有要导出的资料！", vbInformation, "提示！"
    '        Exit Sub
    '    End If
    '    ClsP.Init 100, True
    '    ClsP.ShowProgress 10, "初始化数据..."
       
    strNewFullPathNew = "C:\单项"
    
    If lxflag = 0 Then
         strFileName = DirShare & "\shipping_packing_list.xlsx" '要打开的文件
         strExtsion = ".xlsx"
         DirFileShare1 = "C:\单项"
         strNewFullPath = DirFileShare1 & "\" & strExName & strExtsion    '获取新文件要保存的路径
         If Dir("C:\单项", vbDirectory) = "" Then '判断文件夹是否存在
            MkDir ("C:\单项") '创建文件夹 msgbox ("创建完毕")
            MsgBox ("文件夹已创建！路径为 C:\单项")
         End If

   Else
         strFileName = DirShare & "\shipping_packing_list_2.xlsx" '要打开的文件
         strExtsion = ".xlsx"
         DirFileShare1 = "C:\单项_2"
         strNewFullPath = DirFileShare1 & "\" & strExName & strExtsion    '获取新文件要保存的路径
         If Dir("C:\单项_2", vbDirectory) = "" Then '判断文件夹是否存在
            MkDir ("C:\单项_2") '创建文件夹 msgbox ("创建完毕")
            MsgBox ("文件夹已创建！路径为 C:\单项_2")
         End If
   End If
    'strExName = GetExcelName(Trim(Fra(1).Caption))

    '    strSql = "SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
    '                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
    '                 ",箱号,料号,mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,净重,毛重,尺寸 " & _
    '                 " FROM Vw_InvShippedPLFor37 a  where 单据编号 in ('" & Ordertemp & "')  order by 箱号"


 strSql = "SELECT 0 AS  选择,y.代码简称 + b.单据编号 AS 单据编号,d.Delivery, CONVERT(VARCHAR(100), c.单据日期,23) AS 发货日期, ISNULL(dn_address_new, d.shiptoname) AS shiptoname, ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1 " & _
 " ,  ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2, ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,  ISNULL(x.city_new, d.city) AS city,  ISNULL(x.dn_st_new, d.State) AS state,  ISNULL(x.postal_code_new, d.postalcode) AS postalcode, " & _
 "  ISNULL(x.country_new, d.countrykey) AS countrykey,  ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone, d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo " & _
 " ,f.箱号 ,b.料号 ,d.MarketingPN,SUM(b.数量),d.BatchNumber,d.DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no, d.customerPartNumber, ROUND(CAST(f.重量 AS FLOAT) * 0.4, 2) 净重, f.重量 毛重, " & _
 " f.尺寸 FROM erpdata..tblStockNumTree a INNER JOIN erpdata..tblStocksqfhsub b ON b.箱号 = a.箱号 INNER JOIN erpdata..tblStockSQfh c ON c.单据编号 = b.单据编号 AND c.序号 = b.单据项次 INNER JOIN erpdata..tblstock y " & _
 " ON y.库房代码 = c.仓库编号  INNER JOIN (SELECT a.BOX_ID,  SUBSTRING(a.KEY_VALUE,  1, CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job " & _
 " FROM erpdata .. tblErpInStockDetailInfo a  WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.箱号 = aa.qbox  INNER JOIN (SELECT dn.Delivery, dn.shiptoname, dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3 " & _
 " , dn.city, dn.State, dn.postalcode, dn.countrykey, dn.contactname,  dn.phone, dn.SalesDocument,  dn.PurchasingDocNo, dn.BatchNumber, SUM(dn.Quantity) AS Quantity, dn.customerPartNumber, dn.MarketingPN, dn.DATE_CODE " & _
 " FROM ERPBASE..tblCustomerShippingUp dn WHERE dn.Delivery in ('" & strDNList & "') GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1, dn.shiptostreet2, dn.shiptostreet3,dn.city, dn.State, dn.postalcode, dn.countrykey " & _
 " , dn.contactname,dn.phone, dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber, dn.customerPartNumber, dn.MarketingPN,dn.DATE_CODE) d  ON d.Delivery = a.DN AND d.BatchNumber = aa.job " & _
 " INNER JOIN erpdata..tblStockNumTree e ON e.箱号 = b.箱号 INNER JOIN erpdata..tblstocknumtree f  ON f.序号 = e.上级序号 LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName " & _
 "  WHERE a.DN IN ('" & strDNList & "')GROUP BY  b.单据编号,c.单据日期,ISNULL(dn_address_new, d.shiptoname) , ISNULL(x.ship_to_street1_new, d.shiptostreet1) ,ISNULL(x.ship_to_street2_new, d.shiptostreet2) " & _
 " ,ISNULL(x.ship_to_street3_new, d.shiptostreet3),ISNULL(x.city_new, d.city),ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode), ISNULL(x.country_new, d.countrykey) " & _
 " ,ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) , d.SalesDocument,d.PurchasingDocNo ,f.箱号,b.料号 ,d.MarketingPN,d.BatchNumber ,d.DATE_CODE " & _
 " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber, f.重量 , f.尺寸,y.代码简称,d.Delivery order by Delivery,箱号"

    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
    '    '-------------RS判定和筛选--------------------------------------------
    '    If InStr(strdjbh, ",") > 0 Then
    '        strTmp = Split(strdjbh, ",")
    '        For i = 0 To UBound(strTmp)
    '            strdjTmp = strdjTmp + "单据编号='" + strTmp(i) + "' OR "
    '        Next
    '        strdjTmp = Mid$(strdjTmp, 1, Len(strdjTmp) - 5)
    '    Else
    '        strdjTmp = "单据编号='" & strdjbh & "'"
    '    End If
    '    Rs.Filter = Trim(strdjTmp)          '数据集筛选
    '    Rs.Sort = "单据编号,箱号 ASC"       '数据集排序
    
    'strExtsion = Mid$(StrFileName, InStrRev(StrFileName, "."))      '获取后缀名

    '找该DN的全部单据编号
       If Not rs.EOF Then

        Do While Not rs.EOF
            If ordertemp <> Trim$(rs.Fields("单据编号")) Then
                ordertemp = Trim$(rs.Fields("单据编号"))
                ordertemp1 = ordertemp1 & ordertemp & "','"
            End If
            rs.MoveNext
        Loop

       End If
    
    '如果当前DN没有数据，就执行下一个DN导出，单项导出会有这个问题
    If rs.RecordCount = 0 Then
        Exit Function
    End If
        
     rs.MoveFirst    '数据集移动到第一个
     
    If rs.RecordCount > 0 Then
        '        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '是否显示
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
        '        ExApp.ActiveWindow.DisplayGridlines = False
        
        ' wkbk.ActiveSheet.Range(A3).Select
        
       
        
        DblNum = 0
        DblJZ = 0
        DblMZ = 0
        '

        '赋值到Excel中，表头
        'wkst.Cells(8, 2) = Trim$("" & rs.Fields(2).Value)
        wkst.Cells(8, 17) = Trim$("" & rs.Fields(3).Value)
        'shipto  销售装箱单时才执行 需求Sold to = Ship to 2019-12-5
        If lxflag = 1 Then
            wkst.Cells(10, 2) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(11, 2) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(12, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
            wkst.Cells(13, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
            wkst.Cells(14, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
            wkst.Cells(15, 2) = ""
        End If
        'soldto
        wkst.Cells(17, 2) = Trim$("" & rs.Fields(4).Value)
        wkst.Cells(18, 2) = Trim$("" & rs.Fields(5).Value)
        wkst.Cells(19, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
        wkst.Cells(20, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
        wkst.Cells(22, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)


        If UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION (CAMARILLO)" Or UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION" Then
        
           wkst.Cells(23, 2) = "TAX ID: 95-2119684"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INTERCONNECT" Then
        
           wkst.Cells(23, 2) = "TAX ID: 82-5035949"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INCORPORATED(FEDERAL）" Then
        
            wkst.Cells(23, 2) = "TAX ID: 82-5035949"
        
        Else
           
        wkst.Cells(23, 2) = ""
        
        End If
        wkst.Cells(17, 17) = Trim$("" & rs.Fields(11).Value) 'To
        wkst.Cells(25, 3) = Trim$("" & rs.Fields(14).Value)
        wkst.Cells(25, 6) = Trim$("'" & rs.Fields(15).Value)
        
        lngRows = 28
        
        IntInertRow = rs.RecordCount * 2

        For i = 1 To IntInertRow - 1
            wkst.Rows(lngRows & ":" & lngRows).Select
            ExApp.Selection.Copy
            ExApp.Selection.Insert Shift:=xlDown
            wkst.Rows(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '边框无
        Next i

        IntMaxDetailRow = rs.RecordCount
        
        '        ClsP.ShowProgress 50, "正在导出..."
        
        IntBMegerRow = 27
        IntEMegerRow = 30
        intBegin = 1

        Dim QBX As String

        For i = 0 To rs.RecordCount - 1

            '            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '箱号
            If dnnum1 <> Trim$("" & rs.Fields(2).Value) Then
                dnnum = Trim$("" & rs.Fields(2).Value) + "/" + dnnum
                dnnum1 = Trim$("" & rs.Fields(2).Value)

            End If
             
            strPBigBox = Trim$("" & rs.Fields(16).Value) '箱号

            'QBX = strPBigBox
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & rs.Fields(16).Value) '箱号
                '箱数
                intBoxNum = intBoxNum + 1
                wkst.Cells(lngRows, 1) = "K" & Trim(intBoxNum - 1) '箱号进行转换为客户所要资料
                
                IntBMegerRow = IntBMegerRow + intBegin
                intBegin = 1
            Else
                '                '合并
                '                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).Merge
                '                '设定水平和竖直居中
                '                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).HorizontalAlignment = xlCenter
                '                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).VerticalAlignment = xlCenter
                '--------------------------
                intBegin = intBegin + 1

            End If
            
            If SD <> Trim$("" & rs.Fields(14).Value) Then
                SD = Trim$("" & rs.Fields(14).Value)
                SD1 = SD1 & SD & " "

            End If

            wkst.Cells(25, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & rs.Fields(19).Value)) / 1000 '数量改为已千为单位
            DblNum = DblNum + Val(Trim$("" & rs.Fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(21).Value) 'datacode
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(22).Value) 'lotno

            If strPBigBox <> QBX Then
                wkst.Cells(lngRows, 14) = Trim$("" & rs.Fields(24).Value) '净重
                wkst.Cells(lngRows, 15) = "KG"   '净重单位
                wkst.Cells(lngRows, 18) = "KG"   '毛重单位
                wkst.Cells(lngRows, 19) = Trim$("" & rs.Fields(26).Value)   '尺寸
                wkst.Cells(lngRows, 17) = Trim$("" & rs.Fields(25).Value)   '毛重
            
            End If
           
            DblJZ1 = Val(Trim$("" & rs.Fields(24).Value))
            
            If strPBigBox <> QBX Then
                DblJZ = DblJZ1 + DblJZ

            End If

            DblMZ1 = Val(Trim$("" & rs.Fields(25).Value))

            If strPBigBox <> QBX Then
                DblMZ = DblMZ + DblMZ1

            End If

            '
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(23).Value)
            
            QBX = strPBigBox
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            rs.MoveNext
        Next
        '计算汇总
        wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '总数量改为已千为单位
        wkst.Cells(lngRows + 1, 9) = "KPCS" '单位
        wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)    '箱数
        wkst.Cells(lngRows + 1, 14) = Format(DblJZ, "0.00") '净重
        wkst.Cells(lngRows + 1, 17) = DblMZ '毛重，记录它到后面进行对比
    Else
        '        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Function

    End If

    '查询箱号尺寸，计算体积重
    Dim strXHCC As String       '箱数和尺寸

    Dim DblTJZ  As String       '体积重

   ' Dim order   As String
    ordertemp1 = Mid(ordertemp1, 1, Len(ordertemp1) - 3)
    order1 = Replace(ordertemp1, "A", "")
    order1 = Replace$(order1, "B", "")
    
    strXHCC = ""
    DblTJZ = 0
    'strdjbh1 = Mid(strdjbh, 2, Len(strdjbh) - 1)
    '    strSql = "SELECT COUNT(DISTINCT erpdata.dbo.f_getparent(b.箱号)) 箱数,c.尺寸 " & _
    '             " FROM erpdata..tblStockMove a " & _
    '             " INNER JOIN erpdata..tblStockMovesub b ON a.单据编号 = b.单据编号 AND a.序号=b.单据项次 " & _
    '             " INNER JOIN erpdata..tblStockNumTree c On c.箱号=erpdata.dbo.f_getparent(b.箱号) " & _
    '             " WHERE a.单据编号 IN ('" & order & "')" & _
    '             " GROUP BY c.尺寸"
             
    strSql = "SELECT  COUNT(DISTINCT d.箱号) 箱数,d.尺寸  " & " FROM erpdata..tblStockSQfh  a  " & "  INNER JOIN erpdata..tblStockSQfhsub b ON a.单据编号 = b.单据编号 AND a.序号=b.单据项次 " & "   INNER JOIN erpdata..tblStockNumTree c On c.箱号=b.箱号 AND c.基层标记 = 0 " & "   INNER JOIN erpdata..tblStockNumTree d On d.序号 = c.上级序号 AND d.基层标记 = 1 " & " WHERE a.单据编号 IN ('" & order1 & "','') GROUP BY d.尺寸 "
             
    
    If RsNew.State = adStateOpen Then RsNew.Close
    RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

    If Not RsNew.EOF Then

        Do While Not RsNew.EOF
            '循环得到箱号的箱数和尺寸，进行拼接
            strXHCC = strXHCC & Trim$("" & RsNew!箱数) & "@" & Trim$("" & RsNew!尺寸) & "cm;"

            '对尺寸进行分割，计算体积重
            If Trim$("" & RsNew!尺寸) <> "" And InStr(Trim$("" & RsNew!尺寸), "*") > 0 Then
                strTmp = Split(Trim$("" & RsNew!尺寸), "*") '分割字符
                '计算体积重并汇总
                DblTJZ = DblTJZ + Val(Trim$("" & RsNew!箱数)) * strTmp(0) * strTmp(1) * strTmp(2) / 5000

            End If

            RsNew.MoveNext
        Loop

    End If

    RsNew.Close
    
    wkst.Cells(8, 2) = Mid(dnnum, 1, Len(dnnum) - 1)
    
    wkst.Cells(8, 5) = ""
    
   ' wkst.Cells(7, 10).Select
    
    ' 生成二维码
    Dim strQrCodePath As String
    
    strNewFullPathNew = strNewFullPathNew & "\" & strExName & strExtsion
    'strNewFullPathNew = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\excel" & "\" & strExName & strExtsion
    strQrCodePath = DirQrShare & "\" & strExName & ".JPG"
    strQrCodePath = "\\10.160.1.84\public\FileServer\35.市场部订单维护系统\37\jpg" & "\" & strExName & ".JPG"
    test.Visible = False

    test.QRmaker1.InputData = wkst.Cells(8, 2)
    test.QRmaker1.Refresh
    test.QRmaker1.CreateQrMetaFile hDC, strQrCodePath, 2
    Unload test

    'wkst.Pictures.Insert (App.Path + "\dn.bmp")
    wkst.Shapes.AddPicture _
    strQrCodePath, _
    True, True, 1100, 200, 400, 400
    
    '赋值体积重
    wkst.Cells(lngRows + 3, 4) = Format(DblTJZ, "0.00")

    '比较体积重和毛重看哪个大,就取哪个
    If DblMZ > DblTJZ Then
        wkst.Cells(lngRows + 3, 11) = Format(DblMZ, "0.00")
    Else
        wkst.Cells(lngRows + 3, 11) = Format(DblTJZ, "0.00")

    End If

    '赋值到EXCEL箱数和尺寸
    wkst.Cells(lngRows + 4, 3) = strXHCC

    '    ExApp.Columns.AutoFit '自适应列宽
    With wkst.PageSetup

        '        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
        '        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
        '        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
        '        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
        '        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        ' .RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With

    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Function
        Else

            On Error Resume Next

            Kill strNewFullPath

            If Err.number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Function

            End If

        End If

    End If

    ' wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    ' wkbk.Saved = True
    
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
    wkbk.Saved = True
    
    
    
    '---------------------------------------------------------------------------------------------------------------
    '    ClsP.ShowProgress 100, "导出成功！"
    '    If Not ClsP Is Nothing Then
    '        Set ClsP = Nothing
    '    End If
    ExApp.Visible = True
    
    '    If intFlag = 1 Then
    '        wkst.PrintPreview
    '        wkbk.Close (False)
    '        ExApp.Quit
    '    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing

    End If

    Exit Function
ErrHandle:

    On Error Resume Next

    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing

    End If

    '    If Not ClsP Is Nothing Then
    '        Set ClsP = Nothing
    '    End If
    MsgBox Err.DESCRIPTION, vbInformation, "提示！执行完毕"
    Exit Function


End Function
'shipping invoice
Public Sub ShippingInvoiceExportPrintExcel(ByVal ordertemp As String, ByVal strExName As String)
    Dim strSql              As String
    Dim lngRows             As Long
    Dim rsQuery             As Excel.QueryTable
    Dim ExApp               As Excel.Application
    Dim wkbk                As New Workbook
    Dim wkst                As New Worksheet
    Dim i                   As Long
    Dim j                   As Long
    Dim IntCols             As Integer
    Dim strCols             As String
    Dim strFileName         As String
    Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
    Dim DblNum              As Double
    Dim DblAmt              As Double  '总金额
    Dim intBoxNum           As Integer '箱数
    Dim strPBigBox          As String  '前箱号
    Dim strNBigBox          As String  '新箱号
    Dim IntBMegerRow        As Integer
    Dim IntEMegerRow        As Integer
    Dim DblJZ               As Double   '净重
    Dim DblMZ               As Double   '毛重
    Dim intBegin            As Integer
    Dim strdjTmp            As String
    Dim strTmp()            As String
    Dim SD                  As String
    Dim SD1                  As String
    Dim strExtsion          As String '后缀名
    Dim strNewFullPath      As String '新Excel文件
    Dim rs               As New ADODB.Recordset
    Dim dnnum            As String
    Dim dnnum1            As String
    Dim jine1 As Double
    Dim jine2 As Double
     
    dnnum = ""
    dnnum1 = ""
    strPBigBox = ""
    strNBigBox = ""
    strdjTmp = ""
    intBoxNum = 1
'
'    If Rs.RecordCount <= 0 Then
'        MsgBox "没有要导出的资料！", vbInformation, "提示！"
'        Exit Sub
'    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "初始化数据..."
    
    strFileName = DirShare & "\shipping_invoice.xlsx" '要打开的文件
    
       Dim strDNList As String
        With lstLotID

            For i = 0 To .ListCount - 1
        
                If .Selected(i) = True Then
                    strDNList = strDNList & Trim$("" & .List(i)) & "','"

                End If
                
            Next

        End With

        
          strDNList = Mid(strDNList, 1, Len(strDNList) - 3)

                    
'    strSql = " SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
'                 " ,city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
'                 " ,箱号,料号,mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,单价,AMount,customerPartNumber " & _
'                 "  FROM Vw_InvShippedInvoiceFor37 a  where 单据编号 in ('" & Ordertemp & "')  order by 箱号  "


' strSql = "   select * from ( SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
'                 " ,city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
'                 " ,箱号,料号,replace(mpn_desc,'.P2','') as mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,单价,AMount,customerPartNumber,加工费金额,客供料金额  " & _
'                 "  FROM Vw_InvShippedInvoiceFor37_NEW a  where 单据编号 in ('" & Ordertemp & "')    " & _
'                 "union all " & _
'                 " SELECT 0 选择,单据编号,delivery,发货日期,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
'                 " ,city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
'                 " ,箱号,料号,replace(mpn_desc,'.P2','') as mpn_desc,数量,batchnumber,DATE_CODE,HTlot_no,单价,AMount,customerPartNumber ,加工费金额,客供料金额  " & _
'                 "  FROM Vw_InvShippedInvoiceFor37 a  where 单据编号 in ('" & Ordertemp & "') )  x order by x.箱号  "
'
    
'    strSql = "SELECT 0 选择, h.代码简称+a.单据编号 单据编号,c.delivery,dbo.usp_date(a.单据日期) 发货日期, ISNULL(dn_address_new, c.shiptoname) AS  shiptoname,ISNULL(x.ship_to_street1_new ,c.shiptostreet1)  AS shiptostreet1  " & _
'",ISNULL(x.ship_to_street2_new,c.shiptostreet2) AS shiptostreet2 ,ISNULL(x.ship_to_street3_new,c.shiptostreet3) AS shiptostreet3 ,ISNULL(x.city_new ,c.city) AS city,ISNULL(x.dn_st_new, c.State) AS state,ISNULL(x.postal_code_new,c.postalcode  ) AS  postalcode " & _
'",ISNULL(x.country_new,c.countrykey) AS countrykey,ISNULL(x.contact_new,c.contactname) AS contactname ,ISNULL(x.phone_new,c.phone ) AS phone      " & _
'",c.SalesDocument,'''' + c.PurchasingDocNo AS PurchasingDocNo,RTRIM(erpdata.dbo.f_getparent(b.箱号)) 箱号,b.料号,CASE WHEN RTRIM(gg.MPN_DESC)='UCLAMP0541Z.TFT' THEN RTRIM(gg.MPN_DESC)+'.P2' " & _
'"ELSE REPLACE(REPLACE(gg.MPN_DESC,'.P2',''),'.P3','') END AS mpn_desc,SUM(b.数量) 数量,c.batchnumber,hh.CREATE_DATE DATE_CODE " & _
'",SUBSTRING(aa.qbox ,2,CHARINDEX('-R',aa.qbox)-2)  HTlot_no ,ISNULL(ISNULL( BB.含税单价 / AB.良品数,0)  + ( cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0) AS 单价 " & _
'",ROUND( SUM(b.数量) * ISNULL(ISNULL( BB.含税单价 / AB.良品数,0)  +  (cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0),2) AS AMount,c.customerPartNumber    " & _
'",ROUND( SUM(b.数量) * ISNULL(ISNULL( BB.含税单价 / AB.良品数,0)  +  (cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0),2) -  CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( BB.含税单价 / AB.良品数,0)) AS 加工费金额 " & _
'", CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( BB.含税单价 / AB.良品数,0)) AS 客供料金额 FROM erpdata..tblStockSQfh a           " & _
'"INNER JOIN erpdata..tblStocksqfhsub b ON a.单据编号 = b.单据编号 AND a.序号=b.单据项次 INNER JOIN erpdata..tblStockNumTree g ON g.箱号=b.箱号 " & _
'"INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE,1,CHARINDEX('|',a.KEY_VALUE)-1) AS qbox , SUBSTRING(a.KEY_VALUE,CHARINDEX('|',a.KEY_VALUE)+1,10) AS job " & _
'" FROM erpdata..tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND  a.KEY_VALUE LIKE '%SS%|%')  aa ON g.箱号 = aa.qbox  " & _
'"INNER JOIN (SELECT dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode "
'
'
'     strSql = strSql & ",dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber FROM tblCustomerShippingUp dn GROUP BY dn.Delivery,dn.shiptoname,dn.shiptostreet1 " & _
'",dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,dn.customerPartNumber)  c ON c.Delivery = g.DN  AND c.BatchNumber =  aa.job " & _
'"INNER JOIN dbo.tblstock h ON CONVERT(NVARCHAR(4),h.库房代码) = CONVERT(NVARCHAR(4),a.仓库编号)    INNER JOIN ERPBASE..tblmappingData ff ON ff.SUBSTRATEID = b.流程卡编号 " & _
'"INNER JOIN ERPBASE..tblCustomerOI gg ON CONVERT(VARCHAR(30), gg.ID) = ff.FILENAME AND gg.SOURCE_BATCH_ID = ff.LOTID AND gg.CUSTOMERSHORTNAME = '37' " & _
'"INNER JOIN erpbase..weight37 hh ON hh.WAFERID = REPLACE(b.流程卡编号,'+','') INNER JOIN erpdata..tblErpInStockMainData jj ON jj.BOX_ID = aa.BOX_ID " & _
'"LEFT JOIN erpbase..tbltoinrec_wafer  AB ON ab.批号 = ff.LOTID AND AB.晶圆ID = REPLACE(B.流程卡编号,'+','') LEFT JOIN erpbase..tbltorec_wafer  ww ON  ww.批号 = ab.批号 AND  ww.晶圆ID = ab.晶圆ID  " & _
'"LEFT JOIN ERPBASE..TblToInsub BB ON BB.入库单编号 = AB.入库单编号 AND BB.到货批号 = AB.批号 AND ww.到货单编号 = bb.到货单编号 AND bb.含税单价 IS NOT NULL " & _
'"LEFT JOIN erptemp..tblBB_CSRPO cb ON cb.PO_NUM = gg.PO_NUM AND cb.FAB_DEVICE = gg.MPN_DESC LEFT JOIN ERPBASE..tblmappingData db ON db.SUBSTRATEID = REPLACE(B.流程卡编号,'+','')  " & _
'"LEFT JOIN  erpdata..tblSalerec e ON  e.单据编号 = a.单据编号       AND a.序号 = e.单据项次  AND e.小箱号 = b.箱号    " & _
'"LEFT JOIN erptemp..dn_address x ON dn_address = c.ShipToName WHERE a.客户代码='37' and c.Delivery in ('" & strDNList & "') AND a.单据日期 >= CONVERT(VARCHAR(100),GETDATE()- 8,23) AND  a.单据编号 LIKE 'F%' AND a.良品数量 >0  " & _
'"GROUP BY gg.PO_NUM,h.代码简称,a.单据编号,c.delivery,dbo.usp_date(a.单据日期),c.shiptoname,c.shiptostreet1,c.shiptostreet2         " & _
'",c.shiptostreet3,c.city,c.State,c.postalcode,c.countrykey,c.contactname,c.phone,c.SalesDocument,c.PurchasingDocNo ,erpdata.dbo.f_getparent(b.箱号),b.料号,gg.MPN_DESC,c.batchnumber,hh.CREATE_DATE      " & _
'",SUBSTRING(aa.qbox ,2,CHARINDEX('-R',aa.qbox)-2),e.销售单编号,c.customerPartNumber ,ISNULL( BB.含税单价 / AB.良品数,0) ,  cb.WAFER_PRICE,db.PASSBINCOUNT , cb.DIE_PRICE,dn_address_new " & _
'",x.ship_to_street1_new,x.ship_to_street2_new,x.ship_to_street3_new,x.city_new,x.dn_st_new,x.postal_code_new,x.country_new,x.contact_new ,x.phone_new  order by Delivery "


    strSql = " SELECT 0 AS 选择,y.代码简称 + b.单据编号 AS 单据编号,a.DN, CONVERT(VARCHAR(100), c.单据日期,23) AS 发货日期, ISNULL(dn_address_new, d.shiptoname) AS shiptoname " & _
 ",ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1,ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2 , ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, d.city) AS city, " & _
 " ISNULL(x.dn_st_new, d.State) AS state, ISNULL(x.postal_code_new, d.postalcode) AS postalcode,ISNULL(x.country_new, d.countrykey) AS countrykey, ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone,d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo, f.箱号, " & _
 " b.料号,d.MarketingPN,SUM(b.数量), d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no ,ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  + ( dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0) AS 单价 " & _
 " ,ROUND( SUM(b.数量) * ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  +  (dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0),2) AS AMount ,d.customerPartNumber ,ROUND( SUM(b.数量) * ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  +  (dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0),2) " & _
 "  -  CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( cd.含税单价 / cc.良品数,0)) AS 加工费金额 , CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( cd.含税单价 / cc.良品数,0)) AS 客供料金额 " & _
"   FROM erpdata .. tblStockNumTree a INNER JOIN erpdata .. tblStocksqfhsub b ON b.箱号 = a.箱号 INNER JOIN erpdata .. tblStockSQfh c  ON c.单据编号 = b.单据编号 AND c.序号 = b.单据项次 INNER JOIN erpdata..tblstock y ON y.库房代码 = c.仓库编号 " & _
"  INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE, 1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job  FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.箱号 = aa.qbox  INNER JOIN (SELECT dn.Delivery, " & _
 "  dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State,dn.postalcode, dn.countrykey, dn.contactname,dn.phone, dn.SalesDocument,  dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE  FROM ERPBASE .. tblCustomerShippingUp dn " & _
 "   WHERE dn.Delivery IN ('" & strDNList & "')  GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State, dn.postalcode,  dn.countrykey,dn.contactname,dn.phone, dn.SalesDocument, dn.PurchasingDocNo, dn.BatchNumber, dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE) d " & _
 "   ON d.Delivery = a.DN  AND d.BatchNumber = aa.job INNER JOIN erpdata .. tblStockNumTree e  ON e.箱号 = b.箱号 INNER JOIN erpdata .. tblstocknumtree f  ON f.序号 = e.上级序号  INNER JOIN ERPBASE..tblmappingData qq ON qq.SUBSTRATEID = b.流程卡编号 AND qq.LOTID = b.工单号 LEFT JOIN ERPBASE..tblCustomerOI bb " & _
 " ON CONVERT(VARCHAR(100), bb.ID) = qq.FILENAME AND bb.SOURCE_BATCH_ID = qq.LOTID  LEFT JOIN ERPBASE..tblToRec_Wafer cc ON cc.晶圆ID = REPLACE(qq.SUBSTRATEID,'+','')   AND cc.批号 = qq.LOTID  LEFT JOIN ERPBASE..tblToRecEntry cd ON cd.到货单编号 = cc.到货单编号 AND cd.到货批号 = cc.批号 LEFT JOIN erptemp..tblBB_CSRPO dd " & _
"  ON dd.PO_NUM = bb.PO_NUM AND dd.FAB_DEVICE = bb.MPN_DESC LEFT JOIN  erpdata..tblSalerec j ON j.单据编号 = c.单据编号 AND j.单据项次 = b.单据项次 AND j.小箱号 = b.箱号 LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName  WHERE a.DN IN ('" & strDNList & "')  GROUP BY b.单据编号, c.单据日期, ISNULL(dn_address_new, d.shiptoname), " & _
"  ISNULL(x.ship_to_street1_new, d.shiptostreet1), ISNULL(x.ship_to_street2_new, d.shiptostreet2),ISNULL(x.ship_to_street3_new, d.shiptostreet3), ISNULL(x.city_new, d.city),  ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode),ISNULL(x.country_new, d.countrykey),ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone), " & _
" d.SalesDocument, d.PurchasingDocNo,f.箱号, b.料号,d.MarketingPN,d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber,   a.DN , J.销售单编号, bb.PO_NUM, cd.含税单价, cc.良品数, dd.WAFER_PRICE, qq.PASSBINCOUNT, dd.die_price, y.代码简称 order by DN,箱号 "
  
     rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
'    '-------------RS判定和筛选--------------------------------------------
'    If InStr(strdjbh, ",") > 0 Then
'        strTmp = Split(strdjbh, ",")
'        For i = 0 To UBound(strTmp)
'            strdjTmp = strdjTmp + "单据编号='" + strTmp(i) + "' OR "
'        Next
'        strdjTmp = Mid$(strdjTmp, 1, Len(strdjTmp) - 5)
'    Else
'        strdjTmp = "单据编号='" & strdjbh & "'"
'    End If
'    Rs.Filter = Trim(strdjTmp)          '数据集筛选
'    Rs.Sort = "单据编号,箱号 ASC"       '数据集排序
   strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
   strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
'    Rs.MoveFirst    '数据集移动到第一个
'    '---------------------------------------------------------------------
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '是否显示
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
    
    
    
    
'    '-------------RS判定和筛选--------------------------------------------
'    If InStr(strdjbh, ",") > 0 Then
'        'strTmp = Split(strdjbh, ",")
'       ' For i = 0 To UBound(strTmp)
'            'strdjTmp = strdjTmp + "单据编号='" + strTmp(i) + "' AND "
'       ' Next
'       ' strdjTmp = Mid$(strdjTmp, 1, Len(strdjTmp) - 5)
'        MsgBox "无法同时导出多个单据号！"
'        Exit Sub
'    Else
'        strdjTmp = "单据编号='" & strdjbh & "'"
'    End If
'    Rs.Filter = Trim(strdjTmp)  '数据集筛选
'    Rs.Sort = "单据编号,箱号 ASC" '数据集排序
'    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '获取后缀名
'    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '获取新文件要保存的路径
'    Rs.MoveFirst    '数据集移动到第一个
'    '---------------------------------------------------------------------
'    If Rs.RecordCount > 0 Then
''        ClsP.ShowProgress 30, "初始化Excel..."
'        Set ExApp = New Excel.Application
'        ExApp.Visible = False   '是否显示
'
'        Set wkbk = ExApp.Workbooks.open(strFileName)
'        Set wkst = wkbk.Worksheets(1)
''        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        DblJZ = 0
        DblMZ = 0
        '赋值到Excel中，表头
        'wkst.Cells(13, 10) = Trim$("" & rs.Fields(2).Value)
        wkst.Cells(15, 10) = Trim$("" & rs.Fields(3).Value)
        wkst.Cells(18, 10) = Trim$("" & rs.Fields(11).Value) 'To
        
        wkst.Cells(13, 2) = Trim$("" & rs.Fields(4).Value)
        wkst.Cells(14, 2) = Trim$("" & rs.Fields(5).Value)
        wkst.Cells(15, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
        wkst.Cells(16, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
        
        wkst.Cells(18, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
        wkst.Cells(19, 2) = ""

        'wkst.Cells(23, 3) = Trim$("" & Rs.fields(14).Value)
        wkst.Cells(23, 5) = Trim$("'" & rs.Fields(15).Value)
        
        lngRows = 27
        
        IntInertRow = rs.RecordCount * 2
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
           wkst.Range(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '边框无
        Next i
        IntMaxDetailRow = rs.RecordCount
        
'        ClsP.ShowProgress 50, "正在导出..."
        
        IntBMegerRow = 26
        IntEMegerRow = 29
        intBegin = 1
        Dim QBX As String
        
        For i = 0 To rs.RecordCount - 1
'            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '箱号

             If dnnum1 <> Trim$("" & rs.Fields(2).Value) And InStr(dnnum, Trim$("" & rs.Fields(2).Value)) = 0 Then
                dnnum = Trim$("" & rs.Fields(2).Value) + "/" + dnnum
                dnnum1 = Trim$("" & rs.Fields(2).Value)
             End If
             

            
            strPBigBox = Trim$("" & rs.Fields(16).Value) '箱号
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & rs.Fields(16).Value) '箱号
                '箱数
                intBoxNum = intBoxNum + 1
                wkst.Cells(lngRows, 1) = "K" & Trim(intBoxNum - 1) '箱号进行转换为客户所要资料
                QBX = "K" & Trim(intBoxNum - 1)
                
                IntBMegerRow = IntBMegerRow + intBegin
                intBegin = 1
            Else
'                '合并
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).Merge
'                '设定水平和竖直居中
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).HorizontalAlignment = xlCenter
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).VerticalAlignment = xlCenter
                '--------------------------
                intBegin = intBegin + 1
            End If
              If SD <> Trim$("" & rs.Fields(14).Value) Then
             SD = Trim$("" & rs.Fields(14).Value)
             SD1 = SD1 & SD & " "
             End If
            wkst.Cells(23, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & rs.Fields(19).Value)) / 1000 '数量都为除以1000的值
            DblNum = DblNum + Val(Trim$("" & rs.Fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(21).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(22).Value)
            wkst.Cells(lngRows, 13) = "US$"
            wkst.Cells(lngRows, 14) = Val(Trim$("" & rs.Fields(23).Value)) * 1000 '单价为乘以1000的值
            wkst.Cells(lngRows, 15) = "US$"
            wkst.Cells(lngRows, 16) = Trim$("" & rs.Fields(24).Value)
            DblAmt = DblAmt + Val(Trim$("" & rs.Fields(24).Value))
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(25).Value)
            
            jine1 = jine1 + Val(Trim$("" & rs.Fields(26).Value))
            
            jine2 = jine2 + Val(Trim$("" & rs.Fields(27).Value))
            
            
            
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            rs.MoveNext
        Next
        
        wkst.Cells(13, 10) = Mid(dnnum, 1, Len(dnnum) - 1)
        
        '计算汇总
        wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '数量
        wkst.Cells(lngRows + 1, 9) = "KPCS" '单位
        wkst.Cells(lngRows + 1, 16) = DblAmt
        wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)
        
        wkst.Cells(lngRows + 8, 1) = "Process Amount：US$ " + Str(jine1)
        wkst.Cells(lngRows + 9, 1) = "Wafer Amount：US$ " + Str(jine2)

        
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Sub
    End If
'    ExApp.Columns.AutoFit '自适应列宽
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""楷体_GB2312,常?""&10報表名稱 ：   "  ' & Gsmc
'        .CenterHeader = "&""楷体_GB2312,常規""報表&""宋体,常規""" & Chr(10) & "&""楷体_GB2312,常?""&10日 期："
'        .RightHeader = "" & Chr(10) & "&""楷体_GB2312,常規""&10單位：滬士"
'        .LeftFooter = "&""楷体_GB2312,常規""&10制表人："
'        .CenterFooter = "&""楷体_GB2312,常規""&10制表日期："
        '.RightFooter = "&20" & "第 &P 页，共 &N 页"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Sub
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.DESCRIPTION, vbInformation, "提示！"
    Exit Sub
End Sub

'shipping invoice
Public Sub ShippingInvoiceExportPrintExcel2()

    Dim strDNList As String
    Dim ShiptoName As String
    Dim strExName As String
    Dim i As Integer

    If Dir("C:\合并_voice", vbDirectory) = "" Then '判断文件夹是否存在
        MkDir ("C:\合并_voice") '创建文件夹 msgbox ("创建完毕")
        MsgBox ("文件夹已创建！路径为 C:\合并_voice")
    Else
        'MsgBox ("文件夹已在")
    End If
    
        With lstLotID

            For i = 0 To .ListCount - 1
        
                If .Selected(i) = True Then
                    strDNList = strDNList & Trim$("" & .List(i)) & "','"

                End If
                
            Next

        End With
        strDNList = Mid(strDNList, 1, Len(strDNList) - 3)
        
      With Fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 5
                If Trim(.text) <> ShiptoName Then
                    ShiptoName = Trim(.text)
                    strExName = strDNList
                    Call SPLSVoice(strDNList, strExName, ShiptoName)
                End If
            End If
            Next
        
    End With
    MsgBox "导出完成！"
          

End Sub

Public Function SPLSVoice(strDNList As String, strExName As String, ShiptoName As String)
    Dim strSql              As String
    Dim lngRows             As Long
    Dim rsQuery             As Excel.QueryTable
    Dim ExApp               As Excel.Application
    Dim wkbk                As New Workbook
    Dim wkst                As New Worksheet
    Dim i                   As Long
    Dim j                   As Long
    Dim IntCols             As Integer
    Dim strCols             As String
    Dim strFileName         As String
    Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
    Dim DblNum              As Double
    Dim DblAmt              As Double  '总金额
    Dim intBoxNum           As Integer '箱数
    Dim strPBigBox          As String  '前箱号
    Dim strNBigBox          As String  '新箱号
    Dim IntBMegerRow        As Integer
    Dim IntEMegerRow        As Integer
    Dim DblJZ               As Double   '净重
    Dim DblMZ               As Double   '毛重
    Dim intBegin            As Integer
    Dim strdjTmp            As String
    Dim strTmp()            As String
    Dim SD                  As String
    Dim SD1                  As String
    Dim strExtsion          As String '后缀名
    Dim strNewFullPath      As String '新Excel文件
    Dim rs               As New ADODB.Recordset
    Dim dnnum            As String
    Dim dnnum1            As String
    Dim jine1 As Double
    Dim jine2 As Double
     
    dnnum = ""
    dnnum1 = ""
    strPBigBox = ""
    strNBigBox = ""
    strdjTmp = ""
    intBoxNum = 1
    
    strFileName = DirShare & "\shipping_invoice.xlsx" '要打开的文件

'    strsql = "select * from( SELECT 0 AS 选择,y.代码简称 + b.单据编号 AS 单据编号,a.DN, CONVERT(VARCHAR(100), c.单据日期,23) AS 发货日期, ISNULL(dn_address_new, d.shiptoname) AS shiptoname " & _
' ",ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1,ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2 , ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, d.city) AS city, " & _
' " ISNULL(x.dn_st_new, d.State) AS state, ISNULL(x.postal_code_new, d.postalcode) AS postalcode,ISNULL(x.country_new, d.countrykey) AS countrykey, ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone,d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo, f.箱号, " & _
' " b.料号,d.MarketingPN,SUM(b.数量) as 数量, d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no ,ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  + ( dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0) AS 单价 " & _
' " ,ROUND( SUM(b.数量) * ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  +  (dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0),2) AS AMount ,d.customerPartNumber ,ROUND( SUM(b.数量) * ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  +  (dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0),2) " & _
' "  -  CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( cd.含税单价 / cc.良品数,0)) AS 加工费金额 , CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( cd.含税单价 / cc.良品数,0)) AS 客供料金额 " & _
'"   FROM erpdata .. tblStockNumTree a INNER JOIN erpdata .. tblStocksqfhsub b ON b.箱号 = a.箱号 INNER JOIN erpdata .. tblStockSQfh c  ON c.单据编号 = b.单据编号 AND c.序号 = b.单据项次 INNER JOIN erpdata..tblstock y ON y.库房代码 = c.仓库编号 " & _
'"  INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE, 1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job  FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.箱号 = aa.qbox  INNER JOIN (SELECT dn.Delivery, " & _
' "  dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State,dn.postalcode, dn.countrykey, dn.contactname,dn.phone, dn.SalesDocument,  dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE  FROM ERPBASE .. tblCustomerShippingUp dn " & _
' "   WHERE dn.Delivery IN ('" & strDNList & "')  GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State, dn.postalcode,  dn.countrykey,dn.contactname,dn.phone, dn.SalesDocument, dn.PurchasingDocNo, dn.BatchNumber, dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE) d " & _
' "   ON d.Delivery = a.DN  AND d.BatchNumber = aa.job INNER JOIN erpdata .. tblStockNumTree e  ON e.箱号 = b.箱号 INNER JOIN erpdata .. tblstocknumtree f  ON f.序号 = e.上级序号  INNER JOIN ERPBASE..tblmappingData qq ON qq.SUBSTRATEID = b.流程卡编号 AND qq.LOTID = b.工单号 LEFT JOIN ERPBASE..tblCustomerOI bb " & _
' " ON CONVERT(VARCHAR(100), bb.ID) = qq.FILENAME AND bb.SOURCE_BATCH_ID = qq.LOTID  LEFT JOIN ERPBASE..tblToRec_Wafer cc ON cc.晶圆ID = REPLACE(qq.SUBSTRATEID,'+','')   AND cc.批号 = qq.LOTID  LEFT JOIN ERPBASE..tblToRecEntry cd ON cd.到货单编号 = cc.到货单编号 AND cd.到货批号 = cc.批号 LEFT JOIN erptemp..tblBB_CSRPO dd " & _
'"  ON dd.PO_NUM = bb.PO_NUM AND dd.FAB_DEVICE = bb.MPN_DESC LEFT JOIN  erpdata..tblSalerec j ON j.单据编号 = c.单据编号 AND j.单据项次 = b.单据项次 AND j.小箱号 = b.箱号 LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName  WHERE a.DN IN ('" & strDNList & "')  GROUP BY b.单据编号, c.单据日期, ISNULL(dn_address_new, d.shiptoname), " & _
'"  ISNULL(x.ship_to_street1_new, d.shiptostreet1), ISNULL(x.ship_to_street2_new, d.shiptostreet2),ISNULL(x.ship_to_street3_new, d.shiptostreet3), ISNULL(x.city_new, d.city),  ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode),ISNULL(x.country_new, d.countrykey),ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone), " & _
'" d.SalesDocument, d.PurchasingDocNo,f.箱号, b.料号,d.MarketingPN,d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber,   a.DN , J.销售单编号, bb.PO_NUM, cd.含税单价, cc.良品数, dd.WAFER_PRICE, qq.PASSBINCOUNT, dd.die_price, y.代码简称  )aaa where aaa.shiptoname = '" & ShiptoName & "' order by aaa.shiptoname,aaa.DN,aaa.箱号"
'
'
   strSql = "select * from(SELECT 0 AS 选择,y.代码简称 + b.单据编号 AS 单据编号,a.DN, CONVERT(VARCHAR(100), c.单据日期,23) AS 发货日期, ISNULL(dn_address_new, d.shiptoname) AS shiptoname ,ISNULL(x.ship_to_street1_new " & _
         " , d.shiptostreet1) AS shiptostreet1,ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2 , ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, d.city) AS city " & _
         " ,  ISNULL(x.dn_st_new, d.State) AS state, ISNULL(x.postal_code_new, d.postalcode) AS postalcode,ISNULL(x.country_new, d.countrykey) AS countrykey, ISNULL(x.contact_new, d.contactname) AS contactname " & _
         " , ISNULL(x.phone_new, d.phone) AS phone,d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo, f.箱号,  b.料号,d.MarketingPN,SUM(b.数量) as 数量, d.BatchNumber, d.DATE_CODE " & _
         " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no ,ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  + ( dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0) AS 单价 ,ROUND( SUM(b.数量) * ISNULL(ISNULL( cd.含税单价 / cc.良品数,0) " & _
         "  +  (dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0),2) AS AMount ,d.customerPartNumber,ROUND( SUM(b.数量) * ISNULL(ISNULL( cd.含税单价 / cc.良品数,0)  +  (dd.WAFER_PRICE/cc.良品数 + dd.DIE_PRICE),0),2)   -  CONVERT(DECIMAL(18,2) " & _
         " ,SUM(b.数量) * ISNULL( cd.含税单价 / cc.良品数,0)) AS 加工费金额 , CONVERT(DECIMAL(18,2),SUM(b.数量) * ISNULL( cd.含税单价 / cc.良品数,0)) AS 客供料金额 " & _
         "  FROM erpdata .. tblStockNumTree a INNER JOIN erpdata .. tblStocksqfhsub b ON b.箱号 = a.箱号 INNER JOIN erpdata .. tblStockSQfh c  ON c.单据编号 = b.单据编号 AND c.序号 = b.单据项次 " & _
         " INNER JOIN erpdata..tblstock y ON y.库房代码 = c.仓库编号   INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE, 1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox " & _
         " , SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job  FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.箱号 = aa.qbox " & _
         " INNER JOIN (SELECT dn.Delivery,   dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State,dn.postalcode, dn.countrykey, dn.contactname,dn.phone, dn.SalesDocument " & _
         " ,  dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE  FROM ERPBASE .. tblCustomerShippingUp dn " & _
         " WHERE dn.Delivery IN ('" & strDNList & "')  GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State, dn.postalcode,  dn.countrykey,dn.contactname,dn.phone " & _
         " , dn.SalesDocument, dn.PurchasingDocNo, dn.BatchNumber, dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE) d    ON d.Delivery = a.DN  AND d.BatchNumber = aa.job " & _
         " INNER JOIN erpdata .. tblstocknumtree f  ON f.序号 = a.上级序号  INNER JOIN ERPBASE..tblmappingData qq ON qq.SUBSTRATEID = b.流程卡编号 AND qq.LOTID = b.工单号 " & _
         " LEFT JOIN ERPBASE..tblCustomerOI bb  ON CONVERT(VARCHAR(100), bb.ID) = qq.FILENAME AND bb.SOURCE_BATCH_ID = qq.LOTID  LEFT JOIN ERPBASE..tblToInRec_Wafer cc " & _
         " ON cc.晶圆ID = REPLACE(qq.SUBSTRATEID,'+','')   AND cc.批号 = qq.LOTID  LEFT JOIN ERPBASE..TblToInSub cd ON cd.入库单编号 = cc.入库单编号 AND cd.到货批号 = cc.批号 " & _
         " LEFT JOIN erptemp..tblBB_CSRPO dd   ON dd.PO_NUM = bb.PO_NUM AND dd.FAB_DEVICE = bb.MPN_DESC LEFT JOIN  erpdata..tblSalerec j ON j.单据编号 = c.单据编号 AND j.单据项次 = b.单据项次 " & _
         " AND j.小箱号 = b.箱号 LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName  WHERE a.DN IN ('" & strDNList & "')  GROUP BY b.单据编号, c.单据日期, ISNULL(dn_address_new, d.shiptoname) " & _
         "  ,   ISNULL(x.ship_to_street1_new, d.shiptostreet1), ISNULL(x.ship_to_street2_new, d.shiptostreet2),ISNULL(x.ship_to_street3_new, d.shiptostreet3), ISNULL(x.city_new, d.city) " & _
         "  ,  ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode),ISNULL(x.country_new, d.countrykey),ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) " & _
         " ,  d.SalesDocument, d.PurchasingDocNo,f.箱号, b.料号,d.MarketingPN,d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber,   a.DN " & _
         " , J.销售单编号, bb.PO_NUM, cd.含税单价, cc.良品数, dd.WAFER_PRICE, qq.PASSBINCOUNT, dd.die_price, y.代码简称  )aaa where aaa.shiptoname = '" & ShiptoName & "' order by aaa.shiptoname,aaa.DN,aaa.箱号"
   
   rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     

   strExtsion = ".xlsx"    '获取后缀名
   strNewFullPath = "C:\合并_voice" & "\" & strExName & strExtsion    '获取新文件要保存的路径
'    Rs.MoveFirst    '数据集移动到第一个
'    '---------------------------------------------------------------------
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "初始化Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '是否显示
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
      
        DblNum = 0
        DblAmt = 0
        DblJZ = 0
        DblMZ = 0

        wkst.Cells(15, 10) = Trim$("" & rs.Fields(3).Value)
        wkst.Cells(18, 10) = Trim$("" & rs.Fields(11).Value) 'To
        
        wkst.Cells(13, 2) = Trim$("" & rs.Fields(4).Value)
        wkst.Cells(14, 2) = Trim$("" & rs.Fields(5).Value)
        wkst.Cells(15, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
        wkst.Cells(16, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
        
        wkst.Cells(18, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
      ' wkst.Cells(19, 2) = ""
        If UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION (CAMARILLO)" Or UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION" Then
        
           wkst.Cells(19, 2) = "TAX ID: 95-2119684"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INTERCONNECT" Then
        
           wkst.Cells(19, 2) = "TAX ID: 82-5035949"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INCORPORATED(FEDERAL）" Then
        
            wkst.Cells(19, 2) = "TAX ID: 82-5035949"
        
        Else
           
            wkst.Cells(19, 2) = ""
        
        End If
        'wkst.Cells(23, 3) = Trim$("" & Rs.fields(14).Value)
        wkst.Cells(23, 5) = Trim$("'" & rs.Fields(15).Value)
        
        lngRows = 27
        
        IntInertRow = rs.RecordCount * 2
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
           wkst.Range(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '边框无
        Next i
        IntMaxDetailRow = rs.RecordCount
        
        IntBMegerRow = 26
        IntEMegerRow = 29
        intBegin = 1
        Dim QBX As String
        
        For i = 0 To rs.RecordCount - 1
'            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '箱号

             If dnnum1 <> Trim$("" & rs.Fields(2).Value) And InStr(dnnum, Trim$("" & rs.Fields(2).Value)) = 0 Then
                dnnum = Trim$("" & rs.Fields(2).Value) + "/" + dnnum
                dnnum1 = Trim$("" & rs.Fields(2).Value)
             End If
             

            
            strPBigBox = Trim$("" & rs.Fields(16).Value) '箱号
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & rs.Fields(16).Value) '箱号
                '箱数
                intBoxNum = intBoxNum + 1
                wkst.Cells(lngRows, 1) = "K" & Trim(intBoxNum - 1) '箱号进行转换为客户所要资料
                QBX = "K" & Trim(intBoxNum - 1)
                
                IntBMegerRow = IntBMegerRow + intBegin
                intBegin = 1
            Else
'                '合并
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).Merge
'                '设定水平和竖直居中
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).HorizontalAlignment = xlCenter
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).VerticalAlignment = xlCenter
                '--------------------------
                intBegin = intBegin + 1
            End If
              If SD <> Trim$("" & rs.Fields(14).Value) Then
             SD = Trim$("" & rs.Fields(14).Value)
             SD1 = SD1 & SD & " "
             End If
            wkst.Cells(23, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & rs.Fields(19).Value)) / 1000 '数量都为除以1000的值
            DblNum = DblNum + Val(Trim$("" & rs.Fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(21).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(22).Value)
            wkst.Cells(lngRows, 13) = "US$"
            wkst.Cells(lngRows, 14) = Val(Trim$("" & rs.Fields(23).Value)) * 1000 '单价为乘以1000的值
            wkst.Cells(lngRows, 15) = "US$"
            wkst.Cells(lngRows, 16) = Trim$("" & rs.Fields(24).Value)
            DblAmt = DblAmt + Val(Trim$("" & rs.Fields(24).Value))
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(25).Value)
            
            jine1 = jine1 + Val(Trim$("" & rs.Fields(26).Value))
            
            jine2 = jine2 + Val(Trim$("" & rs.Fields(27).Value))
            
            
            
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            rs.MoveNext
        Next
        
        wkst.Cells(13, 10) = Mid(dnnum, 1, Len(dnnum) - 1)
        
        '计算汇总
        wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '数量
        wkst.Cells(lngRows + 1, 9) = "KPCS" '单位
        wkst.Cells(lngRows + 1, 16) = DblAmt
        wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)
        
        wkst.Cells(lngRows + 8, 1) = "Process Amount：US$ " + Str(jine1)
        wkst.Cells(lngRows + 9, 1) = "Wafer Amount：US$ " + Str(jine2)

        
    Else
'        ClsP.UnLoad_Form
        MsgBox "无需导出数据！", vbInformation, "提示！"
        Exit Function
    End If

    With wkst.PageSetup

    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("此文件已经存在，是否要覆盖原文件?", vbYesNo Or vbQuestion Or vbDefaultButton2, "提示") = vbNo Then
            Exit Function
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.number <> 0 Then
                MsgBox "覆盖文件失败，请手动删除文件再导出。", vbInformation, "提示"
                Exit Function
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "导出成功！"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Function
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.DESCRIPTION, vbInformation, "提示！"
    Exit Function
End Function


'根据Rs数据集语句导出Excel
Public Sub RsExporToExcel(rs As ADODB.Recordset, RptName As String, ExcelFileName As String)
Dim Irowcount       As Long
Dim Icolcount       As Integer
Dim strFileName     As String

    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlQuery As Excel.QueryTable
    Screen.MousePointer = 11
    With rs
        If .RecordCount < 1 Then
            Screen.MousePointer = 0
            MsgBox ("没有可导出的资料")
            Exit Sub
        End If
        Irowcount = .RecordCount
        Icolcount = .Fields.count
    End With
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlApp.Visible = False

    Set xlQuery = xlSheet.QueryTables.Add(rs, xlSheet.Range("a1"))
    
'    With xlQuery
'        .FieldNames = True
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = True
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = True
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .PreserveColumnInfo = True
'    End With
    xlSheet.name = RptName
    xlQuery.FieldNames = True '
    xlQuery.Refresh
    
    With xlSheet
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.name = "宋体"
        '
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = True
        '
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
        '
'        .Range(.Cells(2, 1), .Cells(Irowcount + 1, Icolcount)).Font.Size = 9
    End With
    '另存文件
    strFileName = DirInvRpt + "\" + ExcelFileName
    'xlBook.SaveAs strFileName, xlNormal, "", "", False, False
    xlBook.SaveAs strFileName
    xlBook.Saved = True
    
    Screen.MousePointer = 0
'    xlApp.Visible = True
    Set xlSheet = Nothing
    xlBook.Close
    Set xlBook = Nothing
    xlApp.Quit
    Set xlApp = Nothing
  

End Sub








