VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.ocx"
Begin VB.Form Frm_OPkgPrintGen 
   Caption         =   "通用外箱标签补打"
   ClientHeight    =   9330
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   10635
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
   ScaleHeight     =   9330
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "功能"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   16695
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "退出"
         Height          =   480
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFC0FF&
         Caption         =   "打印"
         Height          =   480
         Left            =   3320
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdBind 
         BackColor       =   &H00FFC0C0&
         Caption         =   "下个大箱"
         Height          =   480
         Left            =   1960
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdGetSubBoxID 
         BackColor       =   &H00C0FFFF&
         Caption         =   "录入箱号"
         Height          =   480
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "箱号List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8535
      Left            =   3720
      TabIndex        =   1
      Top             =   1080
      Width           =   13095
      Begin FPSpreadADO.fpSpread fpS 
         Height          =   8175
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   12615
         _Version        =   524288
         _ExtentX        =   22251
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
         MaxCols         =   4
         MaxRows         =   0
         SpreadDesigner  =   "Frm_OPkgPrintGen.frx":0000
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "箱号录入"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
      Begin VB.TextBox txtScan 
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
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   825
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cbType 
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
         Height          =   360
         ItemData        =   "Frm_OPkgPrintGen.frx":03F8
         Left            =   1200
         List            =   "Frm_OPkgPrintGen.frx":040E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下个大箱扫描:  0000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   1965
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   4440
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱箱号(Q)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户类型"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   960
      End
   End
End
Attribute VB_Name = "Frm_OPkgPrintGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBind_Click()
Call NextWX

End Sub

Private Sub NextWX()
Dim i As Integer

With Fps(0)
    i = .MaxRows
    .Col = 1
    .Row = i
    If .text <> "" Then
        .MaxRows = .MaxRows + 1
        .SetFocus

    End If

End With

txtScan.SetFocus

End Sub

Private Sub cmdExit_Click()

GetGD212JZ (5000)
Unload Me

End Sub

Private Sub cmdGetSubBoxID_Click()
Call OpenShip

End Sub

Private Sub cmdPrint_Click()
Call CloseShip

End Sub

Private Sub Form_Load()
Call InitCtrl

End Sub

Private Sub InitCtrl()

With Fps(0)
    .DAutoCellTypes = False
    .Col = -1
    .Row = -1
    .Lock = True
    .TypeVAlign = TypeVAlignBottom
    .TypeHAlign = TypeHAlignLeft
    .Col = 4
    .Lock = False
    .BackColor = vbGreen
    .SetText 0, 0, "序号"
    .SetText 1, 0, "外箱Q箱号"
    .SetText 2, 0, "外箱重量"
    .SetText 3, 0, "外箱数量"
    .SetText 4, 0, "第几箱"
    .Col = 2
    .Lock = False
    .ColWidth(0) = 8
    .ColWidth(1) = 50
    .ColWidth(2) = 10
    .ColWidth(3) = 10
    .ColWidth(4) = 11

End With

End Sub

Private Sub OpenShip()
If cbType.text = "" Then
    MsgBox "请选择客户类型", vbInformation, "提示"
    Exit Sub

End If

PlaySound ("开始录入外箱箱号")
Fps(0).MaxRows = 0
txtScan.Visible = True
txtScan.SetFocus

End Sub

Private Sub CloseShip()
Dim i As Integer

With Fps(0)
    .Row = .MaxRows
    .Col = 1
    If .text = "" Then
        .MaxRows = .MaxRows - 1

    End If

End With

If Fps(0).MaxRows = 0 Then
    MsgBox "当前未有箱号录入", vbInformation, "提示"
    Exit Sub

End If

PrintLblData
txtScan.Visible = False
Fps(0).MaxRows = 0

End Sub

Private Sub PrintLblData()
Dim i                 As Integer
Dim strPkgID          As String
Dim strSql            As String
Dim strMaxLblID       As String
Dim strBartenderName  As String
Dim strBartenderName2 As String
Dim strPrinterName    As String
Dim strPrinterName2   As String
Dim strTagName1       As String, strTagValue1 As String
Dim strTagName2       As String, strTagValue2 As String
Dim strTagName3       As String, strTagValue3 As String
Dim strTagName4       As String, strTagValue4 As String
Dim lQty              As Long
Dim lAllQty           As Long

Select Case cbType.ListIndex

    Case 0  ' 艾为:HK037/AC70
        strBartenderName = "HK037OUT2.btw"
        strPrinterName = "ALL_OUT_2B1F_2"
        strTagName1 = "CARTON_ID1"
        strTagName2 = "CARTON_ID2"
        strTagName3 = "CARTON_ID3"

    Case 1  ' DA69
        strBartenderName = "DA69OUT1.btw"
        strPrinterName = "ALL_OUT_2B1F_2"
        strTagName1 = "CARTON_ID4"
        strTagName2 = "CARTON_ID5"

    Case 2 ' EQ
        strBartenderName = "ISOUT1.btw"
        strBartenderName2 = "ISOUT2.btw"
        strPrinterName = "ALL_OUT_2B1F_2"
        strPrinterName2 = "ALL_OUT_2B1F_1"
        strTagName1 = "CARTON_ID1"
        strTagName2 = "CARTON_ID2"

    Case 3 ' US023
        strBartenderName = "US023IN1.btw"
        strPrinterName = "W_IN_2B5F_2"
        strTagName1 = "CARTON_ID1"
        strTagName2 = "CARTON_ID2"

    Case 4 ' AH017
        strBartenderName = "AH017OUT1.btw"
        strPrinterName = "ALL_OUT_2B1F_2"
        strTagName1 = "CARTON_ID6"
    
    Case 5 ' GD212
        strBartenderName = "GD212OUT1-NEW.btw"
        strPrinterName = "ALL_OUT_2B1F_2"
        strTagName1 = "CARTON_ID1"
        strTagName2 = "CARTON_ID2"
        strTagName3 = "JZ"
        strTagName4 = "MZ"
    
End Select

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .text = "" Then
            GoTo OUT

        End If

        strPkgID = Split(Trim$(.text), ";")(0)
        .Col = 3
        lAllQty = .text
        strSql = " select MAX(a.id) from erpdata..tblME_PrintInfo a " & " inner join erpdata..tblErpInStockDetailInfo b on a.EVENT_ID = b.KEYID and b.KEY_NAME = 'CONTAINER_NAME' " & " where b.KEY_VALUE = '" & strPkgID & "' and a.BartenderName = '" & strBartenderName & "' and a.PrinterNameID = '" & strPrinterName & "' and charindex('" & strTagName1 & "',a.Content) = 0  "
        strMaxLblID = Get_SqlStr(strSql)
        If strMaxLblID = "" Then
            MsgBox "箱号: " & strPkgID & "查询不到,不能补打", vbInformation, "提示"
            Exit Sub

        End If

        Select Case cbType.ListIndex

            Case 0
                strTagValue1 = i
                strTagValue2 = .MaxRows
                strTagValue3 = strTagValue1 & " OF " & strTagValue2
                If lAllQty <> 0 Then
                    lQty = Get_SqlserverNo(" select SUM(数量) from erpdata..tblPackMainInf where 箱号 = '" & strPkgID & "'")
                    strSql = " INSERT INTO erpdata..tblME_PrintInfo(PrinterNameID,BartenderName,Content,Content2,Content3,flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & " SELECT a.PrinterNameID,a.BartenderName,replace(a.content,'" & lQty & "','" & lAllQty & "') + ';' + '""" & strTagName1 & """' + ',' + '""" & strTagValue1 & """' + ';' + '""" & strTagName2 & """' + ',' + '""" & strTagValue2 & """' + ';' + '""" & strTagName3 & """' + ',' + '""" & strTagValue3 & """' AS Content,a.Content2,a.Content3,'0' AS flag ,GETDATE() AS create_date,a.EVENT_SOURCE AS EVENT_SOURCE,a.EVENT_ID,a.LABEL_ID,'1' " & " FROM erpdata..tblME_PrintInfo a WHERE a.ID = '" & strMaxLblID & "' "
                Else
                    strSql = " INSERT INTO erpdata..tblME_PrintInfo(PrinterNameID,BartenderName,Content,Content2,Content3,flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & " SELECT a.PrinterNameID,a.BartenderName,a.Content + ';' + '""" & strTagName1 & """' + ',' + '""" & strTagValue1 & """' + ';' + '""" & strTagName2 & """' + ',' + '""" & strTagValue2 & """' + ';' + '""" & strTagName3 & """' + ',' + '""" & strTagValue3 & """' AS Content,a.Content2,a.Content3,'0' AS flag ,GETDATE() AS create_date,a.EVENT_SOURCE AS EVENT_SOURCE,a.EVENT_ID,a.LABEL_ID,'1' " & " FROM erpdata..tblME_PrintInfo a WHERE a.ID = '" & strMaxLblID & "' "

                End If

            Case 1
                .Col = 4
                strTagValue1 = Right$("000" & .text, 3)
                .Row = .MaxRows
                .Col = 4
                strTagValue2 = Right$("000" & .text, 3)
                strSql = " INSERT INTO erpdata..tblME_PrintInfo(PrinterNameID,BartenderName,Content,Content2,Content3,flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & " SELECT a.PrinterNameID,a.BartenderName,a.Content + ';' + '""" & strTagName1 & """' + ',' + '""" & strTagValue1 & """' + ';' + '""" & strTagName2 & """' + ',' + '""" & strTagValue2 & """' AS Content,a.Content2,a.Content3,'0' AS flag ,GETDATE() AS create_date,a.EVENT_SOURCE AS EVENT_SOURCE,a.EVENT_ID,a.LABEL_ID,a.PRINT_QTY " & " FROM erpdata..tblME_PrintInfo a WHERE a.ID = '" & strMaxLblID & "' "

            Case 2
                .Col = 4
                strTagValue1 = .text
                .Row = .MaxRows
                .Col = 4
                strTagValue2 = .text
                strSql = " INSERT INTO erpdata..tblME_PrintInfo(PrinterNameID,BartenderName,Content,Content2,Content3,flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & " SELECT a.PrinterNameID,a.BartenderName,a.Content + ';' + '""" & strTagName1 & """' + ',' + '""" & strTagValue1 & """' + ';' + '""" & strTagName2 & """' + ',' + '""" & strTagValue2 & """' AS Content,a.Content2,a.Content3,'0' AS flag ,GETDATE() AS create_date,a.EVENT_SOURCE AS EVENT_SOURCE,a.EVENT_ID,a.LABEL_ID,a.PRINT_QTY " & " FROM erpdata..tblME_PrintInfo a WHERE a.ID = '" & strMaxLblID & "' "

            Case 3
                .Col = 4
                strTagValue1 = .text
                .Row = .MaxRows
                .Col = 4
                strTagValue2 = .text
                strSql = " INSERT INTO erpdata..tblME_PrintInfo(PrinterNameID,BartenderName,Content,Content2,Content3,flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & " SELECT a.PrinterNameID,a.BartenderName,a.Content + ';' + '""" & strTagName1 & """' + ',' + '""" & strTagValue1 & """' + ';' + '""" & strTagName2 & """' + ',' + '""" & strTagValue2 & """' AS Content,a.Content2,a.Content3,'0' AS flag ,GETDATE() AS create_date,a.EVENT_SOURCE AS EVENT_SOURCE,a.EVENT_ID,a.LABEL_ID,a.PRINT_QTY " & " FROM erpdata..tblME_PrintInfo a WHERE a.ID = '" & strMaxLblID & "' "

            Case 4
                strTagValue1 = Right$("00" & i, 2)
                strSql = " INSERT INTO erpdata..tblME_PrintInfo(PrinterNameID,BartenderName,Content,Content2,Content3,flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & " SELECT a.PrinterNameID,a.BartenderName,a.Content + ';' + '""" & strTagName1 & """' + ',' + '""" & strTagValue1 & """' AS Content,a.Content2,a.Content3,'0' AS flag ,GETDATE() AS create_date,a.EVENT_SOURCE AS EVENT_SOURCE,a.EVENT_ID,a.LABEL_ID,a.PRINT_QTY " & " FROM erpdata..tblME_PrintInfo a WHERE a.ID = '" & strMaxLblID & "' "
            
            Case 5
                .Col = 4
                strTagValue1 = .text
                .Row = .MaxRows
                .Col = 4
                strTagValue2 = .text
                
                strTagValue3 = GetGD212JZ(lAllQty)
                strTagValue4 = GetGD212MZ(strPkgID)
                strSql = " INSERT INTO erpdata..tblME_PrintInfo(PrinterNameID,BartenderName,Content,Content2,Content3,flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & " SELECT a.PrinterNameID,a.BartenderName,a.Content + ';' + '""" & strTagName1 & """' + ',' + '""" & strTagValue1 & """' + ';' + '""" & strTagName2 & """' + ',' + '""" & strTagValue2 & """'+ ';' + '""" & strTagName3 & """' + ',' + '""" & strTagValue3 & """'+ ';' + '""" & strTagName4 & """' + ',' + '""" & strTagValue4 & """' AS Content,a.Content2,a.Content3,'0' AS flag ,GETDATE() AS create_date,a.EVENT_SOURCE AS EVENT_SOURCE,a.EVENT_ID,a.LABEL_ID,a.PRINT_QTY " & " FROM erpdata..tblME_PrintInfo a WHERE a.ID = '" & strMaxLblID & "' "
            
        End Select

        If AddSql2(strSql) = 0 Then
            MsgBox "箱号: " & strPkgID & vbCrLf & "没有成功打印", vbInformation, "提示"
            Exit Sub

        End If

        If cbType.ListIndex = 2 Then
            strSql = " select MAX(a.id) from erpdata..tblME_PrintInfo a " & " inner join erpdata..tblErpInStockDetailInfo b on a.EVENT_ID = b.KEYID and b.KEY_NAME = 'CONTAINER_NAME' " & " where b.KEY_VALUE = '" & strPkgID & "' and a.BartenderName = '" & strBartenderName2 & "' and a.PrinterNameID = '" & strPrinterName2 & "' and charindex('" & strTagName1 & "',a.Content) = 0  "
            strMaxLblID = Get_SqlStr(strSql)
            If strMaxLblID = "" Then
                MsgBox "箱号: " & strPkgID & "查询不到,不能补打", vbInformation, "提示"
                Exit Sub

            End If

            .Row = i
            .Col = 4
            strTagValue1 = .text
            .Row = .MaxRows
            .Col = 4
            If .text = "" Then
                MsgBox "请填写实际第几箱", vbExclamation, "提示"
                Exit Sub

            End If

            strTagValue2 = Trim$(.text)
            strSql = " INSERT INTO erpdata..tblME_PrintInfo(PrinterNameID,BartenderName,Content,Content2,Content3,flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & " SELECT a.PrinterNameID,a.BartenderName,a.Content + ';' + '""" & strTagName1 & """' + ',' + '""" & strTagValue1 & """' + ';' + '""" & strTagName2 & """' + ',' + '""" & strTagValue2 & """' AS Content,a.Content2,a.Content3,'0' AS flag ,GETDATE() AS create_date,a.EVENT_SOURCE AS EVENT_SOURCE,a.EVENT_ID,a.LABEL_ID,a.PRINT_QTY " & " FROM erpdata..tblME_PrintInfo a WHERE a.ID = '" & strMaxLblID & "' "
            If AddSql2(strSql) = 0 Then
                MsgBox "箱号: " & strPkgID & vbCrLf & "没有成功打印", vbInformation, "提示"
                Exit Sub

            End If

        End If

OUT:
    Next i

End With

MsgBox "箱号补打完成", vbInformation, "提示"

End Sub

Private Function GetGD212JZ(lGrossDies As Long) As String
GetGD212JZ = Format(lGrossDies * 0.0014 / 1000, "0.00")
If GetGD212JZ = "0.00" Then
    GetGD212JZ = "0.01"
End If

End Function

Private Function GetGD212MZ(strPkgID As String) As String
Dim iReels As Integer
Dim strReels As String
Dim strSql As String

iReels = Get_SqlserverNo("select count(1) from erpdata..tblErpInStockDetailInfo t1 " & _
" inner join erpdata..tblErpInStockDetailInfo t2 on t1.box_id = t2.box_id and t1.key_name = 'CONTAINER_NAME' and t1.key_type = 'B' " & _
" where t2.key_type = 'C' and t2.key_name = 'CONTAINER_NAME' and t2.key_value = '" & strPkgID & "'")

If iReels <= 8 Then
    GetGD212MZ = Format(iReels * 0.168 + (8 - iReels) * 0.052 + 0.371, "0.00")
Else
    GetGD212MZ = Format(iReels * 0.168 + (16 - iReels) * 0.052 + 0.515, "0.00")

End If

End Function


Private Sub txtScan_KeyPress(KeyAscii As Integer)
Dim strScan As String

If KeyAscii <> vbKeyReturn Or Len(Trim(txtScan.text)) = 0 Then Exit Sub
strScan = UCase$(Trim$(txtScan.text))
If strScan = "0000" Then
    NextWX
Else
    If CheckScanData(strScan) = True Then
        Call ShowLblData(strScan)

    End If

End If

txtScan.text = ""

End Sub

Private Function CheckScanData(strData As String) As Boolean
Dim i      As Integer
Dim strSql As String

CheckScanData = False
strSql = " select * from erpdata..tblErpInStockDetailInfo where KEY_VALUE = '" & strData & "' "
If Get_SqlserverCnt(strSql) = 0 Then
    MsgBox "该箱号不存在, 请扫描正确的Q大箱号", vbCritical, "警告"
    Exit Function

End If

With Fps(0)
    If .MaxRows < 1 Then
        CheckScanData = True
        Exit Function

    End If

    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If InStr(.text, strData) > 0 Then
            MsgBox "该箱号:" & strData & "已经扫描过, 请确认是否扫描出错", vbInformation, "提示"
            Exit Function

        End If

    Next

End With

CheckScanData = True

End Function

Private Function ShowLblData(strData As String)
Dim i    As Integer
Dim lQty As Long

lQty = Get_SqlserverNo(" select SUM(数量) from erpdata..tblPackMainInf where 箱号 = '" & strData & "'")

With Fps(0)
    If cbType.ListIndex <> 0 Then
        .MaxRows = .MaxRows + 1

    End If

    .MaxRows = IIf(.MaxRows = 0, 1, .MaxRows)
    i = .MaxRows
    .Col = 1
    .Row = i
    .SetText 1, i, strData & ";" & .text
    .Col = 3
    .Row = i
    If .text = "" Then
        .SetText 3, i, lQty
    Else
        .SetText 3, i, CLng(.text) + lQty

    End If

    .Col = 4
    .text = .MaxRows

End With

PlaySound ("箱号已扫描")

End Function

Private Sub PlaySound(sFileName As String)
Dim sPath   As String
Dim sSuffix As String

sPath = "\\10.160.1.84\public\media_source\"
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix

End Sub
