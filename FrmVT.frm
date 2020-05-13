VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmVT 
   Caption         =   "VT回来资料上传"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10740
   ScaleWidth      =   15510
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   10335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   18230
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "回货资料上传"
      TabPicture(0)   =   "FrmVT.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fpS"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CboCustomer"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "回货资料删除"
      TabPicture(1)   =   "FrmVT.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd_search"
      Tab(1).Control(1)=   "Cmd_del"
      Tab(1).Control(2)=   "Txt_lotid"
      Tab(1).Control(3)=   "fpS_del"
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(5)=   "Label3"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmd_search 
         Caption         =   "查询"
         Height          =   375
         Left            =   -72240
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Cmd_del 
         Caption         =   "删除"
         Height          =   375
         Left            =   -70320
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Txt_lotid 
         Height          =   375
         Left            =   -74160
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin FPSpreadADO.fpSpread fpS_del 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   14
         Top             =   2280
         Width           =   15015
         _Version        =   524288
         _ExtentX        =   26485
         _ExtentY        =   13785
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
         MaxCols         =   0
         MaxRows         =   0
         SpreadDesigner  =   "FrmVT.frx":0038
      End
      Begin VB.Frame Frame3 
         Caption         =   "选择待上传的文件"
         Height          =   2775
         Left            =   360
         TabIndex        =   3
         Top             =   1020
         Width           =   14655
         Begin VB.CommandButton Cmd_GCNewformat 
            Caption         =   "GC回货新格式上传"
            Height          =   495
            Left            =   7680
            TabIndex        =   20
            Top             =   1560
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton Cmd_exportWO 
            Caption         =   "转换WO"
            Height          =   495
            Left            =   6000
            TabIndex        =   10
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox Txt_sqdh 
            Height          =   375
            Left            =   1440
            TabIndex        =   9
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "回货申请"
            Height          =   495
            Left            =   2520
            TabIndex        =   8
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   495
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   840
            Width           =   4935
         End
         Begin VB.CommandButton Command6 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   6
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton Command7 
            Caption         =   "上传DB"
            Height          =   480
            Left            =   480
            TabIndex        =   5
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command8 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   4320
            TabIndex        =   4
            Top             =   1560
            Width           =   1095
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   3000
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label2 
            Caption         =   "申请单号"
            Height          =   255
            Left            =   600
            TabIndex        =   12
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择待上传的xlsx："
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   2
            Left            =   480
            TabIndex        =   11
            Top             =   480
            Width           =   1620
         End
      End
      Begin MSDataListLib.DataCombo CboCustomer 
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   300
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin FPSpreadADO.fpSpread fpS 
         Height          =   5655
         Left            =   360
         TabIndex        =   13
         Top             =   4260
         Width           =   14655
         _Version        =   524288
         _ExtentX        =   25850
         _ExtentY        =   9975
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
         MaxCols         =   5
         MaxRows         =   0
         SpreadDesigner  =   "FrmVT.frx":04AE
         AppearanceStyle =   0
      End
      Begin VB.Label Label4 
         Caption         =   "仅能删除还未生成回货申请的数据"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -72240
         TabIndex        =   19
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "LOTID"
         Height          =   495
         Left            =   -74880
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户："
         Height          =   180
         Left            =   360
         TabIndex        =   1
         Top             =   300
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmVT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vtDataTemp As VTData
Dim shipid As Long
Dim mainItemRS As New ADODB.Recordset
Dim OldBoxList As String
Dim NewBoxList As String
Dim VTformat As String


Private Type GCCUSTDATA  '客户提供的回货信息

BlankRowTemp As Boolean
dateTemp As String
weightTemp As String
C_NOTemp As String
BoxIdTemp As String
lotIdTemp As String
waferIdTemp As String
CustDeviceTemp As String
CustAttributeTemp As String
GccodeTemp As String
GcLevelTemp As String
pieceQtyTemp As String
qtyTemp As String
remarkTemp As String
PackageSizeTemp As String


End Type



Private Function GCformatTranslate()

On Error GoTo ErrHandle

    Dim dT As GCCUSTDATA
    Dim i As Integer
    Dim j As Integer
    Dim VBExcel_Source   As Excel.Application
    Dim xlBook_Source    As Excel.Workbook
    Dim xlSheet_Source   As Excel.Worksheet
    Dim lColsCnt  As Long
    Dim lRowsCnt  As Long
     
    GCformatTranslate = False
    
    If Trim(Text3.text) = "" Then
        MsgBox "请选择源文件所在路径", vbInformation, "提示"
        Exit Function
    End If
    Set VBExcel_Source = CreateObject("excel.application")
    VBExcel_Source.Visible = False
    Set xlBook_Source = VBExcel_Source.Workbooks.Open(Text3.text)
    Set xlSheet_Source = xlBook_Source.Worksheets(1)

   
    lColsCnt = xlSheet_Source.Range("A1").CurrentRegion.Columns.count
    lRowsCnt = xlSheet_Source.Range("A1").CurrentRegion.Rows.count

    'If InStr(Trim(xlSheet_Source.Range("A1").Value), "华天") = 0 Then
    '    MsgBox "请选择客户所提供的格式上传！", vbInformation, "提示"
   '     GoTo EXITPRO
    '    Exit Function
  '  End If


    
    j = 0
    

    cmdStr = "delete from erptemp..GcExcelTranslate"
    AddSql2 (cmdStr)

If VTformat = "new" Then
    If lColsCnt <> 14 Then
        MsgBox "Excel中的列数:" & lColsCnt & "和设定的模版列数:13不一致" & vbCrLf & "请确认Excel是否正确！", vbInformation, "提示"
        GoTo EXITPRO
        Exit Function

    End If
    
    startrow = 2
    If InStr(Trim(xlSheet_Source.Range("B1").Value), "产品型号") = 0 Then
        MsgBox "新模板A1单元格内应为产品型号四个字,请确认格式", vbInformation, "提示"
        Exit Function
    End If
ElseIf VTformat = "old" Then
    If lColsCnt <> 13 Then
        MsgBox "Excel中的列数:" & lColsCnt & "和设定的模版列数:13不一致" & vbCrLf & "请确认Excel是否正确！", vbInformation, "提示"
        GoTo EXITPRO
        Exit Function

    End If
    startrow = 4
    
    If InStr(Trim(xlSheet_Source.Range("A2").Value), "日期") = 0 Then
        MsgBox "旧模板A2单元格内应为日期二字,请确认格式", vbInformation, "提示"
        Exit Function
    End If
End If
    For i = startrow To lRowsCnt
        If getData(dT, xlSheet_Source, i) Then
            If dT.BlankRowTemp = False Then
                j = j + 1
                Call WriteData(dT, j)
            End If
        Else
            GoTo EXITPRO
            Exit Function
        End If
    Next
    
    

    xlBook_Source.Close
    Set VBExcel_Source = Nothing
    Set xlBook_Source = Nothing
    Set xlSheet_Source = Nothing
    
    GCformatTranslate = True

EXITPRO:

On Error Resume Next

MousePointer = 0
If Not VBExcel_Source Is Nothing Then
    xlBook_Source.Close
    Set xlSheet_Source = Nothing
    Set xlBook_Source = Nothing
    Set VBExcel_Source = Nothing
End If

Exit Function
ErrHandle:
GoTo EXITPRO

End Function




Private Function getData(ByRef dT As GCCUSTDATA, xlSheet As Excel.Worksheet, i As Integer)
   
    getData = True
    dT.BlankRowTemp = False
    If Replace(Trim(xlSheet.Range("E" & i)), Chr(13) + Chr(10), "") = "" Then
        dT.BlankRowTemp = True
        Exit Function
    End If
    If VTformat = "old" Then
        dT.dateTemp = GetMergeCellsValue(xlSheet, "A" & i)
        dT.weightTemp = GetMergeCellsValue(xlSheet, "B" & i)
        dT.C_NOTemp = GetMergeCellsValue(xlSheet, "C" & i)
        dT.BoxIdTemp = GetMergeCellsValue(xlSheet, "D" & i)
        dT.waferIdTemp = GetMergeCellsValue(xlSheet, "E" & i)
        If InStr(dT.waferIdTemp, "-") > 0 Then
            dT.lotIdTemp = Split(dT.waferIdTemp, "-")(0)
        Else
            dT.lotIdTemp = dT.waferIdTemp
            MsgBox "D列waferid格式不正确", vbInformation, "提示"
            getData = False
            Exit Function
        End If
        dT.CustDeviceTemp = GetMergeCellsValue(xlSheet, "F" & i)
        dT.CustAttributeTemp = GetMergeCellsValue(xlSheet, "G" & i)
        dT.GccodeTemp = GetMergeCellsValue(xlSheet, "H" & i)
    
        dT.GcLevelTemp = GetMergeCellsValue(xlSheet, "I" & i)
        dT.pieceQtyTemp = GetMergeCellsValue(xlSheet, "J" & i)
        dT.qtyTemp = GetMergeCellsValue(xlSheet, "K" & i)
        dT.remarkTemp = GetMergeCellsValue(xlSheet, "L" & i)
        dT.PackageSizeTemp = GetMergeCellsValue(xlSheet, "M" & i)
    ElseIf VTformat = "new" Then
        If GetMergeCellsValue(xlSheet, "A" & i) = "" Then
            dT.dateTemp = Format(Now(), "yyyy/mm/dd")
        Else
            dT.dateTemp = GetMergeCellsValue(xlSheet, "A" & i)
        End If
        dT.CustDeviceTemp = GetMergeCellsValue(xlSheet, "B" & i)
        dT.GccodeTemp = GetMergeCellsValue(xlSheet, "C" & i)
        dT.BoxIdTemp = GetMergeCellsValue(xlSheet, "D" & i)
        dT.C_NOTemp = GetMergeCellsValue(xlSheet, "E" & i)
        dT.waferIdTemp = GetMergeCellsValue(xlSheet, "F" & i)
        If InStr(dT.waferIdTemp, "-") > 0 Then
            dT.lotIdTemp = Split(dT.waferIdTemp, "-")(0)
        Else
            dT.lotIdTemp = dT.waferIdTemp
            MsgBox "D列waferid格式不正确", vbInformation, "提示"
            getData = False
            Exit Function
        End If
        dT.GcLevelTemp = GetMergeCellsValue(xlSheet, "G" & i)
        dT.pieceQtyTemp = GetMergeCellsValue(xlSheet, "H" & i)
        dT.qtyTemp = GetMergeCellsValue(xlSheet, "J" & i)
        dT.remarkTemp = GetMergeCellsValue(xlSheet, "K" & i)
        dT.CustAttributeTemp = GetMergeCellsValue(xlSheet, "M" & i)
         
        
        dT.weightTemp = ""
        dT.PackageSizeTemp = ""
    
    End If


End Function



Private Function GetMergeCellsValue(xlSheet As Excel.Worksheet, CellAddress As String)
    '合并单元格，获取左上角单元格的value
    Dim left_top_cell As String

    If xlSheet.Range(CellAddress).MergeArea.MergeCells = True Then
        left_top_cell = Split(xlSheet.Range(CellAddress).MergeArea.Address, ":")(0)
        
        GetMergeCellsValue = Replace(Trim(xlSheet.Range(left_top_cell).Value), Chr(13) + Chr(10), "")
    Else
       GetMergeCellsValue = Replace(Trim(xlSheet.Range(CellAddress).Value), Chr(13) + Chr(10), "")
    End If

End Function


Private Sub WriteData(ByRef dT As GCCUSTDATA, j As Integer)

    
Dim cmdStr As String
Dim cmdStr2 As String
Dim LOTID As String
Dim WAFER As String
Dim strgcrev_2 As String
Dim strHtDevice As String
Dim strtype As String

'添加导入Sqlserver

If InStr(dT.waferIdTemp, "-") > 0 Then
    LOTID = Split(dT.waferIdTemp, "-")(0)
    WAFER = Split(dT.waferIdTemp, "-")(1)
Else
    LOTID = ""
    WAFER = ""
End If
strHtDevice = ""
strtype = ""
strgcrev_2 = ""
If Trim(dT.remarkTemp) = "" Then   'WLT
    strtype = "WLT"
ElseIf UCase(Replace(Trim(dT.remarkTemp), " ", "")) = "降MAIN" Then   '转Normal
    strtype = "转Normal"
Else
    strtype = ""
End If

strgcrev_2 = GetGcrevFromWO(LOTID, WAFER)
If Len(strgcrev_2) = 2 Then
    strHtDevice = GetHTDevice(dT.CustDeviceTemp, strtype, Right(strgcrev_2, 1))
End If
cmdStr = "insert into erptemp..GcExcelTranslate(日期,重量,C_NO,大箱_CST,WaferID,型号,属性,二级代码,等级,片数,数量,入库备注,包装尺寸,LotID,Wafer,id,remark1 ,remark2, remark3 ) values  " & _
" ('" & dT.dateTemp & "','" & dT.weightTemp & "','" & dT.C_NOTemp & "','" & dT.BoxIdTemp & "','" & dT.waferIdTemp & "','" & dT.CustDeviceTemp & "','" & dT.CustAttributeTemp & "','" & dT.GccodeTemp & "','" & dT.GcLevelTemp & "','" & dT.pieceQtyTemp & "','" & dT.qtyTemp & "','" & dT.remarkTemp & "','" & dT.PackageSizeTemp & "','" & LOTID & "','" & WAFER & "','" & j & "','" & strtype & "','" & strgcrev_2 & "','" & strHtDevice & "')"

                
AddSql2 (cmdStr)

Exit Sub

    
    

End Sub

Private Sub ExportToExcel()
    Dim xlsApp      As Excel.Application
    Dim xlsBook     As Excel.Workbook
    Dim xlsSheet    As Excel.Worksheet
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    Dim strFileName As String
    On Error GoTo Ert


    Set xlsApp = CreateObject("Excel.Application")
    Set xlsBook = xlsApp.Workbooks.Add
    Set xlsSheet = xlsBook.Worksheets(1)

    With xlsApp
        .Rows(1).Font.Bold = True
    End With

    strSql = "SELECT DISTINCT a.日期, a.大箱_CST,a.型号,a.lotid,WaferId = (STUFF((SELECT ',' +  Wafer FROM erptemp..gcexceltranslate WHERE a.LotID=lotid and a.大箱_CST=大箱_CST  AND a.入库备注=入库备注 order by Wafer FOR XML PATH('')), 1,  1, '')),sum(convert(INT,(a.片数))) as 片数,'华天' as Factory,a.入库备注,a.remark1 as 形式, a.remark3 as 厂内机种 FROM  erptemp..gcexceltranslate  a GROUP BY a.日期, a.大箱_CST,a.型号,a.lotid,a.入库备注,a.remark1 ,a.remark3 "


    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If SMR.RecordCount > 0 Then
    
        xlsSheet.Cells(1, 1) = "日期"
        xlsSheet.Cells(1, 2) = "大箱_CST"
        xlsSheet.Cells(1, 3) = "型号"
        xlsSheet.Cells(1, 4) = "lotid"
        xlsSheet.Cells(1, 5) = "WaferId"
        xlsSheet.Cells(1, 6) = "片数"
        xlsSheet.Cells(1, 7) = "Factory"
        xlsSheet.Cells(1, 8) = "入库备注"
        xlsSheet.Cells(1, 9) = "厂内机种"
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            xlsSheet.Cells(i + 1, 1) = Trim(SMR("日期"))
            xlsSheet.Cells(i + 1, 2) = Trim(SMR("大箱_CST"))
            xlsSheet.Cells(i + 1, 3) = Trim(SMR("型号"))
            xlsSheet.Cells(i + 1, 4) = Trim(SMR("lotid"))
            xlsSheet.Cells(i + 1, 5) = Trim(SMR("WaferId"))
            xlsSheet.Cells(i + 1, 6) = Trim(SMR("片数"))
            xlsSheet.Cells(i + 1, 7) = Trim(SMR("Factory"))
            xlsSheet.Cells(i + 1, 8) = Trim(SMR("入库备注"))
            xlsSheet.Cells(i + 1, 9) = Trim(SMR("厂内机种"))
            SMR.MoveNext
        Next
        With xlsSheet.Range("2:" & i)
            .horizontalAlignment = xlLeft
        End With
        xlsSheet.Range("A1").Select
        xlsApp.Columns.AutoFit
    
    End If
    SMR.Close
    Set SMR = Nothing
    
    xlsApp.Visible = True
    filepath_org = Trim(Text3.text)
    
    strFileName = Left(filepath_org, InStrRev(filepath_org, ".") - 1) & "_tostock" & Format(Now, "YYYYMMDDhhmmss") & Mid(filepath_org, InStrRev(filepath_org, "."), Len(filepath_org) - InStrRev(filepath_org, ".") + 1)

    xlsBook.SaveAs strFileName

    Set xlsApp = Nothing
    Set xlsSheet = Nothing
    Set xlsBook = Nothing

    MsgBox "转换完成", vbInformation, "提示"
Ert:
    MsgBox Err.DESCRIPTION & vbCrLf & "in 正式工程1..ExportToExcel ", vbExclamation + vbOKOnly, "Application Error"
    If Not (xlsApp Is Nothing) Then
        
        Set xlsApp = Nothing
        Set xlsSheet = Nothing
        Set xlsBook = Nothing

    End If
    

End Sub



Private Sub cmd_del_Click()


    Dim strcustlot As String
    Dim strWafer         As String
    Dim i           As Integer
    Dim DelCnt     As Integer
    Dim Delcustlot As String
    Dim DelWaferID As String
    Dim Delmsg As String

    DelCnt = 0
    Delcustlot = ""
    With fpS_del
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text <> "" Then
                If .text = 1 Then
                    .Col = 6    'custlot
                    Delcustlot = Trim(.text)
                    
                    .Col = 7    'custlot
                    DelWaferID = Trim(.text)
                    
                    
                    If Delmsg = "" Then
                        Delmsg = Delcustlot & DelWaferID
                    Else
                        Delmsg = Delmsg & "," & Delcustlot & DelWaferID
                    End If
                    
                    DelCnt = DelCnt + 1
                End If

            End If
        Next i
        If MsgBox("你确认要删除" & Delmsg & ",共" & DelCnt & "笔回货资料吗?", vbOKCancel, "提示") = vbCancel Then
            Exit Sub

        End If
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .text <> "" Then
                If .text = 1 Then
            
                    .Col = 6      'custlot
                    strcustlot = Trim(.text)
                
                    .Col = 7    'wafer
                    strWafer = Trim$(.text)
                    AddSql2 (" UPDATE erptemp..TSV_VT_History_sub SET flag =2,LASTUPDATE_BY='" & gUserName & "',LASTUPDATE_DATE=sysdatetime()    WHERE flag=1 and custlot = '" & strcustlot & "' and waferid='" & strWafer & "'")


                End If

            End If

        Next i
        
        .MaxRows = 0

    End With
   ' cmd_search_Click '查询

End Sub

Private Sub Cmd_exportWO_Click()
ExportToExcel_GCWO
End Sub






Private Sub Cmd_GCNewformat_Click()
    Dim strCust As String
    
    VTformat = "new"
    
    If CboCustomer.text = "" Then
        MsgBox "请先选择客户代码"
        Exit Sub
    
    End If
    '未作回货申请，不可上传新的回货资料
    If Get_SqlserverCnt("select * from erptemp..TSV_VT_History_sub where FLAG=1 AND  CUSTOMERSHORTNAME='" & Trim(CboCustomer.text) & "'") > 0 Then
        MsgBox "有回货资料未做回货申请，请先完成申请再上传！", vbInformation, "提示"
        Exit Sub
    End If
    
    shipid = CStr(GetVTID)
    If CStr(shipid) = "" Then
        MsgBox "获取回货单号异常,请重试", vbInformation, "提示"
        Exit Sub
    End If
    If UCase(Trim(CboCustomer.text)) = "GC" Then
        strCust = UCase(Trim(CboCustomer.text))
        UploadVTData_GC_New (strCust)
    Else
        MsgBox "此功能仅GC使用", vbInformation, "提示"
        Exit Sub
    End If
End Sub

Private Sub cmd_search_Click()
    Dim rs        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    strSql = "select 0,* from erptemp..TSV_VT_History_sub where flag=1 and custlot='" & Trim(Txt_lotid.text) & "'"
    If rs.State = adStateOpen Then SMR.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If rs.RecordCount > 0 Then
        With fpS_del
            Set .DataSource = rs
            For i = 1 To .MaxRows
                .Row = i
                .Col = -1
                .BackColor = &H8000000F
                .Row = i
                .Col = 1
                
                .SetText 1, 0, "选择"
                .CellType = CellTypeCheckBox
                .text = 0
                .TypeHAlign = TypeHAlignCenter
                .TypeVAlign = TypeVAlignCenter
                .Lock = True
                
                .ReDraw = True
    
            Next

    End With
        
        
    Else
    
        MsgBox "未查到此lot的上传记录", vbInformation, "提示"
    End If
    
End Sub











Private Sub fpS_del_Click(ByVal Col As Long, ByVal Row As Long)

If Col <> 1 Then Exit Sub
With fpS_del
    .Col = 1
    .Row = Row
    .Value = Abs(Val(.Value) - 1)

    If Val(.Value) = 1 Then
        .Row = Row
        .Col = -1
        .BackColor = &HC0C0FF
    Else
    
        .Row = Row
        .Col = -1
        .BackColor = &H8000000F
        
        
    End If
End With
End Sub


Private Sub Command1_Click()
    Command1.Enabled = False
    fpS.MaxRows = 0
    If UCase(CboCustomer.text) = "GC" Then
        createvtappllication_GC
    Else
        createvtappllication_KR
    End If
    Command1.Enabled = True
End Sub




Private Sub Command6_Click()

On Error Resume Next
Dim FName
    '帅选文件
    CommonDialog2.Filter = "EXCEL文件(*.xlsx)|*.xlsx|EXCEL文件(*.xls)|*.xls"
    
    CommonDialog2.ShowOpen
    '得到文件名
    FName = CommonDialog2.filename
    If FName <> "" Then
       Text3.text = FName
    End If


End Sub

Private Sub Command7_Click()
Dim strCust As String
VTformat = "new"
If CboCustomer.text = "" Then
    MsgBox "请先选择客户代码"
    Exit Sub

End If
'未作回货申请，不可上传新的回货资料
If Get_SqlserverCnt("select * from erptemp..TSV_VT_History_sub where FLAG=1 AND  CUSTOMERSHORTNAME='" & Trim(CboCustomer.text) & "'") > 0 Then
    MsgBox "有回货资料未做回货申请，请先完成申请再上传！", vbInformation, "提示"
    Exit Sub
End If

'未传WO，不可上传新的回货资料
' If Get_SqlserverCnt("select * from erptemp..TSV_VT_History_sub where FLAG_WO=1 AND  CUSTOMERSHORTNAME='" & Trim(CboCustomer.Text) & "'") > 0 Then
    ' MsgBox "有回货资料未做回货申请，请先完成申请再上传！", vbInformation, "提示"
    ' End Sub
' End If


'shipid = Get_OracleStr("select TSV_VT_SEQ.Nextval from dual")
shipid = CStr(GetVTID)
If CStr(shipid) = "" Then
    MsgBox "获取回货单号异常,请重试", vbInformation, "提示"
    Exit Sub
End If
If UCase(Trim(CboCustomer.text)) = "KR009" Then
    strCust = UCase(Trim(CboCustomer.text))
    UploadVTData_KR009 (strCust)
ElseIf UCase(Trim(CboCustomer.text)) = "GC" Then
    strCust = UCase(Trim(CboCustomer.text))
   ' UploadVTData_GC (strCust)
    UploadVTData_GC_New (strCust)
    
    
Else
    strCust = UCase(Trim(CboCustomer.text))
    UploadVTData (strCust)

End If

End Sub



Private Sub UploadVTData(customerTemp As String)

'上传资料

Dim source_batch_id_Temp As String
'上传OI的CSV
'处理文件名
If Text3.text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
End If
Dim dirName As String
Dim filename As String

'获取文件名
'    If InStrRev(Trim(Text2.Text), "\") > 0 Then
'        strFileName = Mid(Trim(Text2.Text), InStrRev(Trim(Text2.Text), "\") + 1)
'        dirName = Mid$(Trim(Text2.Text), 1, InStrRev(Trim(Text2.Text), "\"))
'    End If
    

'2012-06-27 jiayunzhang 修改读Excel的方式


'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 11 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If


Dim i As Integer
Dim j As Integer
Dim id As Long
Dim TEMP As String
Dim temp2 As String
Dim tempVal As String
   
SumCount = 0
BCResultFlag = False

 vtDataTemp.Created_ByTemp = gUserName

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
 
    TEMP = ""
    source_batch_id_Temp = ""
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
        
        If j = 1 Then
        
            vtDataTemp.SHIPDATETemp = Trim(tempVal)
            
        ElseIf j = 2 Then
            vtDataTemp.DeliveryNoTemp = Trim(tempVal)
            
       ElseIf j = 3 Then
            vtDataTemp.CustDeviceTemp = Trim(tempVal)
            
       ElseIf j = 4 Then
            vtDataTemp.CUSTLOTTemp = Trim(tempVal)
            
       ElseIf j = 5 Then
            vtDataTemp.goodDieQtyTemp = Trim(tempVal)
            
       ElseIf j = 6 Then
            vtDataTemp.ngDieQtyTemp = Trim(tempVal)
            
       ElseIf j = 7 Then
            vtDataTemp.TTLTemp = Trim(tempVal)
            
       ElseIf j = 8 Then
       
            vtDataTemp.NetWeightTemp = Trim(tempVal)
            
       ElseIf j = 9 Then
            
            vtDataTemp.GrossWeightTemp = Trim(tempVal)
            
       ElseIf j = 10 Then
            vtDataTemp.remarkTemp = Trim(tempVal)
            
        End If
        

    Next j

  

    '判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
    If (JudgeFlagVTData(vtDataTemp.DeliveryNoTemp, vtDataTemp.CUSTLOTTemp)) Then
       MsgBox "这笔已存在，无需上传!", vbInformation, "友情提示"
       GoTo NextRecord2

    End If


    Call AddVTCustomer(vtDataTemp, customerTemp)
    SumCount = SumCount + 1

    '上传到DB
NextRecord2:

Next i


     VBExcel.Application.DisplayAlerts = False '关闭文档不弹出提示框
     xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit




If SumCount > 0 Then
    MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"
    
    Else
        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"
            Exit Sub
        End If
    
End If





''读取CSV
'Dim source_batch_id_Temp As String
'Dim customerTemp As String
'
'customerTemp = "GC"
'
''上传OI的CSV
''处理文件名
'If Text3.Text = "" Then
'    MsgBox "先选择待上传的文件"
'    Exit Sub
'End If
'Dim dirName As String
'Dim FileName As String
'
''获取文件名
'    If InStrRev(Trim(Text3.Text), "\") > 0 Then
'        strfilename = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
'        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
'    End If
'
'Dim con As New ADODB.Connection
'Dim Rs As New ADODB.Recordset
'
'
'        con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
'        Rs.open "Select * From " & "[" & strfilename & "]", con, adOpenStatic, adLockReadOnly, adCmdText
'
'        Dim i As Integer
'        Dim j As Integer
'        Dim id As Long
'        Dim temp As String
'        Dim SumCount As Integer
'        Dim GCHeaderFlag As Boolean
'        SumCount = 0
'        Rs.MoveFirst
'
'        GCHeaderFlag = False
'
'        For i = 0 To Rs.RecordCount - 1
'            temp = ""
'            id = 0
'
'            vtDataTemp.SHIPDATETemp = Rs.fields(0).Value
'            vtDataTemp.StockNoTemp = Rs.fields(1).Value
'            vtDataTemp.DeliveryNoTemp = Rs.fields(2).Value
'            vtDataTemp.CustDeviceTemp = Rs.fields(3).Value
'            vtDataTemp.CUSTLOTTemp = Rs.fields(4).Value
'            vtDataTemp.WaferIdTemp = Rs.fields(5).Value
'            vtDataTemp.WLCSPDeviceTemp = Rs.fields(6).Value
'            vtDataTemp.WLCSPLOTTemp = Rs.fields(7).Value
'            vtDataTemp.goodDieQtyTemp = CLng(Rs.fields(8).Value)
'            vtDataTemp.NGDIEQTYTemp = CLng(Rs.fields(9).Value)
'            vtDataTemp.PackingLOTNoTemp = Rs.fields(10).Value
'            vtDataTemp.TTLTemp = IIf(IsNull(Rs.fields(11).Value), "", Rs.fields(11).Value)
'            vtDataTemp.WaferQtyInTemp = IIf(IsNull(Rs.fields(12).Value), "", Rs.fields(12).Value)
'            vtDataTemp.BatchTemp = Rs.fields(13).Value
'            vtDataTemp.SAPCodeTemp = Rs.fields(14).Value
'            vtDataTemp.WorkWeekTemp = IIf(IsNull(Rs.fields(15).Value), "", Rs.fields(15).Value)
'            vtDataTemp.CartonNoTemp = IIf(IsNull(Rs.fields(16).Value), "", Rs.fields(16).Value)
'            vtDataTemp.NetWeightTemp = IIf(IsNull(Rs.fields(17).Value), "", Rs.fields(17).Value)
'            vtDataTemp.GrossWeightTemp = IIf(IsNull(Rs.fields(18).Value), "", Rs.fields(18).Value)
'            vtDataTemp.RemarkTemp = IIf(IsNull(Rs.fields(19).Value), "", Rs.fields(19).Value)
'            vtDataTemp.Created_ByTemp = gUserName
'
'
'
'
'
''                '2013-12-05 jiayun add
''                '判断wo是否为空
''
''                If Trim(gcHeaderTemp.WO_NO) = "" Then
''
''                    MsgBox "WO_NO有空值，请确认！"
''                    Exit Sub
''
''                End If
''
''                '2012-11-07 jiayun 修改Good_Die_Qty 根据市场部规则
''
''            gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
''
''            '2013-12-27 jiayun add
''
''            If gcDetailTemp.Good_Die_Qty <= 0 Then
''                    MsgBox "请确认客户机种对应的Die数是否有维护好！"
''                    Exit Sub
''            End If
''
''
''            '2012-11-05 jiayun 修改 GC
''
''            '判断lotID在Header表中是否已存在
''
''            If (JudgeGCHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)) Then
''
''                If GCHeaderFlag = False Then
''        '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
''                End If
''
''                '2013-12-05 jiayun add 如果lotid,wo_no 已存在，则查询出id
''                '当lotid有隔行时，则查询上次的id
''
''                id = GetGCLotIDWOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)
''
''            Else
''            '上传到Header表中
''                '取目前DB最大的ID号
''                id = GetMaxID()
''                '2013-01-11 jiayun add 客户简称
''
''                If id = 0 Then
''                    MsgBox "DB主表ID生成失败1，请联系资讯！"
''                    Exit Sub
''
''                Else
''
''
''                    Call AddGCHeader(gcHeaderTemp, id, customerTemp)
''                    GCHeaderFlag = True
''
''                End If
''
''            End If
''
''
''            '判断lotID在Detail表中是否已存在
''
''            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_ID)) Then
''               MsgBox "GC 这笔：" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_ID & "已存在，无需上传!"
''
''            Else
''            '上传到Detail表中
''
''                   '2012-11-05 jiayun 修改 GCT
''
''
''                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
''
''
''                If id = 0 Then
''                    MsgBox "DB主表ID生成失败2，请联系资讯！"
''                    Exit Sub
''
''                Else
''                    Call AddGCDetail(gcDetailTemp, customerTemp, id)
''                    SumCount = SumCount + 1
''
''                End If
''
''
''            End If
''
'
'            Rs.MoveNext
'
'        Next i
'
'
'        If SumCount > 0 Then
'            MsgBox "已成功上传" & SumCount & "笔！"
'        End If

End Sub



Private Sub UploadVTData_KR009(customerTemp As String)
'上传资料
Dim source_batch_id_Temp As String
'上传OI的CSV
'处理文件名
If Text3.text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
End If
Dim dirName As String
Dim filename As String

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 18 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If


Dim i As Integer
Dim j As Integer
Dim id As Long
Dim TEMP As String
Dim temp2 As String
Dim tempVal As String
   


 SumCount = 0
 BCResultFlag = False

 vtDataTemp.Created_ByTemp = gUserName

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
 
    TEMP = ""
    source_batch_id_Temp = ""
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
        
       If j = 1 Then
            vtDataTemp.SHIPDATETemp = Trim(tempVal)
            
       ElseIf j = 2 Then
            vtDataTemp.StockNoTemp = Trim(tempVal)
            
       ElseIf j = 3 Then
            vtDataTemp.DeliveryNoTemp = Trim(tempVal)
            
       ElseIf j = 4 Then
            vtDataTemp.CustDeviceTemp = Trim(tempVal)
                 
       ElseIf j = 5 Then
            vtDataTemp.CUSTLOTTemp = Trim(tempVal)
            
       ElseIf j = 6 Then
            vtDataTemp.waferIdTemp = Trim(tempVal)
            
       ElseIf j = 7 Then
            
            vtDataTemp.goodDieQtyTemp = Trim(tempVal)
            
       ElseIf j = 8 Then
            vtDataTemp.ngDieQtyTemp = Trim(tempVal)
       
       ElseIf j = 9 Then
            vtDataTemp.PackingLOTNoTemp = Trim(tempVal)
                   
       ElseIf j = 10 Then
            vtDataTemp.TTLTemp = Trim(tempVal)
            
       ElseIf j = 11 Then
            vtDataTemp.WaferQtyInTemp = Trim(tempVal)
            
       ElseIf j = 12 Then
            vtDataTemp.BatchTemp = Trim(tempVal)
            
       ElseIf j = 13 Then
            vtDataTemp.SAPCodeTemp = Trim(tempVal)
            
       ElseIf j = 14 Then
            vtDataTemp.WorkWeekTemp = Trim(tempVal)
            
       ElseIf j = 15 Then
            vtDataTemp.CartonNoTemp = Trim(tempVal)
            
       ElseIf j = 16 Then
            vtDataTemp.NetWeightTemp = Trim(tempVal)
            
       ElseIf j = 17 Then
            vtDataTemp.GrossWeightTemp = Trim(tempVal)
            
       ElseIf j = 18 Then
            vtDataTemp.remarkTemp = Trim(tempVal)
    
       End If
        
    Next j

    '判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
    If (JudgeFlagVTData_ALL(vtDataTemp)) Then
       MsgBox "这笔已存在，无需上传!", vbInformation, "友情提示"
       GoTo NextRecord2

    End If

'    If (JudgeFlagVTData(vtDataTemp.DeliveryNoTemp, vtDataTemp.CUSTLOTTemp)) Then
'       MsgBox "这笔已存在，无需上传!", vbInformation, "友情提示"
'       GoTo NextRecord2
'    End If


    Call AddVTCustomer_KR009(vtDataTemp, customerTemp)
    SumCount = SumCount + 1

    '上传到DB
NextRecord2:

Next i


     VBExcel.Application.DisplayAlerts = False '关闭文档不弹出提示框
     xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit




If SumCount > 0 Then
    MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"
    
    Else
        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"
            Exit Sub
        End If
    
End If

End Sub


Private Sub UploadVTData_GC(customerTemp As String)
'上传资料
Dim source_batch_id_Temp As String
'上传OI的CSV
'处理文件名
If Text3.text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
End If
Dim dirName As String
Dim filename As String

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 7 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If


Dim i As Integer
Dim j As Integer
Dim id As Long
Dim TEMP As String
Dim temp2 As String
Dim tempVal As String
   


 SumCount = 0
 BCResultFlag = False

 vtDataTemp.Created_ByTemp = gUserName

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
 
    TEMP = ""
    source_batch_id_Temp = ""
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
        
        If j = 1 Then
            vtDataTemp.SHIPDATETemp = Trim(tempVal)
            
        ElseIf j = 2 Then
            vtDataTemp.StockNoTemp = Trim(tempVal)
            
       ElseIf j = 3 Then
            vtDataTemp.DeliveryNoTemp = Trim(tempVal)
            
       ElseIf j = 4 Then
            vtDataTemp.CUSTLOTTemp = Trim(tempVal)
            
       ElseIf j = 5 Then
            vtDataTemp.BatchTemp = Trim(tempVal)
            
       ElseIf j = 6 Then
            vtDataTemp.TTLTemp = Trim(tempVal)
            
       ElseIf j = 7 Then
            vtDataTemp.remarkTemp = Trim(tempVal)
        
       End If
        
    Next j

    If (JudgeFlagVTData_GC(vtDataTemp)) Then
       MsgBox "这笔已存在，无需上传!", vbInformation, "友情提示"
       GoTo NextRecord2

    End If

    Call AddVTCustomer_GC(vtDataTemp, customerTemp)
    SumCount = SumCount + 1

    '上传到DB
NextRecord2:

Next i
     xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

If SumCount > 0 Then
    MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"
    
    Else
        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"
            Exit Sub
        End If
    
End If

End Sub



Private Sub UploadVTData_GC_New(customerTemp As String)
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    Dim errormsg   As String
    
    If GCformatTranslate = False Then
        Exit Sub
    End If
    '20200306merry核对二级代码前两位与WLA是否一致
    
    strSql = "SELECT DISTINCT a.lotid  FROM erptemp..gcexceltranslate a,erpbase..tblCustomerOI b Where a.LOTID = b.SOURCE_BATCH_ID And Left(a.二级代码, 2) <> Left(b.IMAGER_CUSTOMER_REV, 2)"
    Set SMR = Get_SqlserveRs(strSql)
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            errormsg = errormsg & "," & SMR("lotid")
            SMR.MoveNext
        Next
      '  MsgBox "回货数据有误" & errormsg & "二级代码与WLA不一致", vbInformation, "提示"
    '    Exit Sub
    End If
    
    

    SumCount = 0
    vtDataTemp.Created_ByTemp = gUserName
    
    errormsg = ""
    '卡控1，已经上传过
    strSql = "select rtrim(lotid)+rtrim(wafer) as waferid from  erptemp..gcexceltranslate  where  rtrim(lotid)+rtrim(wafer) in (select rtrim(CUSTLOT)+rtrim(WAFERID) from  erptemp..TSV_VT_History_sub where (flag=1 or flag=2))  "
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            errormsg = errormsg & "," & SMR("waferid")
            SMR.MoveNext
        Next
        MsgBox "回货数据有误" & errormsg & "已上传过", vbInformation, "提示"
        Exit Sub
    End If
    '卡控2：remark1栏位只能有两种情况，或为空，或为降main,空是WLT, 降main为转Normal
    If Get_SqlserverCnt("select  REMARK1 from  ERPTEMP..TSV_VT_History_sub  WHERE  FLAG_WO =1 AND Customershortname='GC' And isnull(REMARK1,'')<>'' and replace(isnull(REMARK1,''),' ','')<>'降MAIN'") > 0 Then

        MsgBox "入库备注栏位只能有'降Main'字样，请检查回货资料", vbInformation, "提示"
        Exit Sub
    End If
    '卡控3：二级代码未维护
    strSql = " select DISTINCT a.CUSTDEVICE  from  ERPTEMP..TSV_VT_History_sub a WHERE a.FLAG_WO =1 AND a.Customershortname='GC' AND  a.CUSTDEVICE  +'-3' NOT IN (SELECT b.客户机种名 FROM erptemp..GcCode_Reference  b where  b.制程='WLT')"
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            errormsg = errormsg & "," & SMR("CUSTDEVICE")
            SMR.MoveNext
        Next
        MsgBox errormsg & "未维护WLT二级代码", vbInformation, "提示"
        Exit Sub
    End If

    strSql = " select DISTINCT a.CUSTDEVICE  from  ERPTEMP..TSV_VT_History_sub a WHERE a.FLAG_WO =1 AND a.Customershortname='GC' AND  a.CUSTDEVICE  +'-3' NOT IN (SELECT b.客户机种名 FROM erptemp..GcCode_Reference  b where  b.制程='转normal')"
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            errormsg = errormsg & "," & SMR("CUSTDEVICE")
            SMR.MoveNext
        Next
        MsgBox errormsg & "未维护转normal二级代码", vbInformation, "提示"
        Exit Sub
    End If

    '开始上传
    strSql = "SELECT DISTINCT a.日期, a.大箱_CST,a.型号,a.lotid,WaferId = (STUFF((SELECT ',' +  Wafer FROM erptemp..gcexceltranslate WHERE a.LotID=lotid    and a.大箱_CST=大箱_CST AND a.入库备注=入库备注 order by Wafer FOR XML PATH('')), 1,  1, '')),sum(convert(INT,(a.片数))) as 片数,'华天' as Factory,a.入库备注,a.remark1 as 形式,a.remark3 as 厂内机种 FROM  erptemp..gcexceltranslate  a GROUP BY a.日期, a.大箱_CST,a.型号,a.lotid,a.入库备注,a.remark1,a.remark3"
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            vtDataTemp.SHIPDATETemp = Trim(SMR("日期"))
            vtDataTemp.StockNoTemp = Trim(SMR("大箱_CST"))
            vtDataTemp.DeliveryNoTemp = Trim(SMR("型号"))
            vtDataTemp.CUSTLOTTemp = Trim(SMR("lotid"))
            vtDataTemp.BatchTemp = Trim(SMR("WaferId"))
            vtDataTemp.TTLTemp = Trim(SMR("片数"))
            vtDataTemp.remarkTemp = Trim(SMR("入库备注"))
            vtDataTemp.type = Trim(SMR("形式"))
            vtDataTemp.htdevice = Trim(SMR("厂内机种"))
            
            If (JudgeFlagVTData_GC(vtDataTemp)) Then
               MsgBox "这笔已存在，无需上传!", vbInformation, "友情提示"
            Else
                Call AddVTCustomer_GC(vtDataTemp, customerTemp)
                SumCount = SumCount + 1
            End If
            SMR.MoveNext
        Next
    End If
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"
        ExportToExcel
       ' ExportToExcel_GCWO
    Else
        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"
            Exit Sub
        End If
        
    End If
    SMR.Close
    Set SMR = Nothing
    


End Sub





Private Sub Command8_Click()
Dim customerStr As String

If Trim(CboCustomer.text) = "" Then
    MsgBox "请先选择客户，再导出报表！", vbInformation, "友情提示"
    Exit Sub
ElseIf UCase(Trim(CboCustomer.text)) = "KR009" Then
    customerStr = UCase(Trim(CboCustomer.text))
    ExporToExcel ("  select  SHIPDATE,StockNo,DELIVERYNO,CUSTDEVICE,CUSTLOT,waferId,GOODDIEQTY," & _
         " NGDIEQTY,PackingLOTNo,TTL,WaferQtyIn,Batch,SAPCode,WorkWeek,CartonNo,NETWEIGHT,GROSSWEIGHT,REMARK,回货单号 " & _
                " From TSV_VT_History where customershortname='" & customerStr & "' order by SHIPDATE  ")
ElseIf UCase(Trim(CboCustomer.text)) = "GC" Then
    customerStr = UCase(Trim(CboCustomer.text))
    ExporToExcel ("  select SHIPDATE as ""C/Not"" ,StockNo as ""箱号"",DELIVERYNO as ""型号"",CUSTLOT as ""LOT-ID"",Batch as ""wafer-Id"",TTL as ""数量"",REMARK as ""供应商"" ,回货单号" & _
                " From TSV_VT_History where customershortname='" & customerStr & "' order by id  ")
'     ExporToExcel ("  select SHIPDATE ,StockNo ,DELIVERYNO ,CUSTLOT ,Batch ,TTL ,REMARK " & _
'               " From TSV_VT_History where customershortname='" & customerStr & "' order by id  ")
Else

customerStr = UCase(Trim(CboCustomer.text))

ExporToExcel ("  select id, SHIPDATE,DELIVERYNO,CUSTDEVICE,CUSTLOT,GOODDIEQTY,NGDIEQTY,TTL,NETWEIGHT,GROSSWEIGHT,REMARK,回货单号" & _
               "  Flag, Created_By, created_date " & _
               " From TSV_VT_History where customershortname='" & customerStr & "' order by id  ")
End If


End Sub

Private Sub Form_Load()
IniCustomerName

End Sub


Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CboCustomer.RowSource = mainItemRS
CboCustomer.ListField = mainItemRS("productname").name
CboCustomer.BoundColumn = mainItemRS("PID").name

End Sub

Private Function JudgeFlagVTData_ALL(TEMP As VTData) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from  TSV_VT_History where customershortname = 'KR009' and SHIPDATE = '" & TEMP.SHIPDATETemp & "' and StockNo = '" & TEMP.StockNoTemp & "' and  DELIVERYNO = '" & _
"" & TEMP.DeliveryNoTemp & "' and CUSTDEVICE ='" & TEMP.CustDeviceTemp & "' and  CUSTLOT ='" & TEMP.CUSTLOTTemp & "' and waferId ='" & TEMP.waferIdTemp & _
"' and PackingLOTNo = '" & TEMP.PackingLOTNoTemp & _
 "'and WaferQtyIn = '" & TEMP.WaferQtyInTemp & "'and Batch = '" & TEMP.BatchTemp & "' and SAPCode ='" & TEMP.SAPCodeTemp & "'"

'cmdStr = "select * from  TSV_VT_History where customershortname = 'KR009' and SHIPDATE = '" & TEMP.SHIPDATETemp & "' and StockNo = '" & TEMP.StockNoTemp & "' and  DELIVERYNO = '" & _
'"" & TEMP.DeliveryNoTemp & "' and CUSTDEVICE ='" & TEMP.CustDeviceTemp & "' and  CUSTLOT ='" & TEMP.CUSTLOTTemp & "' and waferId ='" & TEMP.waferIdTemp & _
'"' and GOODDIEQTY = '" & TEMP.goodDieQtyTemp & "' and NGDIEQTY = '" & TEMP.ngDieQtyTemp & "' and PackingLOTNo = '" & TEMP.PackingLOTNoTemp & "' and TTL = '" & _
'"" & TEMP.TTLTemp & "'and WaferQtyIn = '" & TEMP.WaferQtyInTemp & "'and Batch = '" & TEMP.BatchTemp & "' and SAPCode ='" & TEMP.SAPCodeTemp & "' and WorkWeek= '" & _
'"" & TEMP.WorkWeekTemp & "' and CartonNo ='" & TEMP.CartonNoTemp & "' and NETWEIGHT ='" & TEMP.NetWeightTemp & "' and GROSSWEIGHT ='" & TEMP.GrossWeightTemp & _
'"'and REMARK = '" & TEMP.remarkTemp & "'"

slectResult = QueryStr(cmdStr)

JudgeFlagVTData_ALL = slectResult
End Function

Private Function JudgeFlagVTData_GC(TEMP As VTData) As Boolean
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "select * from  TSV_VT_History where customershortname = 'GC' and SHIPDATE = '" & TEMP.SHIPDATETemp & "' and StockNo = '" & TEMP.StockNoTemp & "' and  DELIVERYNO = '" & _
"" & TEMP.DeliveryNoTemp & "'and  CUSTLOT ='" & TEMP.CUSTLOTTemp & "'and Batch = '" & TEMP.BatchTemp & "' and REMARK ='" & TEMP.remarkTemp & "'"

slectResult = QueryStr(cmdStr)

JudgeFlagVTData_GC = slectResult

End Function

' add VT
Private Sub AddVTCustomer(TEMP As VTData, customerTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String
Dim strid As String
Dim strlot As String
Dim strWafer As String
'添加导入Sqlserver
On Error GoTo DealError
strid = Get_OracleStr("select tbl_tsv_VTData_seq.Nextval from dual")

Cnn.BeginTrans

cmdStr = "insert into  TSV_VT_History (" & _
" SHIPDATE,DELIVERYNO,CUSTDEVICE,CUSTLOT,GOODDIEQTY," & _
" NGDIEQTY,TTL,NETWEIGHT,GROSSWEIGHT,REMARK," & _
" Flag , Created_By, created_date,id,customershortname,回货单号) values  " & _
" ('" & TEMP.SHIPDATETemp & "','" & TEMP.DeliveryNoTemp & "','" & TEMP.CustDeviceTemp & "','" & TEMP.CUSTLOTTemp & "'," & _
" " & TEMP.goodDieQtyTemp & "," & TEMP.ngDieQtyTemp & "," & _
" '" & TEMP.TTLTemp & "'," & _
" '" & TEMP.NetWeightTemp & "','" & TEMP.GrossWeightTemp & "','" & TEMP.remarkTemp & "'," & _
" 'Y','" & TEMP.Created_ByTemp & "', sysdate ," & strid & ",'" & customerTemp & "','" & shipid & "')"

                
AddSql (cmdStr)
    strlot = ""
    strWafer = ""
    strlot = Split(TEMP.CUSTLOTTemp, "-")(0)
    If InStr(TEMP.CUSTLOTTemp, "-") Then
        strWafer = Split(TEMP.CUSTLOTTemp, "-")(1)
    Else
        strWafer = ""
    End If
 
    If Get_SqlserverCnt("select * from erptemp..TSV_VT_History_sub where FLAG=1 AND  CUSTLOT='" & strlot & "' and WAFERID='" & strWafer & "'") > 0 Then
   
         strSql = "update erptemp..TSV_VT_History_sub set GOODDIEQTY=GOODDIEQTY+" & TEMP.goodDieQtyTemp & ",NGDIEQTY=NGDIEQTY+" & TEMP.ngDieQtyTemp & " where CUSTLOT='" & strlot & "' and WAFERID='" & strWafer & "'"
         
        AddSql2 (strSql)
    Else
       strSql = " INSERT INTO erptemp..TSV_VT_History_sub(SHIPDATE,StockNo,CUSTOMERSHORTNAME,CUSTDEVICE,CUSTLOT,WAFERID,GOODDIEQTY,NGDIEQTY,FLAG,CREATED_BY,CREATED_DATE,ID,回货单号,FLAG_WO)" & _
            " Values ('" & TEMP.SHIPDATETemp & "','" & TEMP.DeliveryNoTemp & "','" & customerTemp & "','" & TEMP.CustDeviceTemp & "','" & strlot & "','" & strWafer & "' " & _
            ",'" & TEMP.goodDieQtyTemp & "','" & TEMP.ngDieQtyTemp & "',1,'" & TEMP.Created_ByTemp & "',sysdatetime() , " & strid & ",'" & shipid & "',1)"

        AddSql2 (strSql)
    End If
 

 
 
Cnn.CommitTrans

Exit Sub
DealError:
MsgBox TEMP.CUSTLOTTemp & "未成功上传，请确认", vbInformation, "提示"
Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True


End Sub

Private Sub AddVTCustomer_KR009(TEMP As VTData, customerTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String
Dim strid As String

'添加导入Sqlserver
On Error GoTo DealError
strid = Get_OracleStr("select tbl_tsv_VTData_seq.Nextval from dual")
    
Cnn.BeginTrans

cmdStr = "insert into  TSV_VT_History (" & _
" SHIPDATE,StockNo,DELIVERYNO,CUSTDEVICE,CUSTLOT,waferId,GOODDIEQTY," & _
" NGDIEQTY,PackingLOTNo,TTL,WaferQtyIn,Batch,SAPCode,WorkWeek,CartonNo,NETWEIGHT,GROSSWEIGHT,REMARK," & _
" Flag , Created_By, created_date,id,customershortname,回货单号) values  " & _
" ('" & TEMP.SHIPDATETemp & "','" & TEMP.StockNoTemp & "','" & TEMP.DeliveryNoTemp & "','" & TEMP.CustDeviceTemp & "','" & TEMP.CUSTLOTTemp & "','" & _
"" & TEMP.waferIdTemp & "','" & TEMP.goodDieQtyTemp & "','" & TEMP.ngDieQtyTemp & "','" & TEMP.PackingLOTNoTemp & "'," & _
" '" & TEMP.TTLTemp & "','" & TEMP.WaferQtyInTemp & "','" & TEMP.BatchTemp & "','" & TEMP.SAPCodeTemp & "','" & TEMP.WorkWeekTemp & "'," & _
" '" & TEMP.CartonNoTemp & "','" & TEMP.NetWeightTemp & "','" & TEMP.GrossWeightTemp & "','" & TEMP.remarkTemp & "'," & _
" 'Y','" & TEMP.Created_ByTemp & "',sysdate," & strid & ",'" & customerTemp & "','" & shipid & "')"


AddSql (cmdStr)
   If Get_SqlserverCnt("select * from erptemp..TSV_VT_History_sub where FLAG=1 AND CUSTLOT='" & TEMP.CUSTLOTTemp & "' and WAFERID='" & TEMP.waferIdTemp & "'") > 0 Then
   
         strSql = "update erptemp..TSV_VT_History_sub set GOODDIEQTY=GOODDIEQTY+" & TEMP.goodDieQtyTemp & ",NGDIEQTY=NGDIEQTY+" & TEMP.ngDieQtyTemp & " where CUSTLOT='" & TEMP.CUSTLOTTemp & "' and WAFERID='" & TEMP.waferIdTemp & "'"
         
        AddSql2 (strSql)
    Else
       strSql = " INSERT INTO erptemp..TSV_VT_History_sub(SHIPDATE,StockNo,CUSTOMERSHORTNAME,CUSTDEVICE,CUSTLOT,WAFERID,GOODDIEQTY,NGDIEQTY,FLAG,CREATED_BY,CREATED_DATE,ID,回货单号,FLAG_WO)" & _
            " Values ('" & TEMP.SHIPDATETemp & "','" & TEMP.StockNoTemp & "','" & customerTemp & "','" & TEMP.CustDeviceTemp & "','" & TEMP.CUSTLOTTemp & "','" & TEMP.waferIdTemp & "' " & _
            ",'" & TEMP.goodDieQtyTemp & "','" & TEMP.ngDieQtyTemp & "',1" & _
             ",'" & TEMP.Created_ByTemp & "',sysdatetime() , " & strid & ",'" & shipid & "',1)"
 
    
        AddSql2 (strSql)
    End If
 
Cnn.CommitTrans

Exit Sub
DealError:
MsgBox TEMP.CUSTLOTTemp & "未成功上传，请确认", vbInformation, "提示"
Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True
End Sub

Private Sub AddVTCustomer_GC(TEMP As VTData, customerTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String
Dim i As Integer
Dim strWafer As String
Dim strHtDevice As String
Dim strid As String
Dim strtype As String

'添加导入Sqlserver
On Error GoTo DealError
strid = Get_OracleStr("select tbl_tsv_VTData_seq.Nextval from dual")
Cnn.BeginTrans

cmdStr = "insert into  TSV_VT_History (" & _
" SHIPDATE,StockNo,DELIVERYNO,CUSTLOT,Batch,TTL,REMARK," & _
" Flag,Created_By, created_date,id,customershortname,回货单号) values  " & _
" ('" & TEMP.SHIPDATETemp & "','" & TEMP.StockNoTemp & "','" & TEMP.DeliveryNoTemp & "','" & TEMP.CUSTLOTTemp & "','" & _
"" & TEMP.BatchTemp & "','" & TEMP.TTLTemp & "','" & TEMP.remarkTemp & "'," & _
" 'Y','" & TEMP.Created_ByTemp & "',sysdate, " & strid & ",'" & customerTemp & "','" & shipid & "')"

AddSql (cmdStr)

 For i = 0 To UBound(Split(TEMP.BatchTemp, ","))
    strWafer = Split(TEMP.BatchTemp, ",")(i)
    strSql = " INSERT INTO erptemp..TSV_VT_History_sub(SHIPDATE,StockNo,CUSTOMERSHORTNAME,CUSTDEVICE,CUSTLOT,WAFERID,FLAG,CREATED_BY,CREATED_DATE,ID,回货单号,FLAG_WO,REMARK1,REMARK2,REMARK3)" & _
            " Values ('" & TEMP.SHIPDATETemp & "','" & TEMP.StockNoTemp & "','" & customerTemp & "','" & TEMP.DeliveryNoTemp & "','" & TEMP.CUSTLOTTemp & "','" & strWafer & "',1" & _
             ",'" & TEMP.Created_ByTemp & "',sysdatetime() , " & strid & ",'" & shipid & "',1,'" & TEMP.remarkTemp & "','" & TEMP.type & "','" & TEMP.htdevice & "')"
    AddSql2 (strSql)

 Next
Cnn.CommitTrans


Exit Sub
DealError:
MsgBox TEMP.CUSTLOTTemp & "未成功上传，请确认", vbInformation, "提示"
Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True
End Sub

Private Sub init_vtDataTemp(TEMP As VTData)

TEMP.StockNoTemp = ""
TEMP.DeliveryNoTemp = ""
TEMP.CustDeviceTemp = ""
TEMP.CUSTLOTTemp = ""
TEMP.waferIdTemp = ""
TEMP.WLCSPDeviceTemp = ""
WLCSPLOTTemp = ""
TEMP.goodDieQtyTemp = ""
TEMP.ngDieQtyTemp = ""
TEMP.PackingLOTNoTemp = ""
TEMP.TTLTemp = ""
TEMP.WaferQtyInTemp = ""
TEMP.BatchTemp = ""
TEMP.SAPCodeTemp = ""
TEMP.WorkWeekTemp = ""
TEMP.CartonNoTemp = ""
TEMP.NetWeightTemp = ""
TEMP.GrossWeightTemp = ""
TEMP.remarkTemp = ""
TEMP.Created_ByTemp = ""

End Sub


Function createvtappllication_GC()

Dim SMR        As New ADODB.Recordset
Dim rs        As New ADODB.Recordset
Dim strSql As String
Dim strLotList As String
Dim strlot As String
Dim RequestNo As String
Dim strXH As String
Dim strxh_big As String
Dim strgdh As String
Dim strLCK As String
Dim strlps As String
Dim strbls As String
Dim strzcbls As String
Dim strid As String
Dim strmbkf As String
Dim strKF As String
Dim strmatcode As String
Dim strCustCode As String
Dim i As Integer
Dim j As Integer

Dim errormsg As String
errormsg = ""
createvtappllication_GC = False



    If SMR.State = adStateOpen Then SMR.Close
    strSql = "select distinct CUSTLOT from erptemp..TSV_VT_History_sub  where  customershortname='" & Trim(CboCustomer.text) & "' and  flag=1 "
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            If strLotList = "" Then
                strLotList = Trim(SMR("CUSTLOT"))
            Else
                strLotList = strLotList & "," & Trim(SMR("CUSTLOT"))
            End If
            SMR.MoveNext
        Next
    End If
    If strLotList = "" Then
        MsgBox "没有需要申请的回货资料，请确认", vbInformation, "提示"
        Exit Function
    End If
    
    
    For i = 0 To UBound(Split(strLotList, ","))
         strlot = Split(strLotList, ",")(i)
        '已提交过申请
        If SMR.State = adStateOpen Then SMR.Close
        strSql = "select a.CUSTLOT,a.WAFERID from erptemp..TSV_VT_History_sub a where rtrim(a.CUSTLOT)='" & strlot & "' and  a.flag=1 and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID) in (SELECT DISTINCT rtrim(b.WAFER) FROM erptemp..tblstockdb_temp  a,erptemp..tblstockdbsub_temp  b WHERE a.FLAG=1 AND a.ORDER_NUM=b.ORDER_NUM AND a.ITEM=b.ITEM ) "
        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            SMR.MoveFirst
            For j = 1 To SMR.RecordCount
                errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("CUSTLOT"))
                SMR.MoveNext
            Next
            MsgBox "回货资料有误" & errormsg & "已提交回货申请", vbInformation, "提示"
            Exit Function
        End If
        '不在72仓
        errormsg = ""
        If SMR.State = adStateOpen Then SMR.Close
        strSql = "select a.CUSTLOT,a.WAFERID from erptemp..TSV_VT_History_sub a where rtrim(a.CUSTLOT)='" & strlot & "' and    a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID) not in (select rtrim(replace(流程卡编号,'+','')) from erpdata..tblstocknumsub where 库房编号='72' and 合格标记=0 and 数量>0 ) "
        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            SMR.MoveFirst
            For j = 1 To SMR.RecordCount
                errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("WAFERID"))
                SMR.MoveNext
            Next
            MsgBox "回货资料有误" & errormsg & "不在72仓", vbInformation, "提示"
            Exit Function
        End If
        ' errormsg = ""
        ' If SMR.State = adStateOpen Then SMR.Close
        ' strSql = "select a.CUSTLOT,a.WAFERID,isnull(a.GOODDIEQTY,0),b.数量 from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where  rtrim(a.CUSTLOT)='" & strlot & "' and   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. 流程卡编号,'+','')) and b.库房编号='72' and b.合格标记=0 and isnull(a.GOODDIEQTY,0)>b. 数量 "
        ' SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        ' If SMR.RecordCount > 0 Then
            ' SMR.MoveFirst
            ' For j = 1 To SMR.RecordCount
                ' errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("WAFERID"))
                ' SMR.MoveNext
            ' Next
            ' MsgBox "回货资料有误" & errormsg & ",回货数量超出72仓库存", vbInformation, "提示"
            ' Exit Function
        ' End If
    Next
    RequestNo = GetID()
    For i = 0 To UBound(Split(strLotList, ","))
         strlot = Split(strLotList, ",")(i)
         strSql = "select b.箱号,b.流程卡编号,b.工单号,b.数量,b.id from erptemp..TSV_VT_History_sub a  ,erpdata..tblstocknumsub b where  a.flag=1 and  a.CUSTLOT='" & strlot & "' and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b.流程卡编号,'+','')) and b.数量>0 and b.合格标记=0 order by a.WAFERID"
         If SMR.State = adStateOpen Then SMR.Close
         SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            SMR.MoveFirst
            For j = 1 To SMR.RecordCount
                If j = 1 Then
                    strSql = " select top 1 a.原仓库 from erpdata..tblStockdb a, erpdata..tblStockdbsub b where a.id=b.id and  rtrim(b. 流程卡编号)='" & Trim(SMR("流程卡编号")) & "' and a.目标仓库='72'"
                    strmbkf = GetSqlServerStr(strSql)
                End If
                strXH = Trim(SMR("箱号"))     '箱号
                strxh_big = ""   '大箱号
                strgdh = Trim(SMR("工单号"))       '工单号
                strLCK = Trim(SMR("流程卡编号"))    '流程卡编号
                strlps = Trim(SMR("数量"))        '良品数
                strbls = "0"     '不良品数
                strzcbls = "0"     '制程不良数
                strid = Trim(SMR("id"))
                If Get_SqlserverCnt("select * from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid) > 0 Then
                    strSql = "select ITEM from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid
                    intitem = GetSqlServerStr(strSql)
                Else
                    strSql = "select isnull(max(ITEM),0) from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'"
                    intitem = GetSqlServerStr(strSql) + 1

                    If rs.State = adStateOpen Then rs.Close
                    strSql = " select 库房编号,物料编号,客户代码,isnull(单价,0) from erpdata..tblStockNum where id=" & strid
                    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                    If rs.RecordCount = 1 Then
                        rs.MoveFirst
                        strKF = Trim(rs("库房编号"))
                        strmatcode = Trim(rs("物料编号"))
                        strCustCode = Trim(rs("客户代码"))
                    
                    End If
                    
            
                  '上传主表

                 '调拨编号,序号,物料编号, 调拨数量,原仓库,目标仓库,申请人员,申请时间,审核人员,审核时间, 申请部门,状态,REMARK1,REMARK2,REMARK3,REMARK4,REMARK5,ID
                    strSql = "insert into erptemp..tblstockdb_temp(ORDER_NUM,ITEM, MATERIALS,QTY,FORMER, DESTINATION, APPLICANT, APPLICATION_TIME, AUDITOR, AUDIT_TIME, DEPT, FLAG,ID,REMARK1) values( " & _
                    "'" & RequestNo & "'," & intitem & ",'" & strmatcode & "'," & 0 & ",'" & strKF & "','" & strmbkf & "','" & gUserName & "',sysdatetime(),'','','',1," & strid & ",'')"
                
                    AddSql2 (strSql)
                   
                    
                End If
                            
                '上传子表
                
                '调拨编号, 序号, 箱号, 流程卡编号, 工单号, 合格数, 制程不良数, 来料不良数, ID
                 strSql = "insert into erptemp..tblstockdbsub_temp(ORDER_NUM,ITEM,WAFER,LOT,GOOD_DIE,BAD1_DIE,BAD2_DIE,ID,REMARK1,QBOX) values( " & _
                "'" & RequestNo & "'," & intitem & ",'" & strLCK & "','" & strgdh & "'," & strlps & "," & strbls & "," & strzcbls & "," & strid & ",'" & strxh_big & "','" & strXH & "')"
              
                AddSql2 (strSql)
                
                'update主表数量
                strSql = "Update erptemp..tblstockdb_temp set QTY =QTY+" & Val(strlps) + Val(strbls) + Val(strzcbls) & " where ORDER_NUM='" & RequestNo & "' and ITEM=" & intitem
               
                AddSql2 (strSql)

                
                
                SumCount = SumCount + 1
                SMR.MoveNext
            Next
            strSql = "Update erptemp..TSV_VT_History_sub set flag=2 where CUSTLOT='" & strlot & "'"
               
             AddSql2 (strSql)
        End If
    Next

   
    If SumCount > 0 Then
        MsgBox SumCount & "笔记录申请成功", vbInformation, "提示"
        Txt_sqdh.text = RequestNo
        If SMR.State = adStateOpen Then SMR.Close
       ' strSql = "select  CUSTLOT, WAFERID,GOODDIEQTY,NGDIEQTY  from erptemp..TSV_VT_History_sub  where  flag=1  and CUSTOMERSHORTNAME='" & Trim(CboCustomer.Text) & "' order by CUSTLOT,WAFERID "
        strSql = "select *  from erptemp..tblstockdbsub_temp  where  ORDER_NUM='" & RequestNo & "'"
        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            With fpS
                Set .DataSource = SMR
            End With

        End If
            End If
    

createvtappllication_GC = True


End Function


Function createvtappllication_KR()

    Dim SMR        As New ADODB.Recordset
    Dim rs        As New ADODB.Recordset
    Dim strSql As String

    Dim strlot As String
    Dim RequestNo_Gooddie As String
    Dim RequestNo_Ngdie As String
    Dim RequestNo As String
    Dim strXH As String
    Dim strxh_big As String
    Dim strgdh As String
    Dim strLCK As String
    Dim strlps As String
    Dim strbls As String
    Dim strzcbls As String
    Dim strid As String
    Dim strmbkf As String
    Dim strKF As String
    Dim strmatcode As String
    Dim strCustCode As String
    Dim i As Integer
    Dim j As Integer
    Dim SumCount As Integer
    Dim errormsg As String
    Dim strLotList As String
    errormsg = ""
    createvtappllication_KR = False

    If CboCustomer.text = "" Then
        MsgBox "请先选择客户代码"
        Exit Function

    End If

    If SMR.State = adStateOpen Then SMR.Close
    strSql = "select distinct CUSTLOT from erptemp..TSV_VT_History_sub  where customershortname='" & Trim(CboCustomer.text) & "' and  flag=1 "

    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            If strLotList = "" Then
                strLotList = Trim(SMR("CUSTLOT"))
            Else
                strLotList = strLotList & "," & Trim(SMR("CUSTLOT"))
            End If
            SMR.MoveNext
        Next
        
    End If

    'KR001,KR009有不良品，需调至不同的仓别，所以分两次申请
    RequestNo_Gooddie = GetID()
    RequestNo_Ngdie = Left(RequestNo_Gooddie, Len(RequestNo_Gooddie) - 1) & Right(RequestNo_Gooddie, 1) + 1
     For i = 0 To UBound(Split(strLotList, ","))
         strlot = Split(strLotList, ",")(i)
         If CheckData(strlot) = False Then
             Exit Function
         End If
    Next

     If UCase(Trim(CboCustomer.text)) = "KR001" Or UCase(Trim(CboCustomer.text)) = "KR009" Then
     
        '第一次拆箱，按gooddie数量拆
         RequestNo = RequestNo_Gooddie
         boxsplit (1)
         Call CommitApplication(RequestNo, 1, strLotList)
         
         '第二次拆箱，按ngdie数量拆
        ' RequestNo = RequestNo_Ngdie
        ' boxsplit (2)
        ' Call CommitApplication(RequestNo, 2,strLotList)
     Else
        
     End If
   
createvtappllication_KR = True


End Function


Function CheckData(strlot)
Dim SMR        As New ADODB.Recordset
Dim rs        As New ADODB.Recordset
Dim strSql As String
Dim errormsg As String

Dim j As Integer

CheckData = False
    '已提交过申请
    ' If SMR.State = adStateOpen Then SMR.Close
    ' strSql = "select a.CUSTLOT,a.WAFERID from erptemp..TSV_VT_History_sub a where rtrim(a.CUSTLOT)='" & strlot & "' and  a.flag=1 and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID) in (SELECT DISTINCT rtrim(b.WAFER) FROM erptemp..tblstockdb_temp  a,erptemp..tblstockdbsub_temp  b WHERE a.FLAG=1 AND a.ORDER_NUM=b.ORDER_NUM AND a.ITEM=b.ITEM ) "
    ' SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    ' If SMR.RecordCount > 0 Then
        ' SMR.MoveFirst
        ' For j = 1 To SMR.RecordCount
            ' errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("CUSTLOT"))
            ' SMR.MoveNext
        ' Next
        ' MsgBox "回货资料有误" & errormsg & "已提交回货申请", vbInformation, "提示"
        ' Exit Function
    ' End If
    '不在72仓
    errormsg = ""
    If SMR.State = adStateOpen Then SMR.Close
    strSql = "select a.CUSTLOT,a.WAFERID from erptemp..TSV_VT_History_sub a where  a.flag=1 and rtrim(a.CUSTLOT)='" & strlot & "'  and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID) not in (select rtrim(replace(流程卡编号,'+','')) from erpdata..tblstocknumsub where 库房编号='72' and 合格标记=0 and 数量>0 ) "
 
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For j = 1 To SMR.RecordCount
            errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("WAFERID"))
            SMR.MoveNext
        Next
        MsgBox "回货资料有误" & errormsg & "不在72仓", vbInformation, "提示"
        Exit Function
    End If
    '回货数量超出72仓库存
    If UCase(Trim(CboCustomer.text)) = "KR001" Or UCase(Trim(CboCustomer.text)) = "KR009" Then
        
        errormsg = ""
        If SMR.State = adStateOpen Then SMR.Close
        strSql = "select a.CUSTLOT,a.WAFERID, convert(int,isnull(a.GOODDIEQTY,0))+ convert(int,isnull(a.NGDIEQTY,0)),b.数量 from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where  rtrim(a.CUSTLOT)='" & strlot & "' and   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. 流程卡编号,'+','')) and b.库房编号='72' and b.合格标记=0 and  convert(int,isnull(a.GOODDIEQTY,0))+ convert(int,isnull(a.NGDIEQTY,0))>b. 数量 "

        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            SMR.MoveFirst
            For j = 1 To SMR.RecordCount
                errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("WAFERID"))
                SMR.MoveNext
            Next
            MsgBox "回货资料有误" & errormsg & ",回货数量超出72仓库存", vbInformation, "提示"
            Exit Function
        End If
    End If
    CheckData = True

End Function

Function boxsplit(Index As Integer) 'index=1 第一次拆，GoodDie；index=2 第二次拆，NGdie

    Dim SMR        As New ADODB.Recordset
    Dim rs        As New ADODB.Recordset
    Dim strSql As String
    Dim strboxid_old As String
    Dim strboxid_new As String
    Dim splitqty As Integer
    Dim extraqty As Integer
    Dim WAFER As String
    Dim i As Integer
    Dim j As Integer
    Dim qnum As String

    '1.取得新箱号
    NewBoxList = ""
    OldBoxList = ""
    boxsplit = True
    If Index = 1 Then
        strSql = "select distinct b.箱号 from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. 流程卡编号,'+','')) and b.库房编号='72' and b.合格标记=0 and isnull(a.GOODDIEQTY,0)<=b. 数量 and b.箱号 not like '%VT%'"
    ElseIf Index = 2 Then
        strSql = "select distinct b.箱号 from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. 流程卡编号,'+','')) and b.库房编号='72' and b.合格标记=0 and isnull(a.NGDIEQTY,0)<=b. 数量 and b.箱号 not like '%VT%'"
    End If
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For j = 1 To SMR.RecordCount
            If OldBoxList = "" Then
                OldBoxList = Trim(SMR("箱号"))
            Else
                OldBoxList = OldBoxList & "," & Trim(SMR("箱号"))
            End If
            strSql = " SELECT COUNT(*) FROM erpdata..tblStockNumTree c WHERE  c.箱号 LIKE '" & Trim(SMR("箱号")) & "' + '%' "
               
            If rs.State = adStateOpen Then rs.Close
            rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            qnum = Trim(Str(rs.Fields(0).Value))
            If qnum = 1 Then
                strboxid_new = Trim(SMR("箱号")) + "_VT"
            Else
                 strboxid_new = Trim(SMR("箱号")) & "_VT" & qnum - 1
            End If
            If NewBoxList = "" Then
                NewBoxList = strboxid_new
            Else
                NewBoxList = NewBoxList & "," & strboxid_new
            End If
            rs.Close
            Set rs = Nothing

            SMR.MoveNext
        Next
    End If
    SMR.Close
    Set SMR = Nothing
    '2.插入新箱号数据
    strboxid_new = ""
    If NewBoxList <> "" Then
        For i = 0 To UBound(Split(NewBoxList, ","))
            strboxid_new = Split(NewBoxList, ",")(i)
            strboxid_old = Split(OldBoxList, ",")(i)
            AddSql2 ("INSERT INTO erpdata..TBLPACKMAININF(箱号,客户代码,数量,产线标记,合格标记,装箱标记)  VALUES ('" & strboxid_new & "','" & UCase(Trim(CboCustomer.text)) & "',1,'1','0','1')")
            AddSql2 ("INSERT INTO erpdata..tblPackTreeInf(箱号) VALUES ('" & strboxid_new & "')")
            AddSql2 ("INSERT INTO erpdata..tblStockNumTree ( 序号,箱号,上级序号,基层标记,发领标记) SELECT b.序号,b.箱号,b.上级序号,b.基层标记,'0' FROM erpdata..tblPackTreeInf b WHERE b.箱号 = '" & strboxid_new & "' ")
            If Index = 1 Then
                 strSql = "select b.流程卡编号,ISNULL(a.GOODDIEQTY,0),b.数量 -ISNULL(a.GOODDIEQTY,0)  from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. 流程卡编号,'+','')) and b.库房编号='72' and b.合格标记=0 and isnull(a.GOODDIEQTY,0) >0 and isnull(a.GOODDIEQTY,0)<=b. 数量 and  rtrim(b.箱号)='" & strboxid_old & "' "
            Else
                 strSql = "select b.流程卡编号,ISNULL(a.NGDIEQTY,0) ,b.数量 -ISNULL(a.NGDIEQTY,0)  from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. 流程卡编号,'+','')) and b.库房编号='72' and b.合格标记=0 and isnull(a.NGDIEQTY,0) >0 and isnull(a.NGDIEQTY,0)<=b. 数量 and  rtrim(b.箱号)='" & strboxid_old & "' "
            End If
            If SMR.State = adStateOpen Then SMR.Close
            SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            If SMR.RecordCount > 0 Then
                SMR.MoveFirst
                For j = 1 To SMR.RecordCount
                    WAFER = Trim(SMR.Fields(0).Value)
                    splitqty = Trim(SMR.Fields(1).Value)
                    extraqty = Trim(SMR.Fields(2).Value)
               
                    If splitqty > 0 And extraqty > 0 Then
                         AddSql2 (" INSERT INTO erpdata..tblStockNumSub  SELECT '" & strboxid_new & "',a.流程卡编号,a.工单号,'" & splitqty & "',a.料号,a.物料编号,a.合格标记,a.发货标记  ,a.ID,a.库房编号,GETDATE(),a.大工单 FROM erpdata..tblStockNumSub a  WHERE a.箱号 = '" & strboxid_old & "' AND a.流程卡编号 = '" & WAFER & "'")
                         AddSql2 (" UPDATE erpdata..tblStockNumSub SET 数量 = 数量 - " & splitqty & " WHERE 箱号 = '" & strboxid_old & "' AND 流程卡编号 = '" & WAFER & "' ")
                    ElseIf splitqty > 0 Then
                         AddSql2 ("  UPDATE erpdata..tblStockNumSub SET 箱号 = '" & strboxid_new & "' WHERE 箱号 = '" & strboxid_old & "' AND 流程卡编号 = '" & WAFER & "' ")
                    End If
                    SMR.MoveNext
                Next
            End If
            SMR.Close
            Set SMR = Nothing
        Next
    End If
    '---------------
    boxsplit = True
End Function


Function CommitApplication(RequestNo As String, Index As Integer, strLotList As String)
Dim SMR        As New ADODB.Recordset
Dim rs        As New ADODB.Recordset
Dim strSql As String
Dim strlot As String

Dim strXH As String
Dim strxh_big As String
Dim strgdh As String
Dim strLCK As String
Dim strlps As String
Dim strbls As String
Dim strzcbls As String
Dim strid As String
Dim strmbkf As String
Dim strKF As String
Dim strmatcode As String
Dim strCustCode As String
Dim i As Integer
Dim j As Integer
Dim intitem As Integer
Dim SumCount As Integer

     For i = 0 To UBound(Split(strLotList, ","))
         strlot = Split(strLotList, ",")(i)

        If UCase(Trim(CboCustomer.text)) = "GC" Then
            strSql = "select b.箱号,b.流程卡编号,b.工单号,b.数量,b.id from erptemp..TSV_VT_History_sub a  ,erpdata..tblstocknumsub b where  a.flag=1 and  a.CUSTLOT='" & strlot & "' and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b.流程卡编号,'+','')) and b.数量>0 and b.合格标记=0 order by a.WAFERID"
        ElseIf (UCase(Trim(CboCustomer.text)) = "KR001" Or UCase(Trim(CboCustomer.text)) = "KR009") Then
            If Index = 1 Then
                strSql = "select b.箱号,b.流程卡编号,b.工单号,b.数量,b.id from erptemp..TSV_VT_History_sub a  ,erpdata..tblstocknumsub b where  a.flag=1 and  a.CUSTLOT='" & strlot & "' and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b.流程卡编号,'+','')) and b.数量>0  and b.数量=a.GOODDIEQTY and b.合格标记=0 order by a.WAFERID"
            ElseIf Index = 2 Then
                strSql = "select b.箱号,b.流程卡编号,b.工单号,b.数量,b.id from erptemp..TSV_VT_History_sub a  ,erpdata..tblstocknumsub b where  a.flag=1 and  a.CUSTLOT='" & strlot & "' and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b.流程卡编号,'+','')) and b.数量>0  and b.数量=a.NGDIEQTY and b.合格标记=0 order by a.WAFERID"
            End If
        End If

        If SMR.State = adStateOpen Then SMR.Close
        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            SMR.MoveFirst
            For j = 1 To SMR.RecordCount
                If j = 1 Then
                    strSql = " select top 1 rtrim(a.原仓库) from erpdata..tblStockdb a, erpdata..tblStockdbsub b where a.id=b.id and  rtrim(b. 流程卡编号)='" & Trim(SMR("流程卡编号")) & "' and a.目标仓库='72'"
                    strmbkf = GetSqlServerStr(strSql)
                    '不良品要调回对应的不良品仓
                    '07,20 仓委外的，保税不良品都回30
                    '16,19仓委外的，非保税不良品都回28
    
                    If (UCase(Trim(CboCustomer.text)) = "KR001" Or UCase(Trim(CboCustomer.text)) = "KR009") And Index = 2 Then
                         Select Case GetSqlServerStr(strSql)
                         Case "07", "20"
                             strmbkf = "30"
                         Case "16", "19"
                             strmbkf = "30"
                         Case Else
    
                         End Select
                         
                    End If
                End If
                If GetNewBoxId(Trim(SMR("箱号"))) <> Trim(SMR("箱号")) Then
                    SMR.MoveNext
                    GoTo 1  '为避免拆箱出来的2个箱号数量一致，会查询出两笔结果
                End If
                
                strXH = Trim(SMR("箱号"))     '箱号
                strxh_big = ""   '大箱号
                strgdh = Trim(SMR("工单号"))       '工单号
                strLCK = Trim(SMR("流程卡编号"))    '流程卡编号
                strlps = Trim(SMR("数量"))        '良品数
                strbls = "0"     '不良品数
                strzcbls = "0"     '制程不良数
                strid = Trim(SMR("id"))
                If Get_SqlserverCnt("select * from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid) > 0 Then
                    strSql = "select ITEM from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid
                    intitem = GetSqlServerStr(strSql)
                Else
                    strSql = "select isnull(max(ITEM),0) from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'"
                    intitem = GetSqlServerStr(strSql) + 1
    
                    If rs.State = adStateOpen Then rs.Close
                    strSql = " select 库房编号,物料编号,客户代码,isnull(单价,0) from erpdata..tblStockNum where id=" & strid
                   
                    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                    If rs.RecordCount = 1 Then
                        rs.MoveFirst
                        strKF = Trim(rs("库房编号"))
                        strmatcode = Trim(rs("物料编号"))
                        strCustCode = Trim(rs("客户代码"))
                    
                    End If
            
                  '上传主表
    
                 '调拨编号,序号,物料编号, 调拨数量,原仓库,目标仓库,申请人员,申请时间,审核人员,审核时间, 申请部门,状态,REMARK1,REMARK2,REMARK3,REMARK4,REMARK5,ID
                    strSql = "insert into erptemp..tblstockdb_temp(ORDER_NUM,ITEM, MATERIALS,QTY,FORMER, DESTINATION, APPLICANT, APPLICATION_TIME, AUDITOR, AUDIT_TIME, DEPT, FLAG,ID,REMARK1) values( " & _
                    "'" & RequestNo & "'," & intitem & ",'" & strmatcode & "'," & 0 & ",'" & strKF & "','" & strmbkf & "','" & gUserName & "','','','','',1," & strid & ",'')"
                
                    AddSql2 (strSql)
                                      
                End If
                            
                '上传子表
                
                '调拨编号, 序号, 箱号, 流程卡编号, 工单号, 合格数, 制程不良数, 来料不良数, ID
                 strSql = "insert into erptemp..tblstockdbsub_temp(ORDER_NUM,ITEM,WAFER,LOT,GOOD_DIE,BAD1_DIE,BAD2_DIE,ID,REMARK1,QBOX) values( " & _
                "'" & RequestNo & "'," & intitem & ",'" & strLCK & "','" & strgdh & "'," & strlps & "," & strbls & "," & strzcbls & "," & strid & ",'" & strxh_big & "','" & strXH & "')"
              
                AddSql2 (strSql)
                
                'update主表数量
                strSql = "Update erptemp..tblstockdb_temp set QTY =QTY+" & Val(strlps) + Val(strbls) + Val(strzcbls) & " where ORDER_NUM='" & RequestNo & "' and ITEM=" & intitem
               
                AddSql2 (strSql)

               
                SumCount = SumCount + 1
                SMR.MoveNext
1:            Next j
       
            'If Index = 2 Then
                strSql = "Update erptemp..TSV_VT_History_sub set flag=2 where CUSTLOT='" & strlot & "' "
                AddSql2 (strSql)
           ' End If
        End If
    Next
    If SumCount > 0 Then
        MsgBox SumCount & "笔记录申请成功", vbInformation, "提示"
        Txt_sqdh.text = RequestNo
        If SMR.State = adStateOpen Then SMR.Close
        strSql = "select *  from erptemp..tblstockdbsub_temp  where  ORDER_NUM='" & RequestNo & "'"
        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            With fpS
                Set .DataSource = SMR
            End With

        End If
     End If
End Function
     
Function GetID()
'FDP1911140011
'生成方式：FWW+YYMMDD +4位流水码
Dim CODE       As String
Dim strSql     As String
Dim YearStr    As String
Dim MonthStr   As String
Dim DayStr     As String
Dim SMR        As New ADODB.Recordset


YearStr = Right(Year(Now()), 2)
If Len(Month(Now())) = 1 Then
    MonthStr = "0" & Month(Now())
Else
    MonthStr = Month(Now())
End If
If Len(Day(Now())) = 1 Then
    DayStr = "0" & Day(Now())
Else
    DayStr = Day(Now())
End If
CODE = YearStr & MonthStr & DayStr

strSql = "Select Isnull(max(RIGHT(ORDER_NUM,LEN(ORDER_NUM)-3)),0) as ORDER_NUM from erptemp..tblStockdb_temp where left(ORDER_NUM,9)='FWW" & CODE & "'"


If SMR.State = adStateOpen Then SMR.Close
SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If SMR("ORDER_NUM") = 0 Then

    GetID = "FWW" & CODE & "0001"
Else
    GetID = "FWW" & Val(SMR("ORDER_NUM")) + 1
End If
SMR.Close
Set SMR = Nothing

End Function

Function GetNewBoxId(strBoxID As String)

    Dim i As Integer
    Dim strboxid_old As String
    Dim strboxid_new As String

    If Trim(OldBoxList) = "" Then
        GetNewBoxId = strBoxID
        Exit Function
    End If
    For i = 0 To UBound(Split(OldBoxList, ","))
      strboxid_old = Split(OldBoxList, ",")(i)
      strboxid_new = Split(NewBoxList, ",")(i)
      If strboxid_old = strBoxID Then
          GetNewBoxId = strboxid_new
      End If
    Next
    If GetNewBoxId = "" Then
        GetNewBoxId = strBoxID
    End If

End Function

Function GetVTID()
'日期6+4码流水，共10码
'2001151530001
Dim CODE       As String
Dim strSql     As String
Dim YearStr    As String
Dim MonthStr   As String
Dim DayStr     As String
Dim HourStr   As String
Dim MinuteStr     As String
Dim SMR        As New ADODB.Recordset

GetVTID = ""
YearStr = Right(Year(Now()), 2)
If Len(Month(Now())) = 1 Then
    MonthStr = "0" & Month(Now())
Else
    MonthStr = Month(Now())
End If
If Len(Day(Now())) = 1 Then
    DayStr = "0" & Day(Now())
Else
    DayStr = Day(Now())
End If


CODE = YearStr & MonthStr & DayStr

strSql = "Select Isnull(max(回货单号),0) as 回货单号 from erptemp..TSV_VT_History_sub where left(回货单号,6)='" & CODE & "'"

If SMR.State = adStateOpen Then SMR.Close
SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If SMR("回货单号") = 0 Then

    GetVTID = CODE & "001"
Else
    GetVTID = Val(SMR("回货单号")) + 1
End If


SMR.Close
Set SMR = Nothing


End Function



Private Sub ExportToExcel_GCWO()
    Dim xlsApp      As Excel.Application
    Dim xlsBook     As Excel.Workbook
    Dim xlsSheet    As Excel.Worksheet
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    Dim strFileName As String
    On Error GoTo Ert


    Set xlsApp = CreateObject("Excel.Application")
    Set xlsBook = xlsApp.Workbooks.Add
    Set xlsSheet = xlsBook.Worksheets(1)

    With xlsApp
        .Rows(1).Font.Bold = True
    End With


strSql = " SELECT distinct 'GCSH' AS 'Sub Name','HTKS' AS 'Ship To', t1.FAB_CONV_ID AS 'FAB Device',t1.CUSTDEVICE AS 'Customer Device' ,LEFT(t1.IMAGER_CUSTOMER_REV,2) + d.二级代码 AS 'GC Version' ," & _
" t1.PO_NUM AS 'PO NO' ,'' AS WO,'' AS 'Invoice NO', convert(nvarchar(20),GETDATE(),111) AS 'FAB-Out DATE',t1.CUSTLOT AS 'FAB Lot ID',t1.WAFERID AS 'Wafer ID'," & _
" t1.PASSBINCOUNT AS 'Gross Dies','' AS 'Sampling Qty','' AS 'Pass Dies' ,'' AS Yield,t1.REMARK1 AS Remark ,t1.备注,t1. 厂内机种 FROM (" & _
" SELECT c.FAB_CONV_ID ,  a.CUSTDEVICE   ,c.IMAGER_CUSTOMER_REV   ,c.PO_NUM ,c.MTRL_NUM, a.CUSTLOT  , a.WAFERID ,b.PASSBINCOUNT, A.REMARK1," & _
" ISNULL(a.REMARK2,'') AS 备注,ISNULL(a.REMARK3,'')  AS  厂内机种" & _
" FROM ERPTEMP..TSV_VT_History_sub a" & _
" LEFT JOIN erpbase..tblmappingdata b ON a.CUSTLOT =b.lotid  AND a.WAFERID =right(100+b.WAFER_ID,2)  AND a.customershortname=b.customershortname  " & _
" LEFT JOIN erpbase..tblCustomerOI c  ON convert(VARCHAR(50),c.id)=b.filename  AND b.lotid=c.SOURCE_BATCH_ID  AND a.CUSTDEVICE +'-3'=c.MPN_DESC  " & _
" WHERE  a.FLAG_WO =1 AND a.customershortname='GC' ) t1" & _
" LEFT JOIN ERPTEMP..GcCode_Reference d ON  t1.CUSTDEVICE  +'-3'=d.客户机种名 AND d.制程=t1.备注" & _
" ORDER BY t1.CUSTLOT ,t1.WAFERID "

 
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If SMR.RecordCount > 0 Then

        xlsSheet.Cells(1, 1) = "No."
        xlsSheet.Cells(1, 2) = "Sub Name"
        xlsSheet.Cells(1, 3) = "Ship To"
        xlsSheet.Cells(1, 4) = "FAB Device"
        xlsSheet.Cells(1, 5) = "Customer Device"
        xlsSheet.Cells(1, 6) = "GC Version"
        xlsSheet.Cells(1, 7) = "PO NO"
        xlsSheet.Cells(1, 8) = "WO"
        xlsSheet.Cells(1, 9) = "Invoice NO"
        xlsSheet.Cells(1, 10) = "FAB-Out Date"
        xlsSheet.Cells(1, 11) = "FAB Lot ID"
        xlsSheet.Cells(1, 12) = "Wafer ID"
        xlsSheet.Cells(1, 13) = "Gross Dies"
        xlsSheet.Cells(1, 14) = "Sampling Qty"
        xlsSheet.Cells(1, 15) = "Pass Dies"
        xlsSheet.Cells(1, 16) = "Yield"
        xlsSheet.Cells(1, 17) = "Remark"
        xlsSheet.Cells(1, 18) = "备注"
        xlsSheet.Cells(1, 19) = "厂内机种"
       SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            xlsSheet.Cells(i + 1, 1) = i
            xlsSheet.Cells(i + 1, 2) = Trim(SMR("Sub Name"))
            xlsSheet.Cells(i + 1, 3) = Trim(SMR("Ship To"))
            xlsSheet.Cells(i + 1, 4) = Trim(SMR("FAB Device"))
            xlsSheet.Cells(i + 1, 5) = Trim(SMR("Customer Device"))
            xlsSheet.Cells(i + 1, 6) = Trim(SMR("GC Version"))
            xlsSheet.Cells(i + 1, 7) = Trim(SMR("PO NO"))
            xlsSheet.Cells(i + 1, 8) = ""
            xlsSheet.Cells(i + 1, 9) = Trim(SMR("Invoice NO"))
            xlsSheet.Cells(i + 1, 10) = Trim(SMR("FAB-Out Date"))
            xlsSheet.Cells(i + 1, 11) = Trim(SMR("FAB Lot ID"))
            xlsSheet.Cells(i + 1, 12) = Trim(SMR("Wafer ID"))
            xlsSheet.Cells(i + 1, 13) = Trim(SMR("Gross Dies"))
            xlsSheet.Cells(i + 1, 14) = Trim(SMR("Sampling Qty"))
            xlsSheet.Cells(i + 1, 15) = Trim(SMR("Pass Dies"))
            xlsSheet.Cells(i + 1, 16) = Trim(SMR("Yield"))
            xlsSheet.Cells(i + 1, 17) = Trim(SMR("Remark"))
            xlsSheet.Cells(i + 1, 18) = Trim(SMR("备注"))
            xlsSheet.Cells(i + 1, 19) = Trim(SMR("厂内机种"))
            
            
            SMR.MoveNext
        Next
        With xlsSheet.Range("2:" & i)
            .horizontalAlignment = xlLeft
        End With
        xlsSheet.Range("A1").Select
        xlsApp.Columns.AutoFit
    
    End If
    SMR.Close
    Set SMR = Nothing
    
    xlsApp.Visible = True
    filepath_org = Trim(Text3.text)
    strFileName = Left(filepath_org, InStrRev(filepath_org, ".") - 1) & "_WO" & Format(Now, "YYYYMMDDhhmmss") & Mid(filepath_org, InStrRev(filepath_org, "."), Len(filepath_org) - InStrRev(filepath_org, ".") + 1)
    xlsBook.SaveAs strFileName

    Set xlsApp = Nothing
    Set xlsSheet = Nothing
    Set xlsBook = Nothing
    strSql = "Update erptemp..TSV_VT_History_sub set FLAG_WO=2 where FLAG_WO=1"
    AddSql2 (strSql)
    MsgBox "WO转换完成", vbInformation, "提示"
    
               
    
    
Ert:

    If Not (xlsApp Is Nothing) Then
        
        Set xlsApp = Nothing
        Set xlsSheet = Nothing
        Set xlsBook = Nothing

    End If
    

End Sub


Private Sub UploadVTWO_GC()

'GC回货WO上传

End Sub

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
'    wafer_to_string = b(0) + "," + b(1)
End If

'Last = Last - 2
Last = Last - 1


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
If b(Last) = b(Last - 1) + 1 Then
    Last = Last + 1
    String2 = String2 + "-" + b(Last)
    wafer_to_string = String2
Else
    wafer_to_string = String2
End If
End Function

Private Function GetGcrevFromWO(strlot As String, strWaferID As String)
'获取WO中的二级代码前2码
 Dim strSql As String
 strSql = "SELECT distinct left(b.IMAGER_CUSTOMER_REV,2) FROM erpbase..tblmappingData a inner join ERPBASE..TBLCUSTOMEROI b ON  convert(VARCHAR(30),b.ID)=a.FILENAME AND b.SOURCE_BATCH_ID=a.LOTID and a.lotid='" & strlot & "' and convert(int,wafer_id)=" & CInt(strWaferID)
 GetGcrevFromWO = GetSqlServerStr(strSql)

End Function


Private Function GetHTDevice(strCustDevice As String, strtype As String, strGcrev2 As String)
'取得厂内机种
    If Get_SqlserverCnt("select distinct 厂内机种名 from ERPTEMP..GcCode_Reference where 客户机种名='" & strCustDevice & "-3' and 制程='" & strtype & "'") <> 1 Then
    
        If Get_SqlserverCnt("select distinct 厂内机种名 from ERPTEMP..GcCode_Reference where 客户机种名='" & strCustDevice & "-3' and 制程='" & strtype & "' and 二级代码第二位 ='" & strGcrev2 & "'") <> 1 Then
            GetHTDevice = ""
        Else
            GetHTDevice = GetSqlServerStr("select distinct 厂内机种名 from ERPTEMP..GcCode_Reference where 客户机种名='" & strCustDevice & "-3' and 制程='" & strtype & "' and 二级代码第二位 ='" & strGcrev2 & "'")
        End If
    Else
        GetHTDevice = GetSqlServerStr("select distinct 厂内机种名 from ERPTEMP..GcCode_Reference where 客户机种名='" & strCustDevice & "-3' and 制程='" & strtype & "'")
    End If
End Function






