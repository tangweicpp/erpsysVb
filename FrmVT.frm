VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmVT 
   Caption         =   "VT���������ϴ�"
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
      TabCaption(0)   =   "�ػ������ϴ�"
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
      TabCaption(1)   =   "�ػ�����ɾ��"
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
         Caption         =   "��ѯ"
         Height          =   375
         Left            =   -72240
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Cmd_del 
         Caption         =   "ɾ��"
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
         Caption         =   "ѡ����ϴ����ļ�"
         Height          =   2775
         Left            =   360
         TabIndex        =   3
         Top             =   1020
         Width           =   14655
         Begin VB.CommandButton Cmd_GCNewformat 
            Caption         =   "GC�ػ��¸�ʽ�ϴ�"
            Height          =   495
            Left            =   7680
            TabIndex        =   20
            Top             =   1560
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton Cmd_exportWO 
            Caption         =   "ת��WO"
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
            Caption         =   "�ػ�����"
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
            Caption         =   "�ϴ�DB"
            Height          =   480
            Left            =   480
            TabIndex        =   5
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command8 
            Caption         =   "��������"
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
            Caption         =   "���뵥��"
            Height          =   255
            Left            =   600
            TabIndex        =   12
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ѡ����ϴ���xlsx��"
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
         Caption         =   "����ɾ����δ���ɻػ����������"
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
         Caption         =   "�ͻ���"
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


Private Type GCCUSTDATA  '�ͻ��ṩ�Ļػ���Ϣ

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
        MsgBox "��ѡ��Դ�ļ�����·��", vbInformation, "��ʾ"
        Exit Function
    End If
    Set VBExcel_Source = CreateObject("excel.application")
    VBExcel_Source.Visible = False
    Set xlBook_Source = VBExcel_Source.Workbooks.Open(Text3.text)
    Set xlSheet_Source = xlBook_Source.Worksheets(1)

   
    lColsCnt = xlSheet_Source.Range("A1").CurrentRegion.Columns.count
    lRowsCnt = xlSheet_Source.Range("A1").CurrentRegion.Rows.count

    'If InStr(Trim(xlSheet_Source.Range("A1").Value), "����") = 0 Then
    '    MsgBox "��ѡ��ͻ����ṩ�ĸ�ʽ�ϴ���", vbInformation, "��ʾ"
   '     GoTo EXITPRO
    '    Exit Function
  '  End If


    
    j = 0
    

    cmdStr = "delete from erptemp..GcExcelTranslate"
    AddSql2 (cmdStr)

If VTformat = "new" Then
    If lColsCnt <> 14 Then
        MsgBox "Excel�е�����:" & lColsCnt & "���趨��ģ������:13��һ��" & vbCrLf & "��ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        GoTo EXITPRO
        Exit Function

    End If
    
    startrow = 2
    If InStr(Trim(xlSheet_Source.Range("B1").Value), "��Ʒ�ͺ�") = 0 Then
        MsgBox "��ģ��A1��Ԫ����ӦΪ��Ʒ�ͺ��ĸ���,��ȷ�ϸ�ʽ", vbInformation, "��ʾ"
        Exit Function
    End If
ElseIf VTformat = "old" Then
    If lColsCnt <> 13 Then
        MsgBox "Excel�е�����:" & lColsCnt & "���趨��ģ������:13��һ��" & vbCrLf & "��ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        GoTo EXITPRO
        Exit Function

    End If
    startrow = 4
    
    If InStr(Trim(xlSheet_Source.Range("A2").Value), "����") = 0 Then
        MsgBox "��ģ��A2��Ԫ����ӦΪ���ڶ���,��ȷ�ϸ�ʽ", vbInformation, "��ʾ"
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
            MsgBox "D��waferid��ʽ����ȷ", vbInformation, "��ʾ"
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
            MsgBox "D��waferid��ʽ����ȷ", vbInformation, "��ʾ"
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
    '�ϲ���Ԫ�񣬻�ȡ���Ͻǵ�Ԫ���value
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

'��ӵ���Sqlserver

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
ElseIf UCase(Replace(Trim(dT.remarkTemp), " ", "")) = "��MAIN" Then   'תNormal
    strtype = "תNormal"
Else
    strtype = ""
End If

strgcrev_2 = GetGcrevFromWO(LOTID, WAFER)
If Len(strgcrev_2) = 2 Then
    strHtDevice = GetHTDevice(dT.CustDeviceTemp, strtype, Right(strgcrev_2, 1))
End If
cmdStr = "insert into erptemp..GcExcelTranslate(����,����,C_NO,����_CST,WaferID,�ͺ�,����,��������,�ȼ�,Ƭ��,����,��ⱸע,��װ�ߴ�,LotID,Wafer,id,remark1 ,remark2, remark3 ) values  " & _
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

    strSql = "SELECT DISTINCT a.����, a.����_CST,a.�ͺ�,a.lotid,WaferId = (STUFF((SELECT ',' +  Wafer FROM erptemp..gcexceltranslate WHERE a.LotID=lotid and a.����_CST=����_CST  AND a.��ⱸע=��ⱸע order by Wafer FOR XML PATH('')), 1,  1, '')),sum(convert(INT,(a.Ƭ��))) as Ƭ��,'����' as Factory,a.��ⱸע,a.remark1 as ��ʽ, a.remark3 as ���ڻ��� FROM  erptemp..gcexceltranslate  a GROUP BY a.����, a.����_CST,a.�ͺ�,a.lotid,a.��ⱸע,a.remark1 ,a.remark3 "


    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If SMR.RecordCount > 0 Then
    
        xlsSheet.Cells(1, 1) = "����"
        xlsSheet.Cells(1, 2) = "����_CST"
        xlsSheet.Cells(1, 3) = "�ͺ�"
        xlsSheet.Cells(1, 4) = "lotid"
        xlsSheet.Cells(1, 5) = "WaferId"
        xlsSheet.Cells(1, 6) = "Ƭ��"
        xlsSheet.Cells(1, 7) = "Factory"
        xlsSheet.Cells(1, 8) = "��ⱸע"
        xlsSheet.Cells(1, 9) = "���ڻ���"
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            xlsSheet.Cells(i + 1, 1) = Trim(SMR("����"))
            xlsSheet.Cells(i + 1, 2) = Trim(SMR("����_CST"))
            xlsSheet.Cells(i + 1, 3) = Trim(SMR("�ͺ�"))
            xlsSheet.Cells(i + 1, 4) = Trim(SMR("lotid"))
            xlsSheet.Cells(i + 1, 5) = Trim(SMR("WaferId"))
            xlsSheet.Cells(i + 1, 6) = Trim(SMR("Ƭ��"))
            xlsSheet.Cells(i + 1, 7) = Trim(SMR("Factory"))
            xlsSheet.Cells(i + 1, 8) = Trim(SMR("��ⱸע"))
            xlsSheet.Cells(i + 1, 9) = Trim(SMR("���ڻ���"))
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

    MsgBox "ת�����", vbInformation, "��ʾ"
Ert:
    MsgBox Err.DESCRIPTION & vbCrLf & "in ��ʽ����1..ExportToExcel ", vbExclamation + vbOKOnly, "Application Error"
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
        If MsgBox("��ȷ��Ҫɾ��" & Delmsg & ",��" & DelCnt & "�ʻػ�������?", vbOKCancel, "��ʾ") = vbCancel Then
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
   ' cmd_search_Click '��ѯ

End Sub

Private Sub Cmd_exportWO_Click()
ExportToExcel_GCWO
End Sub






Private Sub Cmd_GCNewformat_Click()
    Dim strCust As String
    
    VTformat = "new"
    
    If CboCustomer.text = "" Then
        MsgBox "����ѡ��ͻ�����"
        Exit Sub
    
    End If
    'δ���ػ����룬�����ϴ��µĻػ�����
    If Get_SqlserverCnt("select * from erptemp..TSV_VT_History_sub where FLAG=1 AND  CUSTOMERSHORTNAME='" & Trim(CboCustomer.text) & "'") > 0 Then
        MsgBox "�лػ�����δ���ػ����룬��������������ϴ���", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    shipid = CStr(GetVTID)
    If CStr(shipid) = "" Then
        MsgBox "��ȡ�ػ������쳣,������", vbInformation, "��ʾ"
        Exit Sub
    End If
    If UCase(Trim(CboCustomer.text)) = "GC" Then
        strCust = UCase(Trim(CboCustomer.text))
        UploadVTData_GC_New (strCust)
    Else
        MsgBox "�˹��ܽ�GCʹ��", vbInformation, "��ʾ"
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
                
                .SetText 1, 0, "ѡ��"
                .CellType = CellTypeCheckBox
                .text = 0
                .TypeHAlign = TypeHAlignCenter
                .TypeVAlign = TypeVAlignCenter
                .Lock = True
                
                .ReDraw = True
    
            Next

    End With
        
        
    Else
    
        MsgBox "δ�鵽��lot���ϴ���¼", vbInformation, "��ʾ"
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
    '˧ѡ�ļ�
    CommonDialog2.Filter = "EXCEL�ļ�(*.xlsx)|*.xlsx|EXCEL�ļ�(*.xls)|*.xls"
    
    CommonDialog2.ShowOpen
    '�õ��ļ���
    FName = CommonDialog2.filename
    If FName <> "" Then
       Text3.text = FName
    End If


End Sub

Private Sub Command7_Click()
Dim strCust As String
VTformat = "new"
If CboCustomer.text = "" Then
    MsgBox "����ѡ��ͻ�����"
    Exit Sub

End If
'δ���ػ����룬�����ϴ��µĻػ�����
If Get_SqlserverCnt("select * from erptemp..TSV_VT_History_sub where FLAG=1 AND  CUSTOMERSHORTNAME='" & Trim(CboCustomer.text) & "'") > 0 Then
    MsgBox "�лػ�����δ���ػ����룬��������������ϴ���", vbInformation, "��ʾ"
    Exit Sub
End If

'δ��WO�������ϴ��µĻػ�����
' If Get_SqlserverCnt("select * from erptemp..TSV_VT_History_sub where FLAG_WO=1 AND  CUSTOMERSHORTNAME='" & Trim(CboCustomer.Text) & "'") > 0 Then
    ' MsgBox "�лػ�����δ���ػ����룬��������������ϴ���", vbInformation, "��ʾ"
    ' End Sub
' End If


'shipid = Get_OracleStr("select TSV_VT_SEQ.Nextval from dual")
shipid = CStr(GetVTID)
If CStr(shipid) = "" Then
    MsgBox "��ȡ�ػ������쳣,������", vbInformation, "��ʾ"
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

'�ϴ�����

Dim source_batch_id_Temp As String
'�ϴ�OI��CSV
'�����ļ���
If Text3.text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim filename As String

'��ȡ�ļ���
'    If InStrRev(Trim(Text2.Text), "\") > 0 Then
'        strFileName = Mid(Trim(Text2.Text), InStrRev(Trim(Text2.Text), "\") + 1)
'        dirName = Mid$(Trim(Text2.Text), 1, InStrRev(Trim(Text2.Text), "\"))
'    End If
    

'2012-06-27 jiayunzhang �޸Ķ�Excel�ķ�ʽ


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '���ļ�

    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�

    '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 11 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
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
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
        
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

  

    '�ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
    If (JudgeFlagVTData(vtDataTemp.DeliveryNoTemp, vtDataTemp.CUSTLOTTemp)) Then
       MsgBox "����Ѵ��ڣ������ϴ�!", vbInformation, "������ʾ"
       GoTo NextRecord2

    End If


    Call AddVTCustomer(vtDataTemp, customerTemp)
    SumCount = SumCount + 1

    '�ϴ���DB
NextRecord2:

Next i


     VBExcel.Application.DisplayAlerts = False '�ر��ĵ���������ʾ��
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit




If SumCount > 0 Then
    MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�", , "��������"
    
    Else
        If BCResultFlag = True Then
            MsgBox "�ϴ�ʧ�ܣ���ȷ�����ϸ�ʽ��", , "��������"
            Exit Sub
        End If
    
End If





''��ȡCSV
'Dim source_batch_id_Temp As String
'Dim customerTemp As String
'
'customerTemp = "GC"
'
''�ϴ�OI��CSV
''�����ļ���
'If Text3.Text = "" Then
'    MsgBox "��ѡ����ϴ����ļ�"
'    Exit Sub
'End If
'Dim dirName As String
'Dim FileName As String
'
''��ȡ�ļ���
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
''                '�ж�wo�Ƿ�Ϊ��
''
''                If Trim(gcHeaderTemp.WO_NO) = "" Then
''
''                    MsgBox "WO_NO�п�ֵ����ȷ�ϣ�"
''                    Exit Sub
''
''                End If
''
''                '2012-11-07 jiayun �޸�Good_Die_Qty �����г�������
''
''            gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
''
''            '2013-12-27 jiayun add
''
''            If gcDetailTemp.Good_Die_Qty <= 0 Then
''                    MsgBox "��ȷ�Ͽͻ����ֶ�Ӧ��Die���Ƿ���ά���ã�"
''                    Exit Sub
''            End If
''
''
''            '2012-11-05 jiayun �޸� GC
''
''            '�ж�lotID��Header�����Ƿ��Ѵ���
''
''            If (JudgeGCHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)) Then
''
''                If GCHeaderFlag = False Then
''        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
''                End If
''
''                '2013-12-05 jiayun add ���lotid,wo_no �Ѵ��ڣ����ѯ��id
''                '��lotid�и���ʱ�����ѯ�ϴε�id
''
''                id = GetGCLotIDWOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)
''
''            Else
''            '�ϴ���Header����
''                'ȡĿǰDB����ID��
''                id = GetMaxID()
''                '2013-01-11 jiayun add �ͻ����
''
''                If id = 0 Then
''                    MsgBox "DB����ID����ʧ��1������ϵ��Ѷ��"
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
''            '�ж�lotID��Detail�����Ƿ��Ѵ���
''
''            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_ID)) Then
''               MsgBox "GC ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_ID & "�Ѵ��ڣ������ϴ�!"
''
''            Else
''            '�ϴ���Detail����
''
''                   '2012-11-05 jiayun �޸� GCT
''
''
''                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
''
''
''                If id = 0 Then
''                    MsgBox "DB����ID����ʧ��2������ϵ��Ѷ��"
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
'            MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
'        End If

End Sub



Private Sub UploadVTData_KR009(customerTemp As String)
'�ϴ�����
Dim source_batch_id_Temp As String
'�ϴ�OI��CSV
'�����ļ���
If Text3.text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim filename As String

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '���ļ�

    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�

    '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 18 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
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
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
        
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

    '�ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
    If (JudgeFlagVTData_ALL(vtDataTemp)) Then
       MsgBox "����Ѵ��ڣ������ϴ�!", vbInformation, "������ʾ"
       GoTo NextRecord2

    End If

'    If (JudgeFlagVTData(vtDataTemp.DeliveryNoTemp, vtDataTemp.CUSTLOTTemp)) Then
'       MsgBox "����Ѵ��ڣ������ϴ�!", vbInformation, "������ʾ"
'       GoTo NextRecord2
'    End If


    Call AddVTCustomer_KR009(vtDataTemp, customerTemp)
    SumCount = SumCount + 1

    '�ϴ���DB
NextRecord2:

Next i


     VBExcel.Application.DisplayAlerts = False '�ر��ĵ���������ʾ��
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit




If SumCount > 0 Then
    MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�", , "��������"
    
    Else
        If BCResultFlag = True Then
            MsgBox "�ϴ�ʧ�ܣ���ȷ�����ϸ�ʽ��", , "��������"
            Exit Sub
        End If
    
End If

End Sub


Private Sub UploadVTData_GC(customerTemp As String)
'�ϴ�����
Dim source_batch_id_Temp As String
'�ϴ�OI��CSV
'�����ļ���
If Text3.text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim filename As String

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text3.text)    '���ļ�

    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�

    '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 7 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
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
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
        
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
       MsgBox "����Ѵ��ڣ������ϴ�!", vbInformation, "������ʾ"
       GoTo NextRecord2

    End If

    Call AddVTCustomer_GC(vtDataTemp, customerTemp)
    SumCount = SumCount + 1

    '�ϴ���DB
NextRecord2:

Next i
     xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

If SumCount > 0 Then
    MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�", , "��������"
    
    Else
        If BCResultFlag = True Then
            MsgBox "�ϴ�ʧ�ܣ���ȷ�����ϸ�ʽ��", , "��������"
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
    '20200306merry�˶Զ�������ǰ��λ��WLA�Ƿ�һ��
    
    strSql = "SELECT DISTINCT a.lotid  FROM erptemp..gcexceltranslate a,erpbase..tblCustomerOI b Where a.LOTID = b.SOURCE_BATCH_ID And Left(a.��������, 2) <> Left(b.IMAGER_CUSTOMER_REV, 2)"
    Set SMR = Get_SqlserveRs(strSql)
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            errormsg = errormsg & "," & SMR("lotid")
            SMR.MoveNext
        Next
      '  MsgBox "�ػ���������" & errormsg & "����������WLA��һ��", vbInformation, "��ʾ"
    '    Exit Sub
    End If
    
    

    SumCount = 0
    vtDataTemp.Created_ByTemp = gUserName
    
    errormsg = ""
    '����1���Ѿ��ϴ���
    strSql = "select rtrim(lotid)+rtrim(wafer) as waferid from  erptemp..gcexceltranslate  where  rtrim(lotid)+rtrim(wafer) in (select rtrim(CUSTLOT)+rtrim(WAFERID) from  erptemp..TSV_VT_History_sub where (flag=1 or flag=2))  "
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            errormsg = errormsg & "," & SMR("waferid")
            SMR.MoveNext
        Next
        MsgBox "�ػ���������" & errormsg & "���ϴ���", vbInformation, "��ʾ"
        Exit Sub
    End If
    '����2��remark1��λֻ���������������Ϊ�գ���Ϊ��main,����WLT, ��mainΪתNormal
    If Get_SqlserverCnt("select  REMARK1 from  ERPTEMP..TSV_VT_History_sub  WHERE  FLAG_WO =1 AND Customershortname='GC' And isnull(REMARK1,'')<>'' and replace(isnull(REMARK1,''),' ','')<>'��MAIN'") > 0 Then

        MsgBox "��ⱸע��λֻ����'��Main'����������ػ�����", vbInformation, "��ʾ"
        Exit Sub
    End If
    '����3����������δά��
    strSql = " select DISTINCT a.CUSTDEVICE  from  ERPTEMP..TSV_VT_History_sub a WHERE a.FLAG_WO =1 AND a.Customershortname='GC' AND  a.CUSTDEVICE  +'-3' NOT IN (SELECT b.�ͻ������� FROM erptemp..GcCode_Reference  b where  b.�Ƴ�='WLT')"
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            errormsg = errormsg & "," & SMR("CUSTDEVICE")
            SMR.MoveNext
        Next
        MsgBox errormsg & "δά��WLT��������", vbInformation, "��ʾ"
        Exit Sub
    End If

    strSql = " select DISTINCT a.CUSTDEVICE  from  ERPTEMP..TSV_VT_History_sub a WHERE a.FLAG_WO =1 AND a.Customershortname='GC' AND  a.CUSTDEVICE  +'-3' NOT IN (SELECT b.�ͻ������� FROM erptemp..GcCode_Reference  b where  b.�Ƴ�='תnormal')"
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            errormsg = errormsg & "," & SMR("CUSTDEVICE")
            SMR.MoveNext
        Next
        MsgBox errormsg & "δά��תnormal��������", vbInformation, "��ʾ"
        Exit Sub
    End If

    '��ʼ�ϴ�
    strSql = "SELECT DISTINCT a.����, a.����_CST,a.�ͺ�,a.lotid,WaferId = (STUFF((SELECT ',' +  Wafer FROM erptemp..gcexceltranslate WHERE a.LotID=lotid    and a.����_CST=����_CST AND a.��ⱸע=��ⱸע order by Wafer FOR XML PATH('')), 1,  1, '')),sum(convert(INT,(a.Ƭ��))) as Ƭ��,'����' as Factory,a.��ⱸע,a.remark1 as ��ʽ,a.remark3 as ���ڻ��� FROM  erptemp..gcexceltranslate  a GROUP BY a.����, a.����_CST,a.�ͺ�,a.lotid,a.��ⱸע,a.remark1,a.remark3"
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            vtDataTemp.SHIPDATETemp = Trim(SMR("����"))
            vtDataTemp.StockNoTemp = Trim(SMR("����_CST"))
            vtDataTemp.DeliveryNoTemp = Trim(SMR("�ͺ�"))
            vtDataTemp.CUSTLOTTemp = Trim(SMR("lotid"))
            vtDataTemp.BatchTemp = Trim(SMR("WaferId"))
            vtDataTemp.TTLTemp = Trim(SMR("Ƭ��"))
            vtDataTemp.remarkTemp = Trim(SMR("��ⱸע"))
            vtDataTemp.type = Trim(SMR("��ʽ"))
            vtDataTemp.htdevice = Trim(SMR("���ڻ���"))
            
            If (JudgeFlagVTData_GC(vtDataTemp)) Then
               MsgBox "����Ѵ��ڣ������ϴ�!", vbInformation, "������ʾ"
            Else
                Call AddVTCustomer_GC(vtDataTemp, customerTemp)
                SumCount = SumCount + 1
            End If
            SMR.MoveNext
        Next
    End If
    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�", , "��������"
        ExportToExcel
       ' ExportToExcel_GCWO
    Else
        If BCResultFlag = True Then
            MsgBox "�ϴ�ʧ�ܣ���ȷ�����ϸ�ʽ��", , "��������"
            Exit Sub
        End If
        
    End If
    SMR.Close
    Set SMR = Nothing
    


End Sub





Private Sub Command8_Click()
Dim customerStr As String

If Trim(CboCustomer.text) = "" Then
    MsgBox "����ѡ��ͻ����ٵ�������", vbInformation, "������ʾ"
    Exit Sub
ElseIf UCase(Trim(CboCustomer.text)) = "KR009" Then
    customerStr = UCase(Trim(CboCustomer.text))
    ExporToExcel ("  select  SHIPDATE,StockNo,DELIVERYNO,CUSTDEVICE,CUSTLOT,waferId,GOODDIEQTY," & _
         " NGDIEQTY,PackingLOTNo,TTL,WaferQtyIn,Batch,SAPCode,WorkWeek,CartonNo,NETWEIGHT,GROSSWEIGHT,REMARK,�ػ����� " & _
                " From TSV_VT_History where customershortname='" & customerStr & "' order by SHIPDATE  ")
ElseIf UCase(Trim(CboCustomer.text)) = "GC" Then
    customerStr = UCase(Trim(CboCustomer.text))
    ExporToExcel ("  select SHIPDATE as ""C/Not"" ,StockNo as ""���"",DELIVERYNO as ""�ͺ�"",CUSTLOT as ""LOT-ID"",Batch as ""wafer-Id"",TTL as ""����"",REMARK as ""��Ӧ��"" ,�ػ�����" & _
                " From TSV_VT_History where customershortname='" & customerStr & "' order by id  ")
'     ExporToExcel ("  select SHIPDATE ,StockNo ,DELIVERYNO ,CUSTLOT ,Batch ,TTL ,REMARK " & _
'               " From TSV_VT_History where customershortname='" & customerStr & "' order by id  ")
Else

customerStr = UCase(Trim(CboCustomer.text))

ExporToExcel ("  select id, SHIPDATE,DELIVERYNO,CUSTDEVICE,CUSTLOT,GOODDIEQTY,NGDIEQTY,TTL,NETWEIGHT,GROSSWEIGHT,REMARK,�ػ�����" & _
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
'��ӵ���Sqlserver
On Error GoTo DealError
strid = Get_OracleStr("select tbl_tsv_VTData_seq.Nextval from dual")

Cnn.BeginTrans

cmdStr = "insert into  TSV_VT_History (" & _
" SHIPDATE,DELIVERYNO,CUSTDEVICE,CUSTLOT,GOODDIEQTY," & _
" NGDIEQTY,TTL,NETWEIGHT,GROSSWEIGHT,REMARK," & _
" Flag , Created_By, created_date,id,customershortname,�ػ�����) values  " & _
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
       strSql = " INSERT INTO erptemp..TSV_VT_History_sub(SHIPDATE,StockNo,CUSTOMERSHORTNAME,CUSTDEVICE,CUSTLOT,WAFERID,GOODDIEQTY,NGDIEQTY,FLAG,CREATED_BY,CREATED_DATE,ID,�ػ�����,FLAG_WO)" & _
            " Values ('" & TEMP.SHIPDATETemp & "','" & TEMP.DeliveryNoTemp & "','" & customerTemp & "','" & TEMP.CustDeviceTemp & "','" & strlot & "','" & strWafer & "' " & _
            ",'" & TEMP.goodDieQtyTemp & "','" & TEMP.ngDieQtyTemp & "',1,'" & TEMP.Created_ByTemp & "',sysdatetime() , " & strid & ",'" & shipid & "',1)"

        AddSql2 (strSql)
    End If
 

 
 
Cnn.CommitTrans

Exit Sub
DealError:
MsgBox TEMP.CUSTLOTTemp & "δ�ɹ��ϴ�����ȷ��", vbInformation, "��ʾ"
Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True


End Sub

Private Sub AddVTCustomer_KR009(TEMP As VTData, customerTemp As String)
Dim cmdStr As String
Dim cmdStr2 As String
Dim strid As String

'��ӵ���Sqlserver
On Error GoTo DealError
strid = Get_OracleStr("select tbl_tsv_VTData_seq.Nextval from dual")
    
Cnn.BeginTrans

cmdStr = "insert into  TSV_VT_History (" & _
" SHIPDATE,StockNo,DELIVERYNO,CUSTDEVICE,CUSTLOT,waferId,GOODDIEQTY," & _
" NGDIEQTY,PackingLOTNo,TTL,WaferQtyIn,Batch,SAPCode,WorkWeek,CartonNo,NETWEIGHT,GROSSWEIGHT,REMARK," & _
" Flag , Created_By, created_date,id,customershortname,�ػ�����) values  " & _
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
       strSql = " INSERT INTO erptemp..TSV_VT_History_sub(SHIPDATE,StockNo,CUSTOMERSHORTNAME,CUSTDEVICE,CUSTLOT,WAFERID,GOODDIEQTY,NGDIEQTY,FLAG,CREATED_BY,CREATED_DATE,ID,�ػ�����,FLAG_WO)" & _
            " Values ('" & TEMP.SHIPDATETemp & "','" & TEMP.StockNoTemp & "','" & customerTemp & "','" & TEMP.CustDeviceTemp & "','" & TEMP.CUSTLOTTemp & "','" & TEMP.waferIdTemp & "' " & _
            ",'" & TEMP.goodDieQtyTemp & "','" & TEMP.ngDieQtyTemp & "',1" & _
             ",'" & TEMP.Created_ByTemp & "',sysdatetime() , " & strid & ",'" & shipid & "',1)"
 
    
        AddSql2 (strSql)
    End If
 
Cnn.CommitTrans

Exit Sub
DealError:
MsgBox TEMP.CUSTLOTTemp & "δ�ɹ��ϴ�����ȷ��", vbInformation, "��ʾ"
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

'��ӵ���Sqlserver
On Error GoTo DealError
strid = Get_OracleStr("select tbl_tsv_VTData_seq.Nextval from dual")
Cnn.BeginTrans

cmdStr = "insert into  TSV_VT_History (" & _
" SHIPDATE,StockNo,DELIVERYNO,CUSTLOT,Batch,TTL,REMARK," & _
" Flag,Created_By, created_date,id,customershortname,�ػ�����) values  " & _
" ('" & TEMP.SHIPDATETemp & "','" & TEMP.StockNoTemp & "','" & TEMP.DeliveryNoTemp & "','" & TEMP.CUSTLOTTemp & "','" & _
"" & TEMP.BatchTemp & "','" & TEMP.TTLTemp & "','" & TEMP.remarkTemp & "'," & _
" 'Y','" & TEMP.Created_ByTemp & "',sysdate, " & strid & ",'" & customerTemp & "','" & shipid & "')"

AddSql (cmdStr)

 For i = 0 To UBound(Split(TEMP.BatchTemp, ","))
    strWafer = Split(TEMP.BatchTemp, ",")(i)
    strSql = " INSERT INTO erptemp..TSV_VT_History_sub(SHIPDATE,StockNo,CUSTOMERSHORTNAME,CUSTDEVICE,CUSTLOT,WAFERID,FLAG,CREATED_BY,CREATED_DATE,ID,�ػ�����,FLAG_WO,REMARK1,REMARK2,REMARK3)" & _
            " Values ('" & TEMP.SHIPDATETemp & "','" & TEMP.StockNoTemp & "','" & customerTemp & "','" & TEMP.DeliveryNoTemp & "','" & TEMP.CUSTLOTTemp & "','" & strWafer & "',1" & _
             ",'" & TEMP.Created_ByTemp & "',sysdatetime() , " & strid & ",'" & shipid & "',1,'" & TEMP.remarkTemp & "','" & TEMP.type & "','" & TEMP.htdevice & "')"
    AddSql2 (strSql)

 Next
Cnn.CommitTrans


Exit Sub
DealError:
MsgBox TEMP.CUSTLOTTemp & "δ�ɹ��ϴ�����ȷ��", vbInformation, "��ʾ"
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
        MsgBox "û����Ҫ����Ļػ����ϣ���ȷ��", vbInformation, "��ʾ"
        Exit Function
    End If
    
    
    For i = 0 To UBound(Split(strLotList, ","))
         strlot = Split(strLotList, ",")(i)
        '���ύ������
        If SMR.State = adStateOpen Then SMR.Close
        strSql = "select a.CUSTLOT,a.WAFERID from erptemp..TSV_VT_History_sub a where rtrim(a.CUSTLOT)='" & strlot & "' and  a.flag=1 and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID) in (SELECT DISTINCT rtrim(b.WAFER) FROM erptemp..tblstockdb_temp  a,erptemp..tblstockdbsub_temp  b WHERE a.FLAG=1 AND a.ORDER_NUM=b.ORDER_NUM AND a.ITEM=b.ITEM ) "
        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            SMR.MoveFirst
            For j = 1 To SMR.RecordCount
                errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("CUSTLOT"))
                SMR.MoveNext
            Next
            MsgBox "�ػ���������" & errormsg & "���ύ�ػ�����", vbInformation, "��ʾ"
            Exit Function
        End If
        '����72��
        errormsg = ""
        If SMR.State = adStateOpen Then SMR.Close
        strSql = "select a.CUSTLOT,a.WAFERID from erptemp..TSV_VT_History_sub a where rtrim(a.CUSTLOT)='" & strlot & "' and    a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID) not in (select rtrim(replace(���̿����,'+','')) from erpdata..tblstocknumsub where �ⷿ���='72' and �ϸ���=0 and ����>0 ) "
        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            SMR.MoveFirst
            For j = 1 To SMR.RecordCount
                errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("WAFERID"))
                SMR.MoveNext
            Next
            MsgBox "�ػ���������" & errormsg & "����72��", vbInformation, "��ʾ"
            Exit Function
        End If
        ' errormsg = ""
        ' If SMR.State = adStateOpen Then SMR.Close
        ' strSql = "select a.CUSTLOT,a.WAFERID,isnull(a.GOODDIEQTY,0),b.���� from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where  rtrim(a.CUSTLOT)='" & strlot & "' and   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. ���̿����,'+','')) and b.�ⷿ���='72' and b.�ϸ���=0 and isnull(a.GOODDIEQTY,0)>b. ���� "
        ' SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        ' If SMR.RecordCount > 0 Then
            ' SMR.MoveFirst
            ' For j = 1 To SMR.RecordCount
                ' errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("WAFERID"))
                ' SMR.MoveNext
            ' Next
            ' MsgBox "�ػ���������" & errormsg & ",�ػ���������72�ֿ��", vbInformation, "��ʾ"
            ' Exit Function
        ' End If
    Next
    RequestNo = GetID()
    For i = 0 To UBound(Split(strLotList, ","))
         strlot = Split(strLotList, ",")(i)
         strSql = "select b.���,b.���̿����,b.������,b.����,b.id from erptemp..TSV_VT_History_sub a  ,erpdata..tblstocknumsub b where  a.flag=1 and  a.CUSTLOT='" & strlot & "' and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b.���̿����,'+','')) and b.����>0 and b.�ϸ���=0 order by a.WAFERID"
         If SMR.State = adStateOpen Then SMR.Close
         SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            SMR.MoveFirst
            For j = 1 To SMR.RecordCount
                If j = 1 Then
                    strSql = " select top 1 a.ԭ�ֿ� from erpdata..tblStockdb a, erpdata..tblStockdbsub b where a.id=b.id and  rtrim(b. ���̿����)='" & Trim(SMR("���̿����")) & "' and a.Ŀ��ֿ�='72'"
                    strmbkf = GetSqlServerStr(strSql)
                End If
                strXH = Trim(SMR("���"))     '���
                strxh_big = ""   '�����
                strgdh = Trim(SMR("������"))       '������
                strLCK = Trim(SMR("���̿����"))    '���̿����
                strlps = Trim(SMR("����"))        '��Ʒ��
                strbls = "0"     '����Ʒ��
                strzcbls = "0"     '�Ƴ̲�����
                strid = Trim(SMR("id"))
                If Get_SqlserverCnt("select * from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid) > 0 Then
                    strSql = "select ITEM from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid
                    intitem = GetSqlServerStr(strSql)
                Else
                    strSql = "select isnull(max(ITEM),0) from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'"
                    intitem = GetSqlServerStr(strSql) + 1

                    If rs.State = adStateOpen Then rs.Close
                    strSql = " select �ⷿ���,���ϱ��,�ͻ�����,isnull(����,0) from erpdata..tblStockNum where id=" & strid
                    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                    If rs.RecordCount = 1 Then
                        rs.MoveFirst
                        strKF = Trim(rs("�ⷿ���"))
                        strmatcode = Trim(rs("���ϱ��"))
                        strCustCode = Trim(rs("�ͻ�����"))
                    
                    End If
                    
            
                  '�ϴ�����

                 '�������,���,���ϱ��, ��������,ԭ�ֿ�,Ŀ��ֿ�,������Ա,����ʱ��,�����Ա,���ʱ��, ���벿��,״̬,REMARK1,REMARK2,REMARK3,REMARK4,REMARK5,ID
                    strSql = "insert into erptemp..tblstockdb_temp(ORDER_NUM,ITEM, MATERIALS,QTY,FORMER, DESTINATION, APPLICANT, APPLICATION_TIME, AUDITOR, AUDIT_TIME, DEPT, FLAG,ID,REMARK1) values( " & _
                    "'" & RequestNo & "'," & intitem & ",'" & strmatcode & "'," & 0 & ",'" & strKF & "','" & strmbkf & "','" & gUserName & "',sysdatetime(),'','','',1," & strid & ",'')"
                
                    AddSql2 (strSql)
                   
                    
                End If
                            
                '�ϴ��ӱ�
                
                '�������, ���, ���, ���̿����, ������, �ϸ���, �Ƴ̲�����, ���ϲ�����, ID
                 strSql = "insert into erptemp..tblstockdbsub_temp(ORDER_NUM,ITEM,WAFER,LOT,GOOD_DIE,BAD1_DIE,BAD2_DIE,ID,REMARK1,QBOX) values( " & _
                "'" & RequestNo & "'," & intitem & ",'" & strLCK & "','" & strgdh & "'," & strlps & "," & strbls & "," & strzcbls & "," & strid & ",'" & strxh_big & "','" & strXH & "')"
              
                AddSql2 (strSql)
                
                'update��������
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
        MsgBox SumCount & "�ʼ�¼����ɹ�", vbInformation, "��ʾ"
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
        MsgBox "����ѡ��ͻ�����"
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

    'KR001,KR009�в���Ʒ���������ͬ�Ĳֱ����Է���������
    RequestNo_Gooddie = GetID()
    RequestNo_Ngdie = Left(RequestNo_Gooddie, Len(RequestNo_Gooddie) - 1) & Right(RequestNo_Gooddie, 1) + 1
     For i = 0 To UBound(Split(strLotList, ","))
         strlot = Split(strLotList, ",")(i)
         If CheckData(strlot) = False Then
             Exit Function
         End If
    Next

     If UCase(Trim(CboCustomer.text)) = "KR001" Or UCase(Trim(CboCustomer.text)) = "KR009" Then
     
        '��һ�β��䣬��gooddie������
         RequestNo = RequestNo_Gooddie
         boxsplit (1)
         Call CommitApplication(RequestNo, 1, strLotList)
         
         '�ڶ��β��䣬��ngdie������
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
    '���ύ������
    ' If SMR.State = adStateOpen Then SMR.Close
    ' strSql = "select a.CUSTLOT,a.WAFERID from erptemp..TSV_VT_History_sub a where rtrim(a.CUSTLOT)='" & strlot & "' and  a.flag=1 and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID) in (SELECT DISTINCT rtrim(b.WAFER) FROM erptemp..tblstockdb_temp  a,erptemp..tblstockdbsub_temp  b WHERE a.FLAG=1 AND a.ORDER_NUM=b.ORDER_NUM AND a.ITEM=b.ITEM ) "
    ' SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    ' If SMR.RecordCount > 0 Then
        ' SMR.MoveFirst
        ' For j = 1 To SMR.RecordCount
            ' errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("CUSTLOT"))
            ' SMR.MoveNext
        ' Next
        ' MsgBox "�ػ���������" & errormsg & "���ύ�ػ�����", vbInformation, "��ʾ"
        ' Exit Function
    ' End If
    '����72��
    errormsg = ""
    If SMR.State = adStateOpen Then SMR.Close
    strSql = "select a.CUSTLOT,a.WAFERID from erptemp..TSV_VT_History_sub a where  a.flag=1 and rtrim(a.CUSTLOT)='" & strlot & "'  and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID) not in (select rtrim(replace(���̿����,'+','')) from erpdata..tblstocknumsub where �ⷿ���='72' and �ϸ���=0 and ����>0 ) "
 
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For j = 1 To SMR.RecordCount
            errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("WAFERID"))
            SMR.MoveNext
        Next
        MsgBox "�ػ���������" & errormsg & "����72��", vbInformation, "��ʾ"
        Exit Function
    End If
    '�ػ���������72�ֿ��
    If UCase(Trim(CboCustomer.text)) = "KR001" Or UCase(Trim(CboCustomer.text)) = "KR009" Then
        
        errormsg = ""
        If SMR.State = adStateOpen Then SMR.Close
        strSql = "select a.CUSTLOT,a.WAFERID, convert(int,isnull(a.GOODDIEQTY,0))+ convert(int,isnull(a.NGDIEQTY,0)),b.���� from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where  rtrim(a.CUSTLOT)='" & strlot & "' and   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. ���̿����,'+','')) and b.�ⷿ���='72' and b.�ϸ���=0 and  convert(int,isnull(a.GOODDIEQTY,0))+ convert(int,isnull(a.NGDIEQTY,0))>b. ���� "

        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            SMR.MoveFirst
            For j = 1 To SMR.RecordCount
                errormsg = errormsg & "," & Trim(SMR("CUSTLOT")) & Trim(SMR("WAFERID"))
                SMR.MoveNext
            Next
            MsgBox "�ػ���������" & errormsg & ",�ػ���������72�ֿ��", vbInformation, "��ʾ"
            Exit Function
        End If
    End If
    CheckData = True

End Function

Function boxsplit(Index As Integer) 'index=1 ��һ�β�GoodDie��index=2 �ڶ��β�NGdie

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

    '1.ȡ�������
    NewBoxList = ""
    OldBoxList = ""
    boxsplit = True
    If Index = 1 Then
        strSql = "select distinct b.��� from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. ���̿����,'+','')) and b.�ⷿ���='72' and b.�ϸ���=0 and isnull(a.GOODDIEQTY,0)<=b. ���� and b.��� not like '%VT%'"
    ElseIf Index = 2 Then
        strSql = "select distinct b.��� from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. ���̿����,'+','')) and b.�ⷿ���='72' and b.�ϸ���=0 and isnull(a.NGDIEQTY,0)<=b. ���� and b.��� not like '%VT%'"
    End If
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For j = 1 To SMR.RecordCount
            If OldBoxList = "" Then
                OldBoxList = Trim(SMR("���"))
            Else
                OldBoxList = OldBoxList & "," & Trim(SMR("���"))
            End If
            strSql = " SELECT COUNT(*) FROM erpdata..tblStockNumTree c WHERE  c.��� LIKE '" & Trim(SMR("���")) & "' + '%' "
               
            If rs.State = adStateOpen Then rs.Close
            rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            qnum = Trim(Str(rs.Fields(0).Value))
            If qnum = 1 Then
                strboxid_new = Trim(SMR("���")) + "_VT"
            Else
                 strboxid_new = Trim(SMR("���")) & "_VT" & qnum - 1
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
    '2.�������������
    strboxid_new = ""
    If NewBoxList <> "" Then
        For i = 0 To UBound(Split(NewBoxList, ","))
            strboxid_new = Split(NewBoxList, ",")(i)
            strboxid_old = Split(OldBoxList, ",")(i)
            AddSql2 ("INSERT INTO erpdata..TBLPACKMAININF(���,�ͻ�����,����,���߱��,�ϸ���,װ����)  VALUES ('" & strboxid_new & "','" & UCase(Trim(CboCustomer.text)) & "',1,'1','0','1')")
            AddSql2 ("INSERT INTO erpdata..tblPackTreeInf(���) VALUES ('" & strboxid_new & "')")
            AddSql2 ("INSERT INTO erpdata..tblStockNumTree ( ���,���,�ϼ����,������,������) SELECT b.���,b.���,b.�ϼ����,b.������,'0' FROM erpdata..tblPackTreeInf b WHERE b.��� = '" & strboxid_new & "' ")
            If Index = 1 Then
                 strSql = "select b.���̿����,ISNULL(a.GOODDIEQTY,0),b.���� -ISNULL(a.GOODDIEQTY,0)  from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. ���̿����,'+','')) and b.�ⷿ���='72' and b.�ϸ���=0 and isnull(a.GOODDIEQTY,0) >0 and isnull(a.GOODDIEQTY,0)<=b. ���� and  rtrim(b.���)='" & strboxid_old & "' "
            Else
                 strSql = "select b.���̿����,ISNULL(a.NGDIEQTY,0) ,b.���� -ISNULL(a.NGDIEQTY,0)  from erptemp..TSV_VT_History_sub a ,erpdata..tblstocknumsub b where   a.flag=1 and  rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b. ���̿����,'+','')) and b.�ⷿ���='72' and b.�ϸ���=0 and isnull(a.NGDIEQTY,0) >0 and isnull(a.NGDIEQTY,0)<=b. ���� and  rtrim(b.���)='" & strboxid_old & "' "
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
                         AddSql2 (" INSERT INTO erpdata..tblStockNumSub  SELECT '" & strboxid_new & "',a.���̿����,a.������,'" & splitqty & "',a.�Ϻ�,a.���ϱ��,a.�ϸ���,a.�������  ,a.ID,a.�ⷿ���,GETDATE(),a.�󹤵� FROM erpdata..tblStockNumSub a  WHERE a.��� = '" & strboxid_old & "' AND a.���̿���� = '" & WAFER & "'")
                         AddSql2 (" UPDATE erpdata..tblStockNumSub SET ���� = ���� - " & splitqty & " WHERE ��� = '" & strboxid_old & "' AND ���̿���� = '" & WAFER & "' ")
                    ElseIf splitqty > 0 Then
                         AddSql2 ("  UPDATE erpdata..tblStockNumSub SET ��� = '" & strboxid_new & "' WHERE ��� = '" & strboxid_old & "' AND ���̿���� = '" & WAFER & "' ")
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
            strSql = "select b.���,b.���̿����,b.������,b.����,b.id from erptemp..TSV_VT_History_sub a  ,erpdata..tblstocknumsub b where  a.flag=1 and  a.CUSTLOT='" & strlot & "' and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b.���̿����,'+','')) and b.����>0 and b.�ϸ���=0 order by a.WAFERID"
        ElseIf (UCase(Trim(CboCustomer.text)) = "KR001" Or UCase(Trim(CboCustomer.text)) = "KR009") Then
            If Index = 1 Then
                strSql = "select b.���,b.���̿����,b.������,b.����,b.id from erptemp..TSV_VT_History_sub a  ,erpdata..tblstocknumsub b where  a.flag=1 and  a.CUSTLOT='" & strlot & "' and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b.���̿����,'+','')) and b.����>0  and b.����=a.GOODDIEQTY and b.�ϸ���=0 order by a.WAFERID"
            ElseIf Index = 2 Then
                strSql = "select b.���,b.���̿����,b.������,b.����,b.id from erptemp..TSV_VT_History_sub a  ,erpdata..tblstocknumsub b where  a.flag=1 and  a.CUSTLOT='" & strlot & "' and rtrim(a.CUSTLOT)  + rtrim(a.WAFERID)=rtrim(replace(b.���̿����,'+','')) and b.����>0  and b.����=a.NGDIEQTY and b.�ϸ���=0 order by a.WAFERID"
            End If
        End If

        If SMR.State = adStateOpen Then SMR.Close
        SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        If SMR.RecordCount > 0 Then
            SMR.MoveFirst
            For j = 1 To SMR.RecordCount
                If j = 1 Then
                    strSql = " select top 1 rtrim(a.ԭ�ֿ�) from erpdata..tblStockdb a, erpdata..tblStockdbsub b where a.id=b.id and  rtrim(b. ���̿����)='" & Trim(SMR("���̿����")) & "' and a.Ŀ��ֿ�='72'"
                    strmbkf = GetSqlServerStr(strSql)
                    '����ƷҪ���ض�Ӧ�Ĳ���Ʒ��
                    '07,20 ��ί��ģ���˰����Ʒ����30
                    '16,19��ί��ģ��Ǳ�˰����Ʒ����28
    
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
                If GetNewBoxId(Trim(SMR("���"))) <> Trim(SMR("���")) Then
                    SMR.MoveNext
                    GoTo 1  'Ϊ������������2���������һ�£����ѯ�����ʽ��
                End If
                
                strXH = Trim(SMR("���"))     '���
                strxh_big = ""   '�����
                strgdh = Trim(SMR("������"))       '������
                strLCK = Trim(SMR("���̿����"))    '���̿����
                strlps = Trim(SMR("����"))        '��Ʒ��
                strbls = "0"     '����Ʒ��
                strzcbls = "0"     '�Ƴ̲�����
                strid = Trim(SMR("id"))
                If Get_SqlserverCnt("select * from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid) > 0 Then
                    strSql = "select ITEM from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'and  id=" & strid
                    intitem = GetSqlServerStr(strSql)
                Else
                    strSql = "select isnull(max(ITEM),0) from erptemp..tblstockdb_temp where ORDER_NUM='" & RequestNo & "'"
                    intitem = GetSqlServerStr(strSql) + 1
    
                    If rs.State = adStateOpen Then rs.Close
                    strSql = " select �ⷿ���,���ϱ��,�ͻ�����,isnull(����,0) from erpdata..tblStockNum where id=" & strid
                   
                    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                    If rs.RecordCount = 1 Then
                        rs.MoveFirst
                        strKF = Trim(rs("�ⷿ���"))
                        strmatcode = Trim(rs("���ϱ��"))
                        strCustCode = Trim(rs("�ͻ�����"))
                    
                    End If
            
                  '�ϴ�����
    
                 '�������,���,���ϱ��, ��������,ԭ�ֿ�,Ŀ��ֿ�,������Ա,����ʱ��,�����Ա,���ʱ��, ���벿��,״̬,REMARK1,REMARK2,REMARK3,REMARK4,REMARK5,ID
                    strSql = "insert into erptemp..tblstockdb_temp(ORDER_NUM,ITEM, MATERIALS,QTY,FORMER, DESTINATION, APPLICANT, APPLICATION_TIME, AUDITOR, AUDIT_TIME, DEPT, FLAG,ID,REMARK1) values( " & _
                    "'" & RequestNo & "'," & intitem & ",'" & strmatcode & "'," & 0 & ",'" & strKF & "','" & strmbkf & "','" & gUserName & "','','','','',1," & strid & ",'')"
                
                    AddSql2 (strSql)
                                      
                End If
                            
                '�ϴ��ӱ�
                
                '�������, ���, ���, ���̿����, ������, �ϸ���, �Ƴ̲�����, ���ϲ�����, ID
                 strSql = "insert into erptemp..tblstockdbsub_temp(ORDER_NUM,ITEM,WAFER,LOT,GOOD_DIE,BAD1_DIE,BAD2_DIE,ID,REMARK1,QBOX) values( " & _
                "'" & RequestNo & "'," & intitem & ",'" & strLCK & "','" & strgdh & "'," & strlps & "," & strbls & "," & strzcbls & "," & strid & ",'" & strxh_big & "','" & strXH & "')"
              
                AddSql2 (strSql)
                
                'update��������
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
        MsgBox SumCount & "�ʼ�¼����ɹ�", vbInformation, "��ʾ"
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
'���ɷ�ʽ��FWW+YYMMDD +4λ��ˮ��
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
'����6+4����ˮ����10��
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

strSql = "Select Isnull(max(�ػ�����),0) as �ػ����� from erptemp..TSV_VT_History_sub where left(�ػ�����,6)='" & CODE & "'"

If SMR.State = adStateOpen Then SMR.Close
SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
If SMR("�ػ�����") = 0 Then

    GetVTID = CODE & "001"
Else
    GetVTID = Val(SMR("�ػ�����")) + 1
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


strSql = " SELECT distinct 'GCSH' AS 'Sub Name','HTKS' AS 'Ship To', t1.FAB_CONV_ID AS 'FAB Device',t1.CUSTDEVICE AS 'Customer Device' ,LEFT(t1.IMAGER_CUSTOMER_REV,2) + d.�������� AS 'GC Version' ," & _
" t1.PO_NUM AS 'PO NO' ,'' AS WO,'' AS 'Invoice NO', convert(nvarchar(20),GETDATE(),111) AS 'FAB-Out DATE',t1.CUSTLOT AS 'FAB Lot ID',t1.WAFERID AS 'Wafer ID'," & _
" t1.PASSBINCOUNT AS 'Gross Dies','' AS 'Sampling Qty','' AS 'Pass Dies' ,'' AS Yield,t1.REMARK1 AS Remark ,t1.��ע,t1. ���ڻ��� FROM (" & _
" SELECT c.FAB_CONV_ID ,  a.CUSTDEVICE   ,c.IMAGER_CUSTOMER_REV   ,c.PO_NUM ,c.MTRL_NUM, a.CUSTLOT  , a.WAFERID ,b.PASSBINCOUNT, A.REMARK1," & _
" ISNULL(a.REMARK2,'') AS ��ע,ISNULL(a.REMARK3,'')  AS  ���ڻ���" & _
" FROM ERPTEMP..TSV_VT_History_sub a" & _
" LEFT JOIN erpbase..tblmappingdata b ON a.CUSTLOT =b.lotid  AND a.WAFERID =right(100+b.WAFER_ID,2)  AND a.customershortname=b.customershortname  " & _
" LEFT JOIN erpbase..tblCustomerOI c  ON convert(VARCHAR(50),c.id)=b.filename  AND b.lotid=c.SOURCE_BATCH_ID  AND a.CUSTDEVICE +'-3'=c.MPN_DESC  " & _
" WHERE  a.FLAG_WO =1 AND a.customershortname='GC' ) t1" & _
" LEFT JOIN ERPTEMP..GcCode_Reference d ON  t1.CUSTDEVICE  +'-3'=d.�ͻ������� AND d.�Ƴ�=t1.��ע" & _
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
        xlsSheet.Cells(1, 18) = "��ע"
        xlsSheet.Cells(1, 19) = "���ڻ���"
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
            xlsSheet.Cells(i + 1, 18) = Trim(SMR("��ע"))
            xlsSheet.Cells(i + 1, 19) = Trim(SMR("���ڻ���"))
            
            
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
    MsgBox "WOת�����", vbInformation, "��ʾ"
    
               
    
    
Ert:

    If Not (xlsApp Is Nothing) Then
        
        Set xlsApp = Nothing
        Set xlsSheet = Nothing
        Set xlsBook = Nothing

    End If
    

End Sub


Private Sub UploadVTWO_GC()

'GC�ػ�WO�ϴ�

End Sub

Private Function wafer_to_string(WAFERLIST As String) As String
Dim TEMP As String
Dim String2 As String
Dim bb() As String
Dim b() As String
Dim i As Integer
Dim j As Integer
b = Split(WAFERLIST, ",")

Last = UBound(b) - LBound(b) + 1  '��ȡ�����С

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
'��ȡWO�еĶ�������ǰ2��
 Dim strSql As String
 strSql = "SELECT distinct left(b.IMAGER_CUSTOMER_REV,2) FROM erpbase..tblmappingData a inner join ERPBASE..TBLCUSTOMEROI b ON  convert(VARCHAR(30),b.ID)=a.FILENAME AND b.SOURCE_BATCH_ID=a.LOTID and a.lotid='" & strlot & "' and convert(int,wafer_id)=" & CInt(strWaferID)
 GetGcrevFromWO = GetSqlServerStr(strSql)

End Function


Private Function GetHTDevice(strCustDevice As String, strtype As String, strGcrev2 As String)
'ȡ�ó��ڻ���
    If Get_SqlserverCnt("select distinct ���ڻ����� from ERPTEMP..GcCode_Reference where �ͻ�������='" & strCustDevice & "-3' and �Ƴ�='" & strtype & "'") <> 1 Then
    
        If Get_SqlserverCnt("select distinct ���ڻ����� from ERPTEMP..GcCode_Reference where �ͻ�������='" & strCustDevice & "-3' and �Ƴ�='" & strtype & "' and ��������ڶ�λ ='" & strGcrev2 & "'") <> 1 Then
            GetHTDevice = ""
        Else
            GetHTDevice = GetSqlServerStr("select distinct ���ڻ����� from ERPTEMP..GcCode_Reference where �ͻ�������='" & strCustDevice & "-3' and �Ƴ�='" & strtype & "' and ��������ڶ�λ ='" & strGcrev2 & "'")
        End If
    Else
        GetHTDevice = GetSqlServerStr("select distinct ���ڻ����� from ERPTEMP..GcCode_Reference where �ͻ�������='" & strCustDevice & "-3' and �Ƴ�='" & strtype & "'")
    End If
End Function






