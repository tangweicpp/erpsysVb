VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmWareHousingCharges 
   BackColor       =   &H8000000B&
   Caption         =   "WLP入库即收费通用平台"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11355
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
   ScaleHeight     =   7530
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   11295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   19923
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&I 客户挑料信息(DN)导入"
      TabPicture(0)   =   "FrmWareHousingCharges.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DNContentFps"
      Tab(0).Control(1)=   "DNFileDirText"
      Tab(0).Control(2)=   "OpenFileButton"
      Tab(0).Control(3)=   "SaveDNButton"
      Tab(0).Control(4)=   "Dialog"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&P 相关外层标签打印"
      TabPicture(1)   =   "FrmWareHousingCharges.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ReelIDFps"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "AutoBoxedOption"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ManualBoxedOption"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "BoxedButton"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ScanText"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   -65640
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton SaveDNButton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "保存"
         Height          =   360
         Left            =   -67200
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton OpenFileButton 
         BackColor       =   &H00E0E0E0&
         Caption         =   "浏览..."
         Height          =   360
         Left            =   -68160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DNFileDirText 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   6495
      End
      Begin VB.TextBox ScanText 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   4
         Top             =   1530
         Width           =   3615
      End
      Begin VB.CommandButton BoxedButton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "合箱"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton ManualBoxedOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "手动合箱"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton AutoBoxedOption 
         BackColor       =   &H00E0E0E0&
         Caption         =   "自动合箱"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin FPSpreadADO.fpSpread ReelIDFps 
         Height          =   7815
         Left            =   480
         TabIndex        =   6
         Top             =   1920
         Width           =   10095
         _Version        =   524288
         _ExtentX        =   17806
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
         MaxCols         =   6
         MaxRows         =   0
         SpreadDesigner  =   "FrmWareHousingCharges.frx":0038
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread DNContentFps 
         Height          =   5535
         Left            =   -74760
         TabIndex        =   9
         Top             =   1200
         Width           =   10095
         _Version        =   524288
         _ExtentX        =   17806
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
         MaxCols         =   100
         MaxRows         =   0
         SpreadDesigner  =   "FrmWareHousingCharges.frx":045A
         AppearanceStyle =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描标签"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1560
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmWareHousingCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LabelHistoryPrintData As String

Private Sub InitializeControls()
With DNContentFps
    .ReDraw = False
    .MaxCols = 10
    .MaxRows = 0
    .FontBold = False
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .Row = 0
    .TypeHAlign = TypeHAlignLeft
    .TypeVAlign = TypeVAlignCenter
    

End With

With ReelIDFps
    .TypeMaxEditLen = 500
    .ReDraw = False
    .MaxCols = 2
    .MaxRows = 0
    .FontBold = False
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .Row = 0
    .TypeHAlign = TypeHAlignLeft
    .TypeVAlign = TypeVAlignCenter
    
    .SelForeColor = &HFF8080
    .SetText 1, 0, "箱号唯一码"
    .ColWidth(1) = 16
    
    .SetText 2, 0, "打印记录"
    .ColWidth(2) = 40
End With
End Sub

Private Sub AutoBoxedOption_Click()
BoxedButton.Visible = False
End Sub

Private Sub Form_Activate()
ScanText.SetFocus
End Sub

Private Sub Form_Load()
Call InitializeConfigIni
Call InitializeControls
End Sub

Private Sub InitializeConfigIni()





End Sub

Private Sub ManualBoxedOption_Click()
BoxedButton.Visible = True
End Sub

Private Sub OpenFileButton_Click()

On Error GoTo ErrHandler

Dim name As String

Dialog.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
Dialog.ShowOpen
name = Dialog.filename
If name <> "" Then
    DNFileDirText.text = name
    Call ShowOpenFile

End If

Exit Sub
ErrHandler:
Dialog.filename = ""
End Sub

Private Function ShowOpenFile()

On Error GoTo ErrHandle

Dim VBExcel     As Excel.Application
Dim xlBook      As Excel.Workbook
Dim xlSheet     As Excel.Worksheet
Dim i           As Integer
Dim j           As Integer
Dim strChar     As String
Dim strTmp(10)  As Variant

MousePointer = 11
DNContentFps.MaxRows = 0

Set VBExcel = CreateObject("excel.application")
VBExcel.Visible = False
Set xlBook = VBExcel.Workbooks.Open(DNFileDirText.text)
Set xlSheet = xlBook.Worksheets(1)
If xlSheet.Range("A1").CurrentRegion.Columns.count < 2 Then
    MousePointer = 0
    MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
    GoTo EXITPRO
    Exit Function

End If

With DNContentFps
    For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.count
        strTmp(0) = Trim(xlSheet.Range("A" & i).Value)
        If Len(strTmp(0)) > 0 Then
            If i <> 1 Then .MaxRows = .MaxRows + 1

            For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.count
                If j > 26 Then
                    strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
                Else
                    strChar = Chr(96 + j)

                End If

                If i = 1 Then
                    .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))
                Else
                    .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))

                End If

            Next

        End If

    Next

End With

MousePointer = 0
xlBook.Close
Set xlSheet = Nothing
Set xlBook = Nothing
Set VBExcel = Nothing
VBExcel.Quit
Exit Function
EXITPRO:

On Error Resume Next

MousePointer = 0
If Not VBExcel Is Nothing Then
    xlBook.Close
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
    VBExcel.Quit

End If

Exit Function
ErrHandle:
GoTo EXITPRO

End Function

Private Sub ScanText_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then Exit Sub

'check boxID then list it
Dim text As String

text = UCase(Trim$(ScanText.text))
If Not CheckText(text) Then
    ScanText.text = ""
    Exit Sub
End If

Call ListText(text)

ScanText.text = ""
End Sub

Private Function CheckText(text As String) As Boolean
CheckText = False

Dim sql As String
Dim printerNameID As String
Dim bartenderName As String
Dim lableID As String

printerNameID = GetIni(Me.Tag, "PrinterNameID")
bartenderName = GetIni(Me.Tag, "BartenderName")
lableID = GetIni(Me.Tag, "LABEL_ID")

sql = "select * from WLP_SHIP_BOXID_DETAIL where PSN = '" & text & "'"
If Get_OracleCnt(sql) > 0 Then
    MsgBox "该箱号已经打印过,请勿重复扫描", vbInformation, "提示"
    Exit Function

End If

If IsPSNRepeated(text) = True Then
    MsgBox "该箱号已经扫描过,请勿重复扫描", vbInformation, "提示"
    Exit Function

End If

sql = "select top 1 Content from erpdata.dbo.tblME_PrintInfo " & _
" where PrinterNameID = '" & printerNameID & "' and BartenderName = '" & bartenderName & "' and LABEL_ID = '" & lableID & "' " & _
" and EVENT_SOURCE = 'PKG' and charindex('" & text & "',Content) > 0 " & _
" order by ID desc "
LabelHistoryPrintData = Get_SqlStr(sql)
If LabelHistoryPrintData = "" Then
    MsgBox "查询不到打印历史", vbInformation, "提示"
    Exit Function

End If

'sql = "select * from erpdata..tblStockNumSub where 箱号 = '" & text & "' "
'If Get_SqlserverCnt(sql) = 0 Then
'    MsgBox "库存中查询不到该箱号: " & text & vbCrLf & "请使用该箱号入库,否则禁止出货", vbExclamation, "错误"
'    Exit Function
'
'End If

CheckText = True
End Function

Private Function IsPSNRepeated(text As String) As Boolean
Dim i As Integer
IsPSNRepeated = False

With ReelIDFps
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .text = text Then
            IsPSNRepeated = True
        End If
        
    Next
    
End With

End Function

Private Function ListText(text As String)
Dim i As Long

With ReelIDFps
    .MaxRows = .MaxRows + 1
    i = .MaxRows
    .SetText 1, i, text
    .SetText 2, i, LabelHistoryPrintData
End With

'Call PlaySound("卷盘已扫描")

End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab

    Case 0
    
    Case 1
        ScanText.SetFocus
End Select

End Sub


