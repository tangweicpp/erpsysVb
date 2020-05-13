VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.ocx"
Begin VB.Form FrmDNModule 
   Caption         =   "AH017DN 标签"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17940
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
   ScaleHeight     =   8400
   ScaleWidth      =   17940
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9735
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   17171
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "DN 维护"
      TabPicture(0)   =   "FrmDNModule.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "标签转换"
      TabPicture(1)   =   "FrmDNModule.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtQRCode"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ComboPrinter"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ComboModel"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtPecs"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.TextBox txtPecs 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   6840
         TabIndex        =   20
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox ComboModel 
         BackColor       =   &H00FFC0FF&
         Height          =   315
         Left            =   6840
         TabIndex        =   18
         Top             =   3060
         Width           =   2895
      End
      Begin VB.ComboBox ComboPrinter 
         BackColor       =   &H00FFC0FF&
         Height          =   315
         Left            =   6840
         TabIndex        =   15
         Top             =   2700
         Width           =   2895
      End
      Begin VB.TextBox txtQRCode 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   870
         Width           =   6735
      End
      Begin VB.Frame Frame1 
         Height          =   8655
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   17655
         Begin VB.TextBox txtDNFileName 
            BackColor       =   &H00FFC0FF&
            Height          =   285
            Left            =   960
            TabIndex        =   9
            Top             =   360
            Width           =   5775
         End
         Begin VB.CommandButton btnOpen 
            BackColor       =   &H00E0E0E0&
            Caption         =   ". . ."
            Height          =   285
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton btnSave 
            BackColor       =   &H00C0C0C0&
            Caption         =   "保 存 &S"
            Height          =   405
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton btnQuery 
            BackColor       =   &H00C0C0C0&
            Caption         =   "查 询 &G"
            Height          =   405
            Left            =   2580
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton btnDelete 
            BackColor       =   &H00C0C0C0&
            Caption         =   "删 除 &D"
            Height          =   405
            Left            =   1410
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton btnOutput 
            BackColor       =   &H00C0C0C0&
            Caption         =   "导 出 &E"
            Height          =   405
            Left            =   3750
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton btnExit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "退 出 &Q"
            Height          =   405
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtDN 
            BackColor       =   &H00FFC0FF&
            Height          =   285
            Left            =   960
            TabIndex        =   2
            Top             =   675
            Width           =   2055
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   7080
            Top             =   1080
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin FPSpreadADO.fpSpread Fps 
            Height          =   5775
            Left            =   240
            TabIndex        =   10
            Top             =   2280
            Width           =   16335
            _Version        =   524288
            _ExtentX        =   28813
            _ExtentY        =   10186
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
            MaxCols         =   50
            MaxRows         =   0
            SpreadDesigner  =   "FrmDNModule.frx":0038
            TextTip         =   2
            AppearanceStyle =   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D.N文件"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   405
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D.N"
            Height          =   195
            Left            =   600
            TabIndex        =   11
            Top             =   720
            Width           =   270
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打印份数"
         Height          =   195
         Left            =   6000
         TabIndex        =   19
         Top             =   3480
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标签模板名"
         Height          =   195
         Left            =   5880
         TabIndex        =   17
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打印机"
         Height          =   195
         Left            =   6240
         TabIndex        =   16
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标签二维码"
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   960
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmDNModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type AH033_SN
A_DN As String
B_PO As String
C_DEVICE As String
D_LOT As String
E_QTY As String
F_CPN As String
G_VENDOR_CODE As String
H_SHIP_ADDR As String
I_SHIP_DATE As String
J_SHIP_BY As String

End Type

Private Type LBL_INFO
CPN As String
MPN As String
QTY As String
LOTNO As String
DATECODE As String
MANUFACTURE As String
VENDORCODE As String
COUNTRY As String

End Type

Private Sub btnDelete_Click()
If txtDN.text = "" Then
    MsgBox "请输入要删除的DN", vbInformation, "提示"
    Exit Sub
End If

If MsgBox("是否确定删除", vbYesNo, "提示") = vbNo Then
    Exit Sub
End If

AddSql ("delete from TBL_AH017_DN where A_DN = '" & Trim$(txtDN.text) & "'")
MsgBox "删除成功", vbInformation, "提示"
End Sub

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnOpen_Click()
With CommonDialog1
    .Filter = "Excel文件(*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv"
    .ShowOpen
    txtDNFileName.text = CommonDialog1.filename
End With

End Sub

Private Sub btnOutput_Click()
Call ExportDataByInvoiceNO(Trim(txtDN.text))
End Sub

Private Sub btnQuery_Click()
Call ShowDataByInvoiceNO(Trim(txtDN.text))
End Sub


Private Sub btnSave_Click()
If txtDNFileName.text = "" Then
    MsgBox "请打开你要上传的DN文件", vbCritical, "警告"
    Exit Sub
End If

If MsgBox("是否确定要上传", vbYesNo, "提示") = vbNo Then
    Exit Sub
End If

Call SaveFile(Trim$(txtDNFileName.text))

End Sub

Private Function SaveFile(strFileName As String)
Dim xlApp   As Excel.Application
Dim xlBook  As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim dataArray()    As AH033_SN
Dim lColsCnt
Dim lRowsCnt
Dim i

On Error GoTo hErr

Set xlApp = CreateObject("excel.application")
xlApp.Visible = False
Set xlBook = xlApp.Workbooks.Open(strFileName)
Set xlSheet = xlBook.Worksheets(1)
lColsCnt = xlSheet.Range("A1").CurrentRegion.Columns.count
lRowsCnt = xlSheet.Range("A1").CurrentRegion.Rows.count

If lRowsCnt < 2 Then
    MsgBox "Excel中的行数至少2行", vbInformation, "提示"
    GoTo hErr
End If

ReDim dataArray(lRowsCnt - 1)

For i = 2 To lRowsCnt

    With dataArray(i - 2)
        .A_DN = Trim("" & Replace(Replace(xlSheet.Range("A" & i), Chr(10), ""), Chr(13), ""))
        .B_PO = Trim("" & Replace(Replace(xlSheet.Range("B" & i), Chr(10), ""), Chr(13), ""))
        .C_DEVICE = Trim("" & Replace(Replace(xlSheet.Range("C" & i), Chr(10), ""), Chr(13), ""))
        .D_LOT = Trim("" & Replace(Replace(xlSheet.Range("D" & i), Chr(10), ""), Chr(13), ""))
        .E_QTY = Trim("" & Replace(Replace(xlSheet.Range("E" & i), Chr(10), ""), Chr(13), ""))
        .F_CPN = Trim("" & Replace(Replace(xlSheet.Range("F" & i), Chr(10), ""), Chr(13), ""))
        .G_VENDOR_CODE = Trim("" & Replace(Replace(xlSheet.Range("G" & i), Chr(10), ""), Chr(13), ""))
        .H_SHIP_ADDR = Trim("" & Replace(Replace(xlSheet.Range("H" & i), Chr(10), ""), Chr(13), ""))
        .I_SHIP_DATE = Trim("" & Replace(Replace(xlSheet.Range("I" & i), Chr(10), ""), Chr(13), ""))
        .J_SHIP_BY = Trim("" & Replace(Replace(xlSheet.Range("J" & i), Chr(10), ""), Chr(13), ""))
        
    End With

Next i

Call InsertToDB(dataArray)

hErr:
xlBook.Close
Set xlSheet = Nothing
Set xlBook = Nothing
xlApp.Quit
Set xlApp = Nothing

End Function

Private Function InsertToDB(ByRef dataArray() As AH033_SN)
Dim strSql As String
Dim i As Integer
Dim lUploadID As Long

lUploadID = Get_OracleNo("select SEQ_AH033_UPLOAD_ID.Nextval from dual")

On Error GoTo hErr

Cnn.BeginTrans

For i = 0 To UBound(dataArray) - 1

    With dataArray(i)
        strSql = "insert into TBL_AH017_DN(A_DN,B_PO,C_DEVICE,D_LOT,E_QTY,F_CPN,G_VENDOR_CODE,H_SHIP_ADDR,I_SHIP_DATE,J_SHIP_BY,UPLOAD_BY,UPLOAD_DATE,UPLOAD_ID) " & _
           " values('" & .A_DN & "','" & .B_PO & "','" & .C_DEVICE & "','" & .D_LOT & "','" & .E_QTY & "','" & .F_CPN & "','" & .G_VENDOR_CODE & "','" & .H_SHIP_ADDR & "','" & .I_SHIP_DATE & "','" & .J_SHIP_BY & "','" & gUserRealName & "',sysdate,'" & lUploadID & "')"
        
        AddSql (strSql)

    End With

Next
Cnn.CommitTrans
MsgBox "上传完成", vbInformation, "提示"

Call ShowUploadData(lUploadID)

Exit Function
hErr:

If InStr(Err.DESCRIPTION, "违反唯一约束条件") > 0 Then
    MsgBox "上传重复,系统中已经存在相同的INVOICE NO-机种上传记录" & vbCrLf & "请删除该笔,否则无法重新上传", vbCritical, "错误提示"
Else
    
End If


Cnn.RollbackTrans

End Function

Private Sub ShowUploadData(lUploadID As Long)
Dim strSql As String
Dim rs As New ADODB.Recordset

strSql = "select * from TBL_AH017_DN where UPLOAD_ID = " & lUploadID & " "

Set rs = Get_OracleRs(strSql)

With fps
    .MaxRows = 0
    If rs.RecordCount > 0 Then
        Set .DataSource = rs
    Else
        MsgBox "没有查询到有效数据", vbInformation, "提示"
    End If

End With

End Sub

Private Sub ShowDataByInvoiceNO(strInvoice As String)
Dim strSql As String
Dim rs As New ADODB.Recordset

If strInvoice = "" Then
    strSql = "select * from TBL_AH017_DN order by upload_date desc"
Else
    strSql = "select * from TBL_AH017_DN where A_DN = '" & strInvoice & "'"
End If

Set rs = Get_OracleRs(strSql)

With fps
    .MaxRows = 0
    If rs.RecordCount > 0 Then
        Set .DataSource = rs
    Else
        MsgBox "没有查询到有效数据", vbInformation, "提示"
    End If

End With

End Sub

Private Sub ExportDataByInvoiceNO(strInvoice As String)
Dim strSql As String

If strInvoice = "" Then
    strSql = "select * from TBL_AH017_DN order by upload_date desc"
Else
    strSql = "select * from TBL_AH017_DN where invoice_no = '" & strInvoice & "'"
End If

Call ExporToExcel(strSql)
End Sub

Private Sub Form_Load()

ComboPrinter.AddItem ("W_IN_2B5F_2")
ComboModel.AddItem ("FT_HW")
txtPecs.text = "1"
End Sub

Private Sub txtQRCode_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then Exit Sub
If Len(Trim$(txtQRCode.text)) = 0 Then Exit Sub

If ComboPrinter.text = "" Or ComboModel.text = "" Or txtPecs.text = "" Then
    MsgBox "请选择打印机，标签模板，打印份数", vbCritical, "提示"
    Exit Sub
End If

Call Converter(UCase(Trim$(txtQRCode.text)))
MsgBox "打印完成", vbInformation, "提醒"

End Sub

Private Sub Converter(strQRCode As String)
Dim strArray() As String
Dim strPrinter As String
Dim strModel As String
Dim strModel2 As String
Dim strPecs As String
Dim strLotID As String
Dim strHWPN As String
Dim strVendCode As String
Dim strSql As String
Dim rs As New ADODB.Recordset

PrintIBoxLbl = False

strPrinter = UCase$(Trim$(ComboPrinter.text))
strModel = UCase$(Trim$(ComboModel.text))
strModel2 = strModel & ".btw"
strPecs = UCase$(Trim$(txtPecs.text))

strSql = "select top 1 Content,ID from erpdata.dbo.tblME_PrintInfo where BartenderName = '" & strModel2 & "' and PrinterNameID = '" & strPrinter & "' and EVENT_SOURCE = 'PKG' and LABEL_ID = '" & strModel & "' and charindex('" & strQRCode & "',Content) > 0 "
Set rs = Get_SqlserveRs(strSql)
If rs.RecordCount = 0 Then
    MsgBox "查询不到打印记录，请确定打印机和打印模板是否选错？", vbCritical, "警告"
    Exit Sub
End If

strContent = Trim("" & rs!Content)
strid = Trim("" & rs!id)
If InStr(strContent, ";") = 0 Then
    MsgBox "查询不到该标签的打印记录,无法补打", vbCritical, "提示"
    Exit Sub

End If

strLotID = Mid(strQRCode, InStr(strQRCode, "1T") + 2, 8)

strSql = "select C_DEVICE from TBL_AH017_DN where D_LOT = '" & strLotID & "' "
strHWPN = Get_OracleStr(strSql)

strSql = "select G_VENDOR_CODE from TBL_AH017_DN where D_LOT = '" & strLotID & "' "
strVendCode = Get_OracleStr(strSql)

strContent = strContent & ";" & Chr(34) & "HWPN" & Chr(34) & "," & Chr(34) & strHWPN & Chr(34) & ";"
strContent = strContent & ";" & Chr(34) & "VDCODE" & Chr(34) & "," & Chr(34) & strVendCode & Chr(34) & ";"

iPces = CInt(Trim(txtPecs.text))
strSql = "INSERT INTO erpdata..tblME_PrintInfo(PrinterNameID,BartenderName,Content,Content2,Content3,flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) " & _
"SELECT a.PrinterNameID,a.BartenderName,'" & strContent & "',a.Content2,a.Content3,'0' AS flag ,GETDATE() AS create_date,a.EVENT_SOURCE AS EVENT_SOURCE,a.EVENT_ID,a.LABEL_ID,a.PRINT_QTY FROM erpdata..tblME_PrintInfo a WHERE a.ID = '" & strid & "' "

AddSql2 (strSql)

PrintIBoxLbl = True

End Sub
