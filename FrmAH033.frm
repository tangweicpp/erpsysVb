VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmAH033 
   Caption         =   "AH033 SN导入/PK导出"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17070
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
   ScaleHeight     =   9645
   ScaleWidth      =   17070
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "v"
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   17055
      Begin VB.CommandButton btnExit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "退 出"
         Height          =   285
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   945
         Width           =   975
      End
      Begin VB.CommandButton btnOutput 
         BackColor       =   &H00C0C0C0&
         Caption         =   "导 出"
         Height          =   285
         Left            =   5610
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   945
         Width           =   975
      End
      Begin VB.CommandButton btnDelete 
         BackColor       =   &H00C0C0C0&
         Caption         =   "删 除"
         Height          =   285
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   945
         Width           =   975
      End
      Begin VB.CommandButton btnQuery 
         BackColor       =   &H00C0C0C0&
         Caption         =   "查 询"
         Height          =   285
         Left            =   3390
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   945
         Width           =   975
      End
      Begin VB.TextBox txtInvoiceNo 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   945
         Width           =   1935
      End
      Begin VB.CommandButton btnSave 
         BackColor       =   &H00C0C0C0&
         Caption         =   "保 存"
         Height          =   285
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   13800
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton btnOpen 
         BackColor       =   &H00FFC0FF&
         Caption         =   "..."
         Height          =   285
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   8415
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   7455
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   16815
         _Version        =   524288
         _ExtentX        =   29660
         _ExtentY        =   13150
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
         SpreadDesigner  =   "FrmAH033.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE NO:"
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
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S N:"
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
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmAH033"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type AH033_SN
A_INVOICE_NO As String
B_SHIP_DATE As String
C_SHIP_NO As String
D_ITAM As String
E_INNER_PRODUCT_NO As String
F_OUTER_PRODUCT_NO As String
G_CODE As String
H_SHIP_QTY As String
I_CUST_NAME As String
J_SHIP_TO_CN As String
K_SHIP_TO_EN As String
L_ADDRESS_CN As String
M_ADDRESS_EN As String
N_SHIPPING_MARK As String
O_RECEIVER As String
P_TELEPHONE As String
Q_TELEPHONE_2 As String
R_PURCHASE As String
S_PACKING_CODE As String
T_PO_NO As String
U_CUST_PART_1 As String
V_CUST_PART_2 As String
W_PRODUCT_NAME As String
X_CUSTOMS_NO As String
Y_SECOND_PART As String
Z_PRODUCT_DESC As String
AA_COG_LOT As String
AB_SPLIT As String
AC_REMARK As String
AD_MODIFY_REASON As String
AE_SPECIFIED_NO As String
AF_SPECIFIED_SHIP_TO As String
AG_COF_COUNT_CTR As String
AH_LIMIT_LOTID As String
AI_BILL_TO As String
AJ_SHIP_TO_CODE As String
AK_TERMINAL_CUSTOMER As String

End Type

Private Sub btnDelete_Click()

If txtInvoiceNo.text = "" Then
    MsgBox "请输入要删除的InvoiceNo", vbInformation, "提示"
    Exit Sub
End If

If MsgBox("是否确定删除", vbYesNo, "提示") = vbNo Then
    Exit Sub
End If

AddSql ("delete from TBL_AH033_SN where invoice_no = '" & Trim$(txtInvoiceNo.text) & "'")
MsgBox "删除成功", vbInformation, "提示"

End Sub

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnOpen_Click()

With CommonDialog1
    .Filter = "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
    .ShowOpen
    txtFileName.text = CommonDialog1.filename
End With


End Sub

Private Sub btnOutput_Click()
Call ExportDataByInvoiceNO(Trim(txtInvoiceNo.text))
End Sub

Private Sub btnQuery_Click()

Call ShowDataByInvoiceNO(Trim(txtInvoiceNo.text))

End Sub

Private Sub btnSave_Click()
If txtFileName.text = "" Then
    MsgBox "请点击浏览打开需要上传的SN模板文件", vbInformation, "提示"
    Exit Sub
    
End If

If MsgBox("是否确定要上传", vbYesNo, "提示") = vbNo Then
    Exit Sub
End If

Call SaveFile(Trim$(txtFileName.text))

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
If lColsCnt <> 37 Then
    MsgBox "Excel中的列数:" & lColsCnt & "和设定的模版列数: 37 不一致" & vbCrLf & "请确认Excel是否正确！", vbInformation, "提示"
    GoTo hErr

End If

If lRowsCnt < 2 Then
    MsgBox "Excel中的行数至少2行", vbInformation, "提示"
    GoTo hErr
End If

ReDim dataArray(lRowsCnt - 1)

For i = 2 To lRowsCnt

    With dataArray(i - 2)
        .A_INVOICE_NO = Trim("" & Replace(Replace(xlSheet.Range("A" & i), Chr(10), ""), Chr(13), ""))
        .B_SHIP_DATE = Trim("" & Replace(Replace(xlSheet.Range("B" & i), Chr(10), ""), Chr(13), ""))
        .C_SHIP_NO = Trim("" & Replace(Replace(xlSheet.Range("C" & i), Chr(10), ""), Chr(13), ""))
        .D_ITAM = Trim("" & Replace(Replace(xlSheet.Range("D" & i), Chr(10), ""), Chr(13), ""))
        .E_INNER_PRODUCT_NO = Trim("" & Replace(Replace(xlSheet.Range("E" & i), Chr(10), ""), Chr(13), ""))
        .F_OUTER_PRODUCT_NO = Trim("" & Replace(Replace(xlSheet.Range("F" & i), Chr(10), ""), Chr(13), ""))
        .G_CODE = Trim("" & Replace(Replace(xlSheet.Range("G" & i), Chr(10), ""), Chr(13), ""))
        .H_SHIP_QTY = Trim("" & Replace(Replace(xlSheet.Range("H" & i), Chr(10), ""), Chr(13), ""))
        .I_CUST_NAME = Trim("" & Replace(Replace(xlSheet.Range("I" & i), Chr(10), ""), Chr(13), ""))
        .J_SHIP_TO_CN = Trim("" & Replace(Replace(xlSheet.Range("J" & i), Chr(10), ""), Chr(13), ""))
        .K_SHIP_TO_EN = Trim("" & Replace(Replace(xlSheet.Range("K" & i), Chr(10), ""), Chr(13), ""))
        .L_ADDRESS_CN = Trim("" & Replace(Replace(xlSheet.Range("L" & i), Chr(10), ""), Chr(13), ""))
        .M_ADDRESS_EN = Trim("" & Replace(Replace(xlSheet.Range("M" & i), Chr(10), ""), Chr(13), ""))
        .N_SHIPPING_MARK = Trim("" & Replace(Replace(xlSheet.Range("N" & i), Chr(10), ""), Chr(13), ""))
        .O_RECEIVER = Trim("" & Replace(Replace(xlSheet.Range("O" & i), Chr(10), ""), Chr(13), ""))
        .P_TELEPHONE = Trim("" & Replace(Replace(xlSheet.Range("P" & i), Chr(10), ""), Chr(13), ""))
        .Q_TELEPHONE_2 = Trim("" & Replace(Replace(xlSheet.Range("Q" & i), Chr(10), ""), Chr(13), ""))
        .R_PURCHASE = Trim("" & Replace(Replace(xlSheet.Range("R" & i), Chr(10), ""), Chr(13), ""))
        .S_PACKING_CODE = Trim("" & Replace(Replace(xlSheet.Range("S" & i), Chr(10), ""), Chr(13), ""))
        .T_PO_NO = Trim("" & Replace(Replace(xlSheet.Range("T" & i), Chr(10), ""), Chr(13), ""))
        .U_CUST_PART_1 = Trim("" & Replace(Replace(xlSheet.Range("U" & i), Chr(10), ""), Chr(13), ""))
        .V_CUST_PART_2 = Trim("" & Replace(Replace(xlSheet.Range("V" & i), Chr(10), ""), Chr(13), ""))
        .W_PRODUCT_NAME = Trim("" & Replace(Replace(xlSheet.Range("W" & i), Chr(10), ""), Chr(13), ""))
        .X_CUSTOMS_NO = Trim("" & Replace(Replace(xlSheet.Range("X" & i), Chr(10), ""), Chr(13), ""))
        .Y_SECOND_PART = Trim("" & Replace(Replace(xlSheet.Range("Y" & i), Chr(10), ""), Chr(13), ""))
        .Z_PRODUCT_DESC = Trim("" & Replace(Replace(xlSheet.Range("Z" & i), Chr(10), ""), Chr(13), ""))
        .AA_COG_LOT = Trim("" & Replace(Replace(xlSheet.Range("AA" & i), Chr(10), ""), Chr(13), ""))
        .AB_SPLIT = Trim("" & Replace(Replace(xlSheet.Range("AB" & i), Chr(10), ""), Chr(13), ""))
        .AC_REMARK = Trim("" & Replace(Replace(xlSheet.Range("AC" & i), Chr(10), ""), Chr(13), ""))
        .AD_MODIFY_REASON = Trim("" & Replace(Replace(xlSheet.Range("AD" & i), Chr(10), ""), Chr(13), ""))
        .AE_SPECIFIED_NO = Trim("" & Replace(Replace(xlSheet.Range("AE" & i), Chr(10), ""), Chr(13), ""))
        .AF_SPECIFIED_SHIP_TO = Trim("" & Replace(Replace(xlSheet.Range("AF" & i), Chr(10), ""), Chr(13), ""))
        .AG_COF_COUNT_CTR = Trim("" & Replace(Replace(xlSheet.Range("AG" & i), Chr(10), ""), Chr(13), ""))
        .AH_LIMIT_LOTID = Trim("" & Replace(Replace(xlSheet.Range("AH" & i), Chr(10), ""), Chr(13), ""))
        .AI_BILL_TO = Trim("" & Replace(Replace(xlSheet.Range("AI" & i), Chr(10), ""), Chr(13), ""))
        .AJ_SHIP_TO_CODE = Trim("" & Replace(Replace(xlSheet.Range("AJ" & i), Chr(10), ""), Chr(13), ""))
        .AK_TERMINAL_CUSTOMER = Trim("" & Replace(Replace(xlSheet.Range("AK" & i), Chr(10), ""), Chr(13), ""))
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
        strSql = "insert into TBL_AH033_SN(invoice_no,ship_date,ship_no,itam,inner_product_no,outer_product_no,code,ship_qty,cust_name,ship_to_cn,ship_to_en,address_cn,address_en,shipping_mark,receiver,telephone,telephone_2,purchase,packing_code,po_no,cust_part_1,cust_part_2,product_name,customs_no,second_part,product_desc,cog_lot,split,remark,modify_reason,specified_no,specified_ship_to,cof_count_ctr,limit_lotid,bill_to,ship_to_code,terminal_customer,UPLOAD_BY,UPLOAD_DATE,UPLOAD_ID) " & _
           " values('" & .A_INVOICE_NO & "','" & .B_SHIP_DATE & "','" & .C_SHIP_NO & "','" & .D_ITAM & "','" & .E_INNER_PRODUCT_NO & "','" & .F_OUTER_PRODUCT_NO & "','" & .G_CODE & "','" & .H_SHIP_QTY & "','" & .I_CUST_NAME & "','" & .J_SHIP_TO_CN & "','" & .K_SHIP_TO_EN & "','" & .L_ADDRESS_CN & "','" & .M_ADDRESS_EN & "','" & .N_SHIPPING_MARK & "','" & .O_RECEIVER & "','" & .P_TELEPHONE & "','" & .Q_TELEPHONE_2 & "','" & .R_PURCHASE & "','" & .S_PACKING_CODE & "','" & .T_PO_NO & "', " & _
           " '" & .U_CUST_PART_1 & "','" & .V_CUST_PART_2 & "','" & .W_PRODUCT_NAME & "','" & .X_CUSTOMS_NO & "','" & .Y_SECOND_PART & "','" & .Z_PRODUCT_DESC & "','" & .AA_COG_LOT & "','" & .AB_SPLIT & "','" & .AC_REMARK & "','" & .AD_MODIFY_REASON & "','" & .AE_SPECIFIED_NO & "','" & .AF_SPECIFIED_SHIP_TO & "','" & .AG_COF_COUNT_CTR & "','" & .AH_LIMIT_LOTID & "','" & .AI_BILL_TO & "','" & .AJ_SHIP_TO_CODE & "','" & .AK_TERMINAL_CUSTOMER & "','" & gUserRealName & "',sysdate,'" & lUploadID & "')"
        
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

strSql = "select * from TBL_AH033_SN where UPLOAD_ID = " & lUploadID & " "

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
    strSql = "select * from TBL_AH033_SN order by upload_date desc"
Else
    strSql = "select * from TBL_AH033_SN where invoice_no = '" & strInvoice & "'"
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
    strSql = "select * from TBL_AH033_SN order by upload_date desc"
Else
    strSql = "select * from TBL_AH033_SN where invoice_no = '" & strInvoice & "'"
End If

Call ExporToExcel(strSql)
End Sub

