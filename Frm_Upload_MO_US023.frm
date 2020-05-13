VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Upload_MO_US023 
   Caption         =   "上传MO_US023"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16095
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
   ScaleHeight     =   10440
   ScaleWidth      =   16095
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   16095
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   14280
         Top             =   2760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Upload_MO_US023.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Upload_MO_US023.frx":0C52
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Upload_MO_US023.frx":18A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Upload_MO_US023.frx":24F6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1418
         Width           =   10935
      End
      Begin MSComDlg.CommonDialog com 
         Left            =   13560
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   600
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   1058
         ButtonWidth     =   2408
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "导入数据"
               Key             =   "IMPORT"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存数据"
               Key             =   "SAVE"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "导出数据"
               Key             =   "EXPORT"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出窗体"
               Key             =   "EXIT"
               ImageIndex      =   4
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   7815
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   12015
         _Version        =   524288
         _ExtentX        =   21193
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
         MaxCols         =   7
         MaxRows         =   0
         SpreadDesigner  =   "Frm_Upload_MO_US023.frx":2848
         TextTip         =   2
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上传文件:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   240
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   1035
      End
   End
End
Attribute VB_Name = "Frm_Upload_MO_US023"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    With fps(0)
        .DAutoSizeCols = DAutoSizeColsBest
        
        .Col = -1
        .Row = -1
        .Lock = True
        
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        .ColWidth(1) = 12
        .ColWidth(2) = 12
        .ColWidth(3) = 16
        .ColWidth(4) = 6
        .ColWidth(5) = 16
        .ColWidth(6) = 16
        .ColWidth(7) = 16
    End With

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key

        Case "IMPORT"
            importData
    
        Case "EXPORT"
            exportData
            
        Case "SAVE"
            saveData
            
        Case "EXIT"
            Unload Me

    End Select

End Sub

Private Sub exportData()

Dim strsql As String
strsql = "select * from TBL_MO_US023 where create_date > sysdate - 30 order by create_date desc"
ExporToExcel (strsql)
End Sub

Private Sub importData()

 On Error GoTo ErrHandler

    Dim FName
    
    com.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
    com.ShowOpen

    FName = com.filename

    If FName <> "" Then
        txtPath.Text = FName
        showData_upload

    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "错误产生", vbInformation, "提示"
    Exit Sub

End Sub

Private Sub showData_upload()
 On Error GoTo ErrHandle

    Dim VBExcel     As Excel.Application

    Dim xlBook      As Excel.Workbook

    Dim xlSheet     As Excel.Worksheet

    Dim strFilename As String

    Dim I           As Integer

    Dim J           As Integer

    Dim strChar     As String

    Dim strTmp  As String
    
    MousePointer = 11

    fps(0).MaxRows = 0

    If InStrRev(Trim(txtPath.Text), "\") > 0 Then
        strFilename = Mid(Trim(txtPath.Text), InStrRev(Trim(txtPath.Text), "\") + 1)

        If InStr(strFilename, ".") > 0 Then
            strFilename = Mid(strFilename, 1, InStr(strFilename, ".") - 1)

        End If

    End If

    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(txtPath.Text)
    Set xlSheet = xlBook.Worksheets(1)
  
    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 7 Then
        MousePointer = 0
        MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        GoTo EXITPRO
        Exit Sub

    End If

    With fps(0)

        For I = 1 To xlSheet.Range("A1").CurrentRegion.Rows.Count
            strTmp = Trim(xlSheet.Range("A" & I).Value)

            If Len(strTmp) > 0 Then
                If I <> 1 Then .MaxRows = .MaxRows + 1

                For J = 1 To 7

                    If J > 26 Then
                        strChar = Chr(96 + Int(J / 26 - 0.001)) & IIf(J Mod 26 = 0, "Z", Chr(96 + (J Mod 26)))
                    Else
                        strChar = Chr(96 + J)

                    End If

                    If I = 1 Then
                        .SetText J, .MaxRows, Trim$(xlSheet.Range(strChar & I))
                    Else
                        .SetText J, .MaxRows, Trim$(xlSheet.Range(strChar & I))

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
    
    Exit Sub
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

    Exit Sub
ErrHandle:
    GoTo EXITPRO
End Sub

Private Sub saveData()

    On Error GoTo ErrHandle

    Dim strsql              As String

    Dim I                   As Integer, J As Integer
    
    Dim strMoReceiptDate    As String

    Dim strWaferReceiptDate As String

    Dim strWaferID          As String

    Dim strFabLotNo         As String

    Dim strAsmLotNo         As String

    Dim strProduct          As String
    
    Dim strArr()            As String
    
    If fps(0).MaxRows <= 0 Then
        MsgBox "没有要保存的资料", vbInformation, "提示"
        Exit Sub

    End If
    
    If MsgBox("是否要保存吗？", vbInformation + vbYesNo, "提示") = vbNo Then Exit Sub
    MousePointer = 11

    With fps(0)

        For I = 1 To .MaxRows
            .Row = I
          
            .Col = 1
            strMoReceiptDate = Trim$(.Text)
            
            .Col = 2
            strWaferReceiptDate = Trim$(.Text)
            
            .Col = 3
            strWaferID = Trim$(.Text)
                
            .Col = 5
            strFabLotNo = Replace(Trim$(.Text), " ", "")
                
            .Col = 6
            strAsmLotNo = Trim$(.Text)
            
            .Col = 7
            strProduct = Trim$(.Text)
            
            If InStr(strWaferID, ",") > 0 Then
                strArr = Split(Replace$(strWaferID, "#", ""), ",")
                
                For J = 0 To UBound(strArr)
                    
                    strWaferID = strArr(J)
                    
                    strWaferID = Split(strFabLotNo, "(")(0) & Right$("0" & strWaferID, 2)
                    
                    strsql = "select * from TBL_MO_US023 where wafer_id = '" & strWaferID & "'"

                    If Get_OracleCnt(strsql) > 0 Then
                        MsgBox "WaferID: " & strWaferID & vbCrLf & "已经存在MO数据, 请勿重复上传", vbInformation, "提示"
                    Else
                        strsql = "insert into TBL_MO_US023(MO_RECEIPT_DATE,WAFER_RECEIPT_DATE,WAFER_ID,FAB_LOTNO,ASM_LOTNO,PRODUCT,CREATE_DATE,CREATE_BY) values('" & strMoReceiptDate & "','" & strWaferReceiptDate & "','" & strWaferID & "','" & strFabLotNo & "','" & strAsmLotNo & "','" & strProduct & "',sysdate, '" & gUserName & "')"
                        AddSql (strsql)
                    
                    End If
                    
                Next
                
            Else
                strWaferID = Replace$(strWaferID, "#", "")
                
                strWaferID = Split(strFabLotNo, "(")(0) & Right$("0" & strWaferID, 2)
                
                strsql = "select * from TBL_MO_US023 where wafer_id = '" & strWaferID & "'"

                If Get_OracleCnt(strsql) > 0 Then
                    MsgBox "WaferID" & strWaferID & "已经存在MO数据, 请勿重复上传", vbInformation, "提示"
                Else
                    strsql = "insert into TBL_MO_US023(MO_RECEIPT_DATE,WAFER_RECEIPT_DATE,WAFER_ID,FAB_LOTNO,ASM_LOTNO,PRODUCT,CREATE_DATE,CREATE_BY) values('" & strMoReceiptDate & "','" & strWaferReceiptDate & "','" & strWaferID & "','" & strFabLotNo & "','" & strAsmLotNo & "','" & strProduct & "',sysdate, '" & gUserName & "')"
                    AddSql (strsql)
                    
                End If

            End If
                
        Next

    End With

    MousePointer = 0
    
    MsgBox "资料保存成功！", vbInformation, Me.Caption
    
    Exit Sub
    
ErrHandle:
    MousePointer = 0
    MsgBox Err.Description, vbCritical + vbInformation, "警告"

End Sub
