VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_UploadShipSide 
   Caption         =   "上传出货地址"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12060
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
   ScaleHeight     =   7665
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "上传"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6360
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H000000FF&
         Caption         =   ".."
         Height          =   405
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FF80FF&
         Caption         =   "导出报表"
         Height          =   840
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtText1 
         Height          =   405
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   4215
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00FFFF00&
         Caption         =   "上传"
         Height          =   840
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径:"
         Height          =   195
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   420
      End
   End
End
Attribute VB_Name = "Frm_UploadShipSide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ShipSideTmp As ShipSideData

Private Sub cmd_Click()

On Error Resume Next
Dim FName
    '帅选文件
    CommonDialog1.Filter = "EXCEL文件(*.xlsx)|*.xlsx"
    
    CommonDialog1.ShowOpen
    '得到文件名
    FName = CommonDialog1.filename
    If FName <> "" Then
       txtText1.Text = FName
    End If
    
End Sub

Private Sub cmdOpen_Click()

Dim source_batch_id_Temp As String
Dim dirName As String
Dim filename As String

If txtText1.Text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
End If


'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(txtText1.Text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 5 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        Exit Sub

    End If


Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String



SumCount = 0
BCResultFlag = False

 ShipSideTmp.Created_ByTemp = gUserName

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count

    temp = ""
    source_batch_id_Temp = ""
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值

        If j = 1 Then

            ShipSideTmp.CustomerCode = Trim(tempVal)

        ElseIf j = 2 Then
            ShipSideTmp.GULFDeviceName = Trim(tempVal)

        ElseIf j = 3 Then
            ShipSideTmp.GULFLotID = Trim(tempVal)

        ElseIf j = 4 Then
            ShipSideTmp.WaferQTY = Trim(tempVal)

        ElseIf j = 5 Then
            ShipSideTmp.ShipTo = Trim(tempVal)

        End If

    Next j

    If (JudgeShipSideData(ShipSideTmp.GULFLotID)) Then
       MsgBox "这笔已存在，无需上传!", vbInformation, "友情提示"
       GoTo NextRecord2

    End If


    Call AddShipSideData(ShipSideTmp)
    SumCount = SumCount + 1

    '上传到DB
NextRecord2:

Next i



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



Private Sub CmdSave_Click()

SqlServer2ExporToExcel ("SELECT ID, CustCode, DeviceName, LotID, WaferQty, ShipTo, Memo " & _
"FROM tblSale_Shipto ")

End Sub
