VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmUpEDC 
   Caption         =   "更新EDC进Mes的值"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "EDC_Excel"
      Height          =   2295
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   9015
      Begin VB.CommandButton Command5 
         Caption         =   "导出报表"
         Height          =   480
         Left            =   4080
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "上传DB"
         Height          =   480
         Left            =   1200
         TabIndex        =   3
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   ".."
         Height          =   495
         Left            =   6120
         TabIndex        =   2
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   495
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   4935
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3000
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         MaxFileSize     =   10000
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择待上传的Excel："
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   1635
      End
   End
End
Attribute VB_Name = "FrmUpEDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BcRS        As New ADODB.Recordset

Private Sub CmdDel_Click()
Dim idtemp As String

idtemp = Trim$(TxtBatchId.Text)

If idtemp = "" Then
    MsgBox "BatchId不可以为空"
    Exit Sub
    
End If

'判断输入的Lot号，是否存在于BC表中
If (Not JudgeBCExist(idtemp)) Then
   MsgBox "这笔：" & idtemp & " 不存在，无需删除！"
Exit Sub

End If


Call DelBC(idtemp)

End Sub

Private Sub Command1_Click()
'修改数量
Dim idtemp As String

idtemp = Trim$(TxtBatchId.Text)

If idtemp = "" Then
    MsgBox "BatchId不可以为空"
    Exit Sub
    
End If

If Trim(TxtQty1.Text) = "" Then
'先根据BatchId带出原来数量
    MsgBox "先输入BatchId，后回车，带出原来数量！"
    Exit Sub
End If

If Trim(TxtQty2.Text) = "" Then
    MsgBox "请输入现在BC中的数量！"
    Exit Sub
End If

'判断数量是否大于原来数量

If CLng(Trim(TxtQty2.Text)) > CLng(TxtQty1.Text) Then
    MsgBox "输入的数量不可以大于原来的数量！"
    Exit Sub
End If



Call ModifyBC(idtemp, CLng(Trim(TxtQty2.Text)))




End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim FName
    '帅选文件
    CommonDialog1.Filter = "EXCEL文件(*.xls)|*.xls"
    CommonDialog1.ShowOpen
    '得到文件名
    FName = CommonDialog1.FileName
    If FName <> "" Then
  
        Text2.Text = Replace(FName, Chr(160), ",")
        
        
    End If
End Sub

Private Sub Command3_Click()
    SumCount = 0
    ErrorInf = ""
    If Text2.Text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
    
    End If
    
    Dim FileName As String
    FileName = Text2.Text
    Dim dirtemp() As String
    Dim i As Integer
    
    If InStr(1, FileName, ",") > 0 Then
        dirtemp = Split(FileName, ",")
        
        For i = 1 To UBound(dirtemp)
            UpBI (dirtemp(0) + "\" + dirtemp(i))
        Next
        
    Else
        
        UpBI (FileName)
    End If
    
    
    If SumCount > 0 Then
        MsgBox "已成功上传" & SumCount & "笔！"
    End If
    
    If ErrorInf <> "" Then
           MsgBox "上传失败的有:" + ErrorInf + "数据库中已存在！"
    End If




'
'
'
'
'
'
'Dim source_batch_id_Temp As String
''上传BI的CSV
''处理文件名
'If Text2.Text = "" Then
'    MsgBox "先选择待上传的文件"
'    Exit Sub
'End If
'Dim dirName As String
'Dim FileName As String
'
''获取文件名
'    If InStrRev(Trim(Text2.Text), "\") > 0 Then
'        strfilename = Mid(Trim(Text2.Text), InStrRev(Trim(Text2.Text), "\") + 1)
'        dirName = Mid$(Trim(Text2.Text), 1, InStrRev(Trim(Text2.Text), "\"))
'    End If
'
'
'Dim con As New ADODB.Connection
'Dim Rs As New ADODB.Recordset
'
''con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
''Rs.open "Select * From " & strfilename, con, adOpenStatic, adLockReadOnly, adCmdText
'
''2012-07-03 jiayunzhang 修改读CSV的方式
'
'  '1)打开Excel
'
'    Set VBExcel = CreateObject("excel.application")     '创建Excle对象
'
'    VBExcel.Visible = False
'
'    Set xlBook = VBExcel.Workbooks.open(Text2.Text)    '打开文件
'
'    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表
'
'
'  '判定最大列Excel中的和设定列是否相同
'  '2012-10-08 jiayunzhang 市场部要求新增一列 comp_code
'
'    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 16 Then
'
'        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
'        Exit Sub
'
'    End If
'
'
'
'
'
'
'
'Dim i As Integer
'Dim j As Integer
'Dim id As Long
'Dim temp As String
'Dim temp2 As String
'Dim tempVal As String
'Dim WV_inspect As String
'Dim Comp_codeTemp As String
'
'
'
'Dim SumCount As Integer
'SumCount = 0
''Rs.MoveFirst
''For i = 0 To Rs.RecordCount - 1
'
' For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
'
'
'    temp = ""
'    source_batch_id_Temp = ""
'      For j = 1 To 16
'
'        strChar = Chr(96 + j)
'        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
'
'        If j = 1 Then
'            source_batch_id_Temp = Trim(tempVal)  'LotId
'        End If
'
'        If j = 15 Then
'        tempVal = Format(tempVal, "HH:MM:SS")
'
'        End If
'
'
'        temp = temp & "," & newStr("" & tempVal)
'
'    Next j
'
'    temp = temp & ",'','Y','" & gUserName & "',sysdate,'',''"
'
''    Debug.Print temp
'
''             '判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
'    If (JudgeFlagStautsBI(source_batch_id_Temp)) Then
'       MsgBox "这笔：" & source_batch_id_Temp & "已存在，无需上传!"
'       GoTo NextRecord2
'
'    End If
'
'    '处理前面的,
'
'    temp = Mid(temp, 2, Len(temp) - 1)
'
'
'    Call AddBI(temp)
'     SumCount = SumCount + 1
'
'    '上传到DB
'
'NextRecord2:
''    Rs.MoveNext
'
'Next i
'
'
'If SumCount > 0 Then
'    MsgBox "已成功上传" & SumCount & "笔！"
'End If

End Sub


Private Function newStr(temp As String)
If temp <> "" Then
newStr = "'" & temp & "'"
Else
newStr = "''"

End If

End Function

Private Sub Command5_Click()

Dim sqlTemp As String
 
 sqlTemp = "select  Batch_id ,Mtrl_Num ,Mtrl_Desc ,Design_Id ,Current_Wafer_Qty ,Die_Qty ,Fabrication_Facility ,Imager_Customer_Rev ," & _
           " Dark_Bond_Pads ,Ship_Comment ,Probe_Ship_Part_type ,Aptina_Doc_Number ,Awb ,Created_Date ,Created_Time ," & _
           " OffShore_Probe_Facility , Qtech_Created_by, Qtech_Created_date,(select b.createdate from   CustomerBCtbl b where b.batchid=a.batch_id and b.flag='Y') as BC_UploadDate " & _
           " from CustomerBItbl a where flag='Y' order by Qtech_Created_date desc "
               
  ExporToExcel sqlTemp
    

End Sub



Private Sub TxtBatchId_KeyPress(KeyAscii As Integer)
Dim idtemp As String

If KeyAscii = 13 Then
    idtemp = Trim$(TxtBatchId.Text)

    '判断输入的Lot号，是否存在于BC表中
    If (Not JudgeBCExist(idtemp)) Then
       MsgBox "这笔：" & idtemp & " 不存在，无需删除！"
    Exit Sub
    
    End If
    
    Set BcRS = GetDecBCQty(idtemp)

    TxtQty1.Text = BcRS.fields("dieqty").Value


End If

End Sub


Private Sub TxtQty2_KeyPress(KeyAscii As Integer)
Dim Numbers As String
Numbers = "1234567890" + Chr(8)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If


End Sub

Private Sub Form_Load()
CommonDialog1.Flags = &H80200



End Sub

Private Sub UpBI(dirtemp As String)


Dim TPriceFlag As Boolean


Dim source_batch_id_Temp As String
'上传BI的CSV
'处理文件名
If Text2.Text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

TPriceFlag = False


'获取文件名
    If InStrRev(Trim(dirtemp), "\") > 0 Then
        strFileName = Mid(Trim(dirtemp), InStrRev(Trim(dirtemp), "\") + 1)
        dirName = Mid$(Trim(dirtemp), 1, InStrRev(Trim(dirtemp), "\"))
        FileName = Replace(UCase(Mid(Trim(strFileName), InStrRev(Trim(strFileName), "_") + 1)), ".XLS", "")
    End If
    

Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset

'con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
'Rs.open "Select * From " & strfilename, con, adOpenStatic, adLockReadOnly, adCmdText

'2012-07-03 jiayunzhang 修改读CSV的方式

  '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(dirtemp)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表


  '判定最大列Excel中的和设定列是否相同
  '2012-10-08 jiayunzhang 市场部要求新增一列 comp_code

'    If xlSheet.Range("A1").CurrentRegion.Columns.Count = 17 Then
'       TPriceFlag = True
'
'    ElseIf xlSheet.Range("A1").CurrentRegion.Columns.Count = 16 Then
'        TPriceFlag = False
'
'    Else
'        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
'        Exit Sub
'
'    End If







Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
Dim WV_inspect As String
Dim Comp_codeTemp As String

Dim typeName1 As String
Dim valueTemp1 As Integer




'Dim SumCount As Integer
'SumCount = 0
'Rs.MoveFirst
'For i = 0 To Rs.RecordCount - 1

'xlSheet.Range("A1").CurrentRegion.Rows.Count

' For i = 2 To 48
'
'     If i = 43 Then
'
'     typeName1 = "槽上开口A"
'
'    strChar = Chr(96 + 3)
'    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值
'
'     valueTemp1 = CInt(Trim(tempVal))
'
'     Call UpdateEDCData(FileName, typeName1, valueTemp1)
'     SumCount = SumCount + 1
'
'
'     End If
'
'      If i = 45 Then
'
'     typeName1 = "槽上开口B"
'
'    strChar = Chr(96 + 3)
'    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值
'
'     valueTemp1 = CInt(Trim(tempVal))
'
'     Call UpdateEDCData(FileName, typeName1, valueTemp1)
'     SumCount = SumCount + 1
'
'
'     End If
'
'
'      If i = 47 Then
'
'     typeName1 = "槽上开口1"
'
'    strChar = Chr(96 + 3)
'    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值
'
'     valueTemp1 = CInt(Trim(tempVal))
'
'     Call UpdateEDCData(FileName, typeName1, valueTemp1)
'     SumCount = SumCount + 1
'
'
'     End If
'
'       If i = 44 Then
'
'     typeName1 = "槽下开口A"
'
'    strChar = Chr(96 + 3)
'    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值
'
'     valueTemp1 = CInt(Trim(tempVal))
'
'     Call UpdateEDCData(FileName, typeName1, valueTemp1)
'     SumCount = SumCount + 1
'
'
'     End If
'
'     If i = 46 Then
'
'     typeName1 = "槽下开口B"
'
'    strChar = Chr(96 + 3)
'    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值
'
'     valueTemp1 = CInt(Trim(tempVal))
'
'     Call UpdateEDCData(FileName, typeName1, valueTemp1)
'     SumCount = SumCount + 1
'
'
'     End If
'
'
'       If i = 48 Then
'
'     typeName1 = "槽下开口1"
'
'    strChar = Chr(96 + 3)
'    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值
'
'     valueTemp1 = CInt(Trim(tempVal))
'
'     Call UpdateEDCData(FileName, typeName1, valueTemp1)
'     SumCount = SumCount + 1
'
'
'     End If
'
'
'
'
'NextRecord2:
''    Rs.MoveNext
'
'Next i


' 38-44行

For i = 2 To 44

     If i = 39 Then

     typeName1 = "槽上开口A"

    strChar = Chr(96 + 3)
    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值

     valueTemp1 = CInt(Trim(tempVal))

     Call UpdateEDCData(FileName, typeName1, valueTemp1)
     SumCount = SumCount + 1


     End If

      If i = 41 Then

     typeName1 = "槽上开口B"

    strChar = Chr(96 + 3)
    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值

     valueTemp1 = CInt(Trim(tempVal))

     Call UpdateEDCData(FileName, typeName1, valueTemp1)
     SumCount = SumCount + 1


     End If


      If i = 43 Then

     typeName1 = "槽上开口1"

    strChar = Chr(96 + 3)
    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值

     valueTemp1 = CInt(Trim(tempVal))

     Call UpdateEDCData(FileName, typeName1, valueTemp1)
     SumCount = SumCount + 1


     End If

       If i = 40 Then

     typeName1 = "槽下开口A"

    strChar = Chr(96 + 3)
    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值

     valueTemp1 = CInt(Trim(tempVal))

     Call UpdateEDCData(FileName, typeName1, valueTemp1)
     SumCount = SumCount + 1


     End If

     If i = 42 Then

     typeName1 = "槽下开口B"

    strChar = Chr(96 + 3)
    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值

     valueTemp1 = CInt(Trim(tempVal))

     Call UpdateEDCData(FileName, typeName1, valueTemp1)
     SumCount = SumCount + 1


     End If


       If i = 44 Then

     typeName1 = "槽下开口1"

    strChar = Chr(96 + 3)
    tempVal = xlSheet.Range(strChar & i).Value    '临时保存值

     valueTemp1 = CInt(Trim(tempVal))

     Call UpdateEDCData(FileName, typeName1, valueTemp1)
     SumCount = SumCount + 1


     End If




NextRecord2:
'    Rs.MoveNext

Next i





  VBExcel.Quit
 Set VBExcel = Nothing
    
    


End Sub
