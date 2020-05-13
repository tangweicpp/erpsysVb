VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmUpLoadONBC 
   Caption         =   "ON 上传BC资料"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14145
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
   ScaleHeight     =   7500
   ScaleWidth      =   14145
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdClear 
      Caption         =   "清空 "
      Height          =   480
      Left            =   6960
      TabIndex        =   21
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存"
      Height          =   480
      Left            =   2880
      TabIndex        =   20
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox TxtPeace 
      Height          =   375
      Left            =   7200
      TabIndex        =   19
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox TxtCS 
      Height          =   375
      Left            =   10560
      TabIndex        =   17
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox TxtDie 
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox TxtDevice 
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox TxtInvoice 
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox TxtLotID 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "BC_xls"
      Height          =   2295
      Left            =   840
      TabIndex        =   0
      Top             =   2400
      Width           =   10815
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
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择待上传的xls："
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   10080
      TabIndex        =   10
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Format          =   180158465
      CurrentDate     =   40882
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "片数："
      Height          =   195
      Left            =   6600
      TabIndex        =   18
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"FormOn_BC.frx":0000
      Height          =   390
      Left            =   8640
      TabIndex        =   16
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Die数："
      Height          =   195
      Left            =   4680
      TabIndex        =   14
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DeviceID："
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   840
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"FormOn_BC.frx":001E
      Height          =   390
      Left            =   8880
      TabIndex        =   11
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "InvoiceID："
      Height          =   195
      Left            =   4440
      TabIndex        =   8
      Top             =   240
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LotID："
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   570
   End
End
Attribute VB_Name = "FrmUpLoadONBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BcRS        As New ADODB.Recordset

Private Sub CmdDel_Click()
Dim idTemp As String

idTemp = Trim$(TxtBatchId.Text)

If idTemp = "" Then
    MsgBox "BatchId不可以为空"
    Exit Sub
    
End If

'判断输入的Lot号，是否存在于BC表中
If (Not JudgeBCExist(idTemp)) Then
   MsgBox "这笔：" & idTemp & " 不存在，无需删除！"
Exit Sub

End If


Call DelBC(idTemp)

End Sub

Private Sub Command1_Click()
'修改数量
Dim idTemp As String

idTemp = Trim$(TxtBatchId.Text)

If idTemp = "" Then
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



Call ModifyBC(idTemp, CLng(Trim(TxtQty2.Text)))




End Sub

Private Sub CmdClear_Click()
TxtLotID.Text = ""
TxtInvoice.Text = ""
TxtDevice.Text = ""
TxtDie.Text = ""
TxtPeace.Text = ""
TxtCS.Text = ""



End Sub

Private Sub CmdSave_Click()

Dim cmdStr As String
Dim cmdStr2 As String

Dim idTemp As Long
Dim lotidtemp As String
Dim invoiceTemp As String
Dim transDateTemp As String
Dim diveictTemp As String
Dim dieQtyTemp As Long
Dim pieceQtyTemp As Integer
Dim csTemp As String

idTemp = GetMaxID()
lotidtemp = UCase(Trim(TxtLotID.Text))
invoiceTemp = UCase(Trim(TxtInvoice.Text))
transDateTemp = CStr(DTPicker1.Value)
diveictTemp = UCase(Trim(TxtDevice.Text))
If TxtDie.Text = "" Then
dieQtyTemp = 0
Else

dieQtyTemp = CLng(Trim(TxtDie.Text))
End If

pieceQtyTemp = CInt(Trim(TxtPeace.Text))
csTemp = UCase(Trim(TxtCS.Text))



'添加导入Sqlserver
On Error GoTo DealError
        
'Cnn.BeginTrans

cmdStr = " insert into CustomerBCtbl(ID ,BATCHID ,APTINADOCNUMBER  ,LOTRECDATE ,MTRLNUM ," & _
" DIEQTY,DESIGNID ,CURRENT_WAFER_QTY ,FLAG ,CREATEBY ,CreateDate) " & _
" values(" & idTemp & " ,'" & lotidtemp & "','" & invoiceTemp & "','" & transDateTemp & "','" & diveictTemp & "'," & _
" " & dieQtyTemp & ",'" & csTemp & "'," & pieceQtyTemp & ",'Y','Auto',sysdate) "

cmdStr2 = " insert into [ERPBASE].[dbo].[tblCustomerBC](ID ,BATCHID ,APTINADOCNUMBER  ,LOTRECDATE ,MTRLNUM ," & _
" DIEQTY,DESIGNID ,CURRENT_WAFER_QTY ,FLAG ,CREATEBY ,CreateDate) " & _
" values(" & idTemp & " ,'" & lotidtemp & "','" & invoiceTemp & "','" & transDateTemp & "','" & diveictTemp & "'," & _
" " & dieQtyTemp & ",'" & csTemp & "'," & pieceQtyTemp & ",'Y','Auto',getdate()) "
                        
AddSql (cmdStr)
AddSql2 (cmdStr2)
 
'Cnn.CommitTrans
 MsgBox "保存成功!", vbInformation, "友情提示"
  
Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True




End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim FName
    '帅选文件
    CommonDialog1.Filter = "EXCEL文件(*.xls)|*.xls|EXCEL文件(*.xlsx)|*.xlsx"
    

    CommonDialog1.ShowOpen
    '得到文件名
    FName = CommonDialog1.FileName
    If FName <> "" Then
       Text2.Text = FName
    End If
End Sub

Private Sub Command3_Click()
'上传资料

Dim source_batch_id_Temp As String
Dim lotTypeTemp As String

'上传OI的CSV
'处理文件名
If Text2.Text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

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

    Set xlBook = VBExcel.Workbooks.open(Text2.Text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.Count = 7 Or xlSheet.Range("A1").CurrentRegion.Columns.Count = 6 Then

       source_batch_id_Temp = ""

    Else
          MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
          Exit Sub
    
    End If


Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
Dim dieQtyTemp As Long
Dim pcsQtemp As Integer
Dim dateFormatTemp As String



   
lotTypeTemp = "P"

SumCount = 0
BCResultFlag = False

Cnn.BeginTrans

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    pcsQtemp = 0
    
    '判断 是不是 "Σ"
    
    strChar = Chr(96 + 1)
    tempVal = xlSheet.Range(strChar & i).Value
    
    If InStr(tempVal, "Σ") > 0 Then
    
      GoTo NextRecord2
    
    End If
    


    
    source_batch_id_Temp = ""
    
    '样品 后面加一列  手动加样品标志
    If xlSheet.Range("A1").CurrentRegion.Columns.Count = 7 Then
    
             For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
                 strChar = Chr(96 + j)
                 tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
                    
                 If j = 1 Then
                     source_batch_id_Temp = Right(UCase(Trim(tempVal)), Len(UCase(Trim(tempVal))) - 2)  'LotId
                     
                     temp = temp & "," & newStr("" & Right(UCase(Trim(tempVal)), Len(UCase(Trim(tempVal))) - 2))
                     
                 End If
                 
                 If j = 4 Then
                    ' dieQtyTemp = CLng(Trim(tempVal))  'qty
                      'temp = temp & "," & newStr("" & tempVal)
                      temp = temp & "," & newStr("" & tempVal)
                 End If
                 
                 If j = 3 Then
            
                     
                     dateFormatTemp = tempVal
                     
                     tempVal = newDateStr(dateFormatTemp)
                     
                     temp = temp & "," & newStr("" & tempVal)
                     
                                
                 End If
                  
                     
                 If j = 2 Then
                     temp = temp & "," & newStr("" & tempVal)
                   
                 End If
                 
                 If j = 5 Then
                 
                    pcsQtemp = CInt(Trim(tempVal))
                   
                 End If
                 
                 If j = 6 Then
                     temp = temp & "," & newStr("" & tempVal)
                   
                 End If
                 
                 If j = 7 Then
                     lotTypeTemp = "S"
                 
                 End If
                 
                 
             Next j
    
    
    Else
          '正常拷下来的
          
         For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
                 strChar = Chr(96 + j)
                 tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
                    
                 If j = 1 Then
                     source_batch_id_Temp = Right(UCase(Trim(tempVal)), Len(UCase(Trim(tempVal))) - 2)  'LotId
                     
                     temp = temp & "," & newStr("" & Right(UCase(Trim(tempVal)), Len(UCase(Trim(tempVal))) - 2))
                     
                 End If
                 
                 If j = 4 Then
                    ' dieQtyTemp = CLng(Trim(tempVal))  'qty
                      'temp = temp & "," & newStr("" & tempVal)
                      temp = temp & "," & newStr("" & tempVal)
                 End If
                 
                 If j = 3 Then
                    '6/30/15 12:00:00 AM
                    
                     dateFormatTemp = tempVal
                     
                     tempVal = newDateStr(dateFormatTemp)
                     
                     temp = temp & "," & newStr("" & tempVal)
                                
                 End If
                  
                     
                 If j = 2 Then
                     temp = temp & "," & newStr("" & tempVal)
                   
                 End If
                 
                 If j = 5 Then
                 
                 pcsQtemp = CInt(Trim(tempVal))
                 
'                   dieQtyTemp = CLng(Trim(tempVal))  'qty
'                      temp = temp & "," & newStr("" & tempVal)
                      
                   '  temp = temp & "," & newStr("" & tempVal)
                   
                 End If
                 
                 If j = 6 Then
                     temp = temp & "," & newStr("" & tempVal)
                   
                 End If
                 
'                 If j = 7 Then
'                 pcsQtemp = CInt(Trim(tempVal))
'
'                 End If
    
              Next j
    
    End If
    

    
    
    
    
    '取目前DB最大的ID号
    id = GetMaxID()
    temp = id & temp
'    temp2 = temp & ",'Y','Auto',GETDATE(),'','','AA',0"

    temp2 = temp & "," & pcsQtemp & ",'Y','Auto',GETDATE(),  '" & lotTypeTemp & "'"
    temp = temp & "," & pcsQtemp & ",'Y','Auto',sysdate,  '" & lotTypeTemp & "'"

'    Debug.Print temp

             '判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
    If (JudgeFlagStautsBC(source_batch_id_Temp)) Then
       MsgBox "这笔：" & source_batch_id_Temp & "已存在，无需上传!"
       GoTo NextRecord2

    End If
    
    
'    If (Not JudgeFlagStautsBCQty(source_batch_id_Temp, dieQtyTemp)) Then
'       MsgBox "这笔：" & source_batch_id_Temp & "与BI中的Die数量不一致!"
'       GoTo NextRecord2
'
'    End If


    Call AddONBC(temp, temp2)
    SumCount = SumCount + 1
     
    '上传到DB
NextRecord2:

Next i

Cnn.CommitTrans

     
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


Private Function newStr(temp As String)
If temp <> "" Then
newStr = "'" & temp & "'"
Else
newStr = "''"

End If

End Function

Private Function newDateStr(temp As String)
'6/30/15 12:00:00 AM

Dim str1 As String

Dim resultTemp As String
Dim monthTemp As String
Dim dayTemp As String
Dim yearTemp As String
Dim i As Integer

str1 = temp

i = InStr(str1, "/")
monthTemp = Left(str1, i - 1)

str1 = Right(str1, Len(str1) - i)

i = InStr(str1, "/")
dayTemp = Left(str1, i - 1)

str1 = Right(str1, Len(str1) - i)

i = InStr(str1, " ")
yearTemp = Left(str1, i - 1)

yearTemp = Right("20" & yearTemp, 4)

newDateStr = yearTemp & "-" & monthTemp & "-" & dayTemp


End Function


Private Sub Command5_Click()

 ExporToExcel (" select ID,BATCHID,MTRLNUM, MTRLDESC,DESIGNID,DIEQTY,APTINADOCNUMBER, LOTRECDATE ,CURRENT_WAFER_QTY from CustomerBCtbl order by id desc ")
      
      
End Sub


Private Sub TxtBatchId_KeyPress(KeyAscii As Integer)
Dim idTemp As String

If KeyAscii = 13 Then
    idTemp = Trim$(TxtBatchId.Text)

    '判断输入的Lot号，是否存在于BC表中
    If (Not JudgeBCExist(idTemp)) Then
       MsgBox "这笔：" & idTemp & " 不存在，无需删除！"
    Exit Sub
    
    End If
    
    Set BcRS = GetDecBCQty(idTemp)

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

Private Sub Form_Activate()
TxtLotID.SetFocus

DTPicker1.Value = Format(Now, "yyyy-mm-dd")

End Sub

