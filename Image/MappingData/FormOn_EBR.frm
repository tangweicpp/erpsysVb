VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FormOn_EBR 
   Caption         =   "ON EBR��Ϣά��"
   ClientHeight    =   9105
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
   ScaleHeight     =   9105
   ScaleWidth      =   14145
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtAttr2 
      Height          =   375
      Left            =   9480
      TabIndex        =   12
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox TxtAttr1 
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox TxtPT 
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "��� "
      Height          =   480
      Left            =   6960
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "����"
      Height          =   480
      Left            =   2880
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox TxtContact 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox TxtERB 
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox TxtLotID 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   5895
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   13815
      _Version        =   196613
      _ExtentX        =   24368
      _ExtentY        =   10398
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
      SpreadDesigner  =   "FormOn_EBR.frx":0000
      TextTip         =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ATTR2��"
      Height          =   195
      Left            =   8760
      TabIndex        =   13
      Top             =   1320
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ATTR1��"
      Height          =   195
      Left            =   4920
      TabIndex        =   11
      Top             =   1320
      Width           =   660
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact��"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   750
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part��"
      Height          =   195
      Left            =   8880
      TabIndex        =   4
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EBR Number��"
      Height          =   195
      Left            =   4560
      TabIndex        =   2
      Top             =   480
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LotID��"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   570
   End
End
Attribute VB_Name = "FormOn_EBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BcRS        As New ADODB.Recordset

Private Enum E_FPS0          'Detail�֭��
    E_SeqId = 1                '���
    E_LotID                   'LotID
    E_ERB                     'EBR
    E_PT                      'PT
    E_Contact                'Contact
  
    
    E_End
    
End Enum


Private Sub CmdDel_Click()
Dim idTemp As String

idTemp = Trim$(TxtBatchId.Text)

If idTemp = "" Then
    MsgBox "BatchId������Ϊ��"
    Exit Sub
    
End If

'�ж������Lot�ţ��Ƿ������BC����
If (Not JudgeBCExist(idTemp)) Then
   MsgBox "��ʣ�" & idTemp & " �����ڣ�����ɾ����"
Exit Sub

End If


Call DelBC(idTemp)

End Sub

Private Sub Command1_Click()
'�޸�����
Dim idTemp As String

idTemp = Trim$(TxtBatchId.Text)

If idTemp = "" Then
    MsgBox "BatchId������Ϊ��"
    Exit Sub
    
End If

If Trim(TxtQty1.Text) = "" Then
'�ȸ���BatchId����ԭ������
    MsgBox "������BatchId����س�������ԭ��������"
    Exit Sub
End If

If Trim(TxtQty2.Text) = "" Then
    MsgBox "����������BC�е�������"
    Exit Sub
End If

'�ж������Ƿ����ԭ������

If CLng(Trim(TxtQty2.Text)) > CLng(TxtQty1.Text) Then
    MsgBox "��������������Դ���ԭ����������"
    Exit Sub
End If



Call ModifyBC(idTemp, CLng(Trim(TxtQty2.Text)))




End Sub

Private Sub CmdClear_Click()
TxtLotId.Text = ""
TxtERB.Text = ""
TxtPT.Text = ""
TxtContact.Text = ""
TxtAttr1.Text = ""
TxtAttr1.Text = ""



End Sub

Private Sub CmdSave_Click()

Dim cmdStr As String


Dim idTemp As Long
Dim lotidtemp As String
Dim erbTemp As String
Dim ptTemp As String

Dim contactTemp As String
Dim attr1Temp As String
Dim attr2Temp As String

idTemp = GetMaxID()

lotidtemp = UCase(Trim(TxtLotId.Text))
erbTemp = UCase(Trim(TxtERB.Text))
ptTemp = UCase(Trim(TxtPT.Text))
contactTemp = Trim(TxtContact.Text)
attr1Temp = Trim(TxtAttr1.Text)
attr2Temp = Trim(TxtAttr2.Text)


'��ӵ���Sqlserver
On Error GoTo DealError
        
'Cnn.BeginTrans

cmdStr = " insert into CUSTOMEREBRtbl(ID,BATCHID ,EBRNumber,PT,Contact,Attr1,Attr2 ,FLAG ,CREATEBY , CREATEDATE) " & _
" values(" & idTemp & " ,'" & lotidtemp & "','" & erbTemp & "','" & ptTemp & "','" & contactTemp & "'," & _
" '" & attr1Temp & "','" & attr2Temp & "','Y','Auto',sysdate) "

AddSql (cmdStr)

'Cnn.CommitTrans
 MsgBox "����ɹ�!", vbInformation, "������ʾ"
 
 ShowData_Where
  
Exit Sub
DealError:

Cnn.RollbackTrans
SumCount = SumCount - 1
BCResultFlag = True




End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim FName
    '˧ѡ�ļ�
    CommonDialog1.Filter = "EXCEL�ļ�(*.xls)|*.xls|EXCEL�ļ�(*.xlsx)|*.xlsx"
    

    CommonDialog1.ShowOpen
    '�õ��ļ���
    FName = CommonDialog1.FileName
    If FName <> "" Then
       Text2.Text = FName
    End If
End Sub

Private Sub Command3_Click()
'�ϴ�����

Dim source_batch_id_Temp As String
Dim lotTypeTemp As String

'�ϴ�OI��CSV
'�����ļ���
If Text2.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

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

    Set xlBook = VBExcel.Workbooks.open(Text2.Text)    '���ļ�

    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�

    '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count = 7 Or xlSheet.Range("A1").CurrentRegion.Columns.Count = 6 Then

       source_batch_id_Temp = ""

    Else
          MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
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
    
    '�ж� �ǲ��� "��"
    
    strChar = Chr(96 + 1)
    tempVal = xlSheet.Range(strChar & i).Value
    
    If InStr(tempVal, "��") > 0 Then
    
      GoTo NextRecord2
    
    End If
    


    
    source_batch_id_Temp = ""
    
    '��Ʒ �����һ��  �ֶ�����Ʒ��־
    If xlSheet.Range("A1").CurrentRegion.Columns.Count = 7 Then
    
             For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
                 strChar = Chr(96 + j)
                 tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
                    
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
          '������������
          
         For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
                 strChar = Chr(96 + j)
                 tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
                    
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
    

    
    
    
    
    'ȡĿǰDB����ID��
    id = GetMaxID()
    temp = id & temp
'    temp2 = temp & ",'Y','Auto',GETDATE(),'','','AA',0"

    temp2 = temp & "," & pcsQtemp & ",'Y','Auto',GETDATE(),  '" & lotTypeTemp & "'"
    temp = temp & "," & pcsQtemp & ",'Y','Auto',sysdate,  '" & lotTypeTemp & "'"

'    Debug.Print temp

             '�ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
    If (JudgeFlagStautsBC(source_batch_id_Temp)) Then
       MsgBox "��ʣ�" & source_batch_id_Temp & "�Ѵ��ڣ������ϴ�!"
       GoTo NextRecord2

    End If
    
    
'    If (Not JudgeFlagStautsBCQty(source_batch_id_Temp, dieQtyTemp)) Then
'       MsgBox "��ʣ�" & source_batch_id_Temp & "��BI�е�Die������һ��!"
'       GoTo NextRecord2
'
'    End If


    Call AddONBC(temp, temp2)
    SumCount = SumCount + 1
     
    '�ϴ���DB
NextRecord2:

Next i

Cnn.CommitTrans

     
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

    '�ж������Lot�ţ��Ƿ������BC����
    If (Not JudgeBCExist(idTemp)) Then
       MsgBox "��ʣ�" & idTemp & " �����ڣ�����ɾ����"
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



Private Sub Text1_Change()

End Sub

Private Sub Form_Load()

 With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        '�]�m�榡
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .SetText E_FPS0.E_SeqId, 0, "��¼��"
        .SetText E_FPS0.E_LotID, 0, "LotID"
        .SetText E_FPS0.E_ERB, 0, "EBR Number"
        .SetText E_FPS0.E_PT, 0, "PT"
        .SetText E_FPS0.E_Contact, 0, "Contact "
        
        
        .ColWidth(E_FPS0.E_SeqId) = 5
        .ColWidth(E_FPS0.E_LotID) = 10
        .ColWidth(E_FPS0.E_ERB) = 12
        .ColWidth(E_FPS0.E_PT) = 25
        .ColWidth(E_FPS0.E_Contact) = 25
       
        
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
ShowData_Where


End Sub


Private Sub ShowData_Where()
Set reportRS = GetEBRData()

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub
