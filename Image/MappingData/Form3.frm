VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmUpLoadBC 
   Caption         =   "�ϴ�BC����"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
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
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "���˴���"
      Height          =   2775
      Left            =   960
      TabIndex        =   6
      Top             =   2760
      Width           =   9015
      Begin VB.CommandButton Command1 
         Caption         =   "�޸�����"
         Height          =   480
         Left            =   4080
         TabIndex        =   14
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox TxtQty2 
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox TxtQty1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "ɾ���� Lot"
         Height          =   480
         Left            =   1200
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox TxtBatchId 
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Ϊ���˺󣬿�滹ʣ�µ�����"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6000
         TabIndex        =   15
         Top             =   1440
         Width           =   2475
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   195
         Left            =   3240
         TabIndex        =   13
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ������"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label LblBatchId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BatchId��"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "BC_xls"
      Height          =   2295
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   9015
      Begin VB.CommandButton Command5 
         Caption         =   "��������"
         Height          =   480
         Left            =   4080
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�ϴ�DB"
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
         Caption         =   "ѡ����ϴ���xls��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmUpLoadBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BcRS        As New ADODB.Recordset

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

Private Sub Command2_Click()
On Error Resume Next
Dim FName
    '˧ѡ�ļ�
    CommonDialog1.Filter = "EXCEL�ļ�(*.xls)|*.xls"
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

    Set xlSheet = xlBook.Worksheets("sheet1")        '��sheet�еı�

    '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 7 Then

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


   


SumCount = 0
BCResultFlag = False

Cnn.BeginTrans

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
    temp = ""
    pcsQtemp = 0
    
    source_batch_id_Temp = ""
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
           
        If j = 1 Then
            source_batch_id_Temp = Trim(tempVal)  'LotId
            
            temp = temp & "," & newStr("" & tempVal)
            
        End If
        
        If j = 4 Then
            dieQtyTemp = CLng(Trim(tempVal))  'qty
             temp = temp & "," & newStr("" & tempVal)
        End If
        
        If j = 3 Then
   
            
            temp = temp & "," & newStr("" & tempVal) & "," & newStr("" & Left(Trim(tempVal), 4))
                       
        End If
         
            
        If j = 2 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 5 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 6 Then
            temp = temp & "," & newStr("" & tempVal)
          
        End If
        
        If j = 7 Then
        pcsQtemp = CInt(Trim(tempVal))
        
        End If
        
        
    Next j

    'ȡĿǰDB����ID��
    id = GetMaxID()
    temp = id & temp
'    temp2 = temp & ",'Y','Auto',GETDATE(),'','','AA',0"

    temp2 = temp & ",'Y','Auto',GETDATE(),'',''," & pcsQtemp
    temp = temp & ",'Y','Auto',sysdate,'','',1, " & pcsQtemp

'    Debug.Print temp

             '�ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
    If (JudgeFlagStautsBC(source_batch_id_Temp)) Then
       MsgBox "��ʣ�" & source_batch_id_Temp & "�Ѵ��ڣ������ϴ�!"
       GoTo NextRecord2

    End If
    
    
    If (Not JudgeFlagStautsBCQty(source_batch_id_Temp, dieQtyTemp)) Then
       MsgBox "��ʣ�" & source_batch_id_Temp & "��BI�е�Die������һ��!"
       GoTo NextRecord2

    End If


    Call AddBC(temp, temp2)
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
