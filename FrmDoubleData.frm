VERSION 5.00
Begin VB.Form FrmDoubleData 
   Caption         =   "�׹��� �����ظ� �쳣����"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
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
   ScaleHeight     =   5310
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "�ύ"
      Height          =   360
      Left            =   6120
      TabIndex        =   2
      Top             =   480
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ţ�"
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   840
   End
End
Attribute VB_Name = "FrmDoubleData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'�жϡ�������
Dim workOrderTemp As String
Dim createdDateTemp As Date
Dim createdDateTemp2 As Date
Dim date1 As Date

 If UCase(Trim(Text1.Text)) = "" Then
      MsgBox "�����Ų�����Ϊ�գ�"
     Exit Sub
 
 End If
 
 '�����ų���
  If Len(UCase(Trim(Text1.Text))) <> 12 Then
      MsgBox "�������������������ȷ�ϣ�"
     Exit Sub
 
 End If
 
 '�����ſ������ڣ�����һ��Ĺ���������ɾ��
 
 workOrderTemp = UCase(Trim(Text1.Text))
 date1 = Now
  Set oiRS = GetWoCreatedDate(workOrderTemp)
If (oiRS.RecordCount > 0) Then
    createdDateTemp = oiRS.fields("erpcreationdate").Value
    
    If date1 - createdDateTemp > 1 Then
           MsgBox "�������ѿ�������һ�죬������ɾ����"
           Exit Sub
    End If
      
End If


 '�����ſ������ڣ�����һ��Сʱ�ģ�����ɾ��
 
  Set oiRS = GetWoCreatedDate2(workOrderTemp)
If (oiRS.RecordCount > 0) Then
    createdDateTemp2 = oiRS.fields("txntimestamp").Value
    
    If (date1 - createdDateTemp2) * 24 > 1 Then
           MsgBox "�����ѿ�������һ��Сʱ��������ɾ����"
           Exit Sub
    End If
      
End If

'ִ��ɾ��

'Cnn.Execute del_Double_Data(workOrderTemp)

Dim Cmd As New ADODB.Command
Set Cmd = New ADODB.Command
Set Cmd.ActiveConnection = Cnn
 Cmd.CommandType = adCmdStoredProc
 Cmd.CommandText = "del_Double_Data"
 Cmd.Parameters.Append Cmd.CreateParameter("woTemp", adVarChar, adParamInput, workOrderTemp)
 
 Cmd.Execute

                  

End Sub
