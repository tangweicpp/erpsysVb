VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmTestSetUp 
   Caption         =   "���̿��ϵĲ��԰汾���趨(һ��һ)"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "���԰汾��"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      Begin VB.TextBox CmbTest 
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton CmdCardModify 
         Caption         =   "�޸�"
         Height          =   360
         Left            =   4320
         TabIndex        =   4
         Top             =   1680
         Width           =   990
      End
      Begin VB.CommandButton CmdCardAdd 
         Caption         =   "����"
         Height          =   360
         Left            =   2520
         TabIndex        =   3
         Top             =   1680
         Width           =   990
      End
      Begin MSDataListLib.DataCombo CmbProduct 
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���԰汾�ţ�"
         Height          =   195
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʒ����"
         Height          =   195
         Left            =   1620
         TabIndex        =   1
         Top             =   360
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmTestSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mainItemRS As New ADODB.Recordset
Dim pNameRS As New ADODB.Recordset
Dim testNameRS As New ADODB.Recordset

Private Sub CmdCardAdd_Click()
'2012-08-06 jiayun add

Dim alterNameTemp As String
Dim testVersionNameTemp As String

Dim sqlTemp As String
Dim judgeExist As Boolean
judgeExist = False


alterNameTemp = CmbProduct.Text
testVersionNameTemp = UCase(Trim(CmbTest.Text))

If alterNameTemp = "" Or testVersionNameTemp = "" Then

MsgBox "ѡ������գ���Ϊ������", vbInformation, "������ʾ"

Exit Sub

End If


'�жϳ����Ƿ񳬹�20λ
If Len(testVersionNameTemp) > 20 Then
 MsgBox "���԰汾�Ų����Գ���20λ !", vbInformation, "������ʾ"
 CmbTest.Text = ""
 CmbTest.SetFocus
 
 Exit Sub
 End If




'�ж������Ʒ�Ƿ��Ѵ���
judgeExist = JudgeDataExist(alterNameTemp)

 If judgeExist = True Then
 
 MsgBox "�Ѵ��ڣ����޸ģ�", vbInformation, "������ʾ"

 Exit Sub
 
 Else
 
sqlTemp = "insert into TSVCard_EDT( id,productname,testedition,createdby,createddate,flag) values (RCardTestVersionId.Nextval,'" & alterNameTemp & "','" & testVersionNameTemp & "','Auto',sysdate,'Y')"
AddSql (sqlTemp)

 MsgBox "��ӳɹ�!", vbInformation, "������ʾ"
 

 End If




End Sub

Public Function JudgeDataExist(alterNameTemp As String) As Boolean
'�ж�����Ƿ���ڦs�b
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select *  from  TSVCard_EDT where productname='" + alterNameTemp + "' and flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeDataExist = slectResult
End Function

Private Sub CmdCardModify_Click()
'�޸�

'2012-08-06 jiayun add

Dim alterNameTemp As String
Dim testVersionNameTemp As String

Dim sqlTemp As String
Dim judgeExist As Boolean
judgeExist = False


alterNameTemp = CmbProduct.Text
testVersionNameTemp = UCase(Trim(CmbTest.Text))

If alterNameTemp = "" Or testVersionNameTemp = "" Then

MsgBox "ѡ������գ���Ϊ������", vbInformation, "������ʾ"

Exit Sub

End If

'�жϳ����Ƿ񳬹�20λ
If Len(testVersionNameTemp) > 20 Then
 MsgBox "���԰汾�Ų����Գ���20λ !", vbInformation, "������ʾ"
 CmbTest.Text = ""
 CmbTest.SetFocus
 
 Exit Sub
 End If
 


'�ж������Ʒ�Ƿ��Ѵ���
judgeExist = JudgeDataExist(alterNameTemp)

 If judgeExist = True Then
 
 sqlTemp = "update TSVCard_EDT set testedition='" & testVersionNameTemp & "',lastupdateby='Auto', lastupdatedate=sysdate where productname = '" & alterNameTemp & "' and flag='Y'"
 
AddSql (sqlTemp)

 MsgBox "�޸ĳɹ�!", vbInformation, "������ʾ"
 

 
 Else
 

 
MsgBox "��ʲ����ڣ�����������", vbInformation, "������ʾ"

 Exit Sub
 
 

 End If



End Sub

Private Sub Form_Load()


'Ʒ����ʼ��
IniPName

''���԰汾�ų�ʼ��
'IniTestName




End Sub


Private Sub IniPName()
Set pNameRS = GetInitPName()
Set CmbProduct.RowSource = pNameRS
CmbProduct.ListField = pNameRS("ALTERNATENAME").Name
CmbProduct.BoundColumn = pNameRS("PID").Name

End Sub

'Private Sub IniTestName()
'Set testNameRS = GetInitTestName()
'Set CmbTest.RowSource = testNameRS
'CmbTest.ListField = testNameRS("QtechTestVersion").Name
'CmbTest.BoundColumn = testNameRS("PID").Name
'
'End Sub



