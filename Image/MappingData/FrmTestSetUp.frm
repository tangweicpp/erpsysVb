VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmTestSetUp 
   Caption         =   "流程卡上的测试版本号设定(一对一)"
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
      Caption         =   "测试版本号"
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
         Caption         =   "修改"
         Height          =   360
         Left            =   4320
         TabIndex        =   4
         Top             =   1680
         Width           =   990
      End
      Begin VB.CommandButton CmdCardAdd 
         Caption         =   "新增"
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
         Caption         =   "测试版本号："
         Height          =   195
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "品名："
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

MsgBox "选项不可留空，都为必填项", vbInformation, "友情提示"

Exit Sub

End If


'判断长度是否超过20位
If Len(testVersionNameTemp) > 20 Then
 MsgBox "测试版本号不可以超过20位 !", vbInformation, "友情提示"
 CmbTest.Text = ""
 CmbTest.SetFocus
 
 Exit Sub
 End If




'判断这个产品是否已存在
judgeExist = JudgeDataExist(alterNameTemp)

 If judgeExist = True Then
 
 MsgBox "已存在，请修改！", vbInformation, "友情提示"

 Exit Sub
 
 Else
 
sqlTemp = "insert into TSVCard_EDT( id,productname,testedition,createdby,createddate,flag) values (RCardTestVersionId.Nextval,'" & alterNameTemp & "','" & testVersionNameTemp & "','Auto',sysdate,'Y')"
AddSql (sqlTemp)

 MsgBox "添加成功!", vbInformation, "友情提示"
 

 End If




End Sub

Public Function JudgeDataExist(alterNameTemp As String) As Boolean
'判断这笔是否存在
Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False

cmdStr = "select *  from  TSVCard_EDT where productname='" + alterNameTemp + "' and flag='Y'"

slectResult = QueryStr(cmdStr)
JudgeDataExist = slectResult
End Function

Private Sub CmdCardModify_Click()
'修改

'2012-08-06 jiayun add

Dim alterNameTemp As String
Dim testVersionNameTemp As String

Dim sqlTemp As String
Dim judgeExist As Boolean
judgeExist = False


alterNameTemp = CmbProduct.Text
testVersionNameTemp = UCase(Trim(CmbTest.Text))

If alterNameTemp = "" Or testVersionNameTemp = "" Then

MsgBox "选项不可留空，都为必填项", vbInformation, "友情提示"

Exit Sub

End If

'判断长度是否超过20位
If Len(testVersionNameTemp) > 20 Then
 MsgBox "测试版本号不可以超过20位 !", vbInformation, "友情提示"
 CmbTest.Text = ""
 CmbTest.SetFocus
 
 Exit Sub
 End If
 


'判断这个产品是否已存在
judgeExist = JudgeDataExist(alterNameTemp)

 If judgeExist = True Then
 
 sqlTemp = "update TSVCard_EDT set testedition='" & testVersionNameTemp & "',lastupdateby='Auto', lastupdatedate=sysdate where productname = '" & alterNameTemp & "' and flag='Y'"
 
AddSql (sqlTemp)

 MsgBox "修改成功!", vbInformation, "友情提示"
 

 
 Else
 

 
MsgBox "这笔不存在，请先新增！", vbInformation, "友情提示"

 Exit Sub
 
 

 End If



End Sub

Private Sub Form_Load()


'品名初始化
IniPName

''测试版本号初始化
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



