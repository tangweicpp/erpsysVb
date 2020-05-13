VERSION 5.00
Begin VB.Form FrmSemtechPoAlter 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8550
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
   ScaleHeight     =   5025
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtText2 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtText1 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmdPO 
      Caption         =   "更改PO"
      Height          =   360
      Left            =   2520
      TabIndex        =   0
      Top             =   3120
      Width           =   990
   End
   Begin VB.Label lblPO_NUM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入PO_NUM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   840
      TabIndex        =   4
      Top             =   2640
      Width           =   1440
   End
   Begin VB.Label lblPO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入waferId"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1200
   End
End
Attribute VB_Name = "FrmSemtechPoAlter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPO_Click()
Dim cmd As New ADODB.Command
Dim RS  As New ADODB.Recordset
Dim strSql As String
Dim strValue As String


'输入框不能为空
If txtText1.Text = "" Then
MsgBox ("请输入要修改PO号的waferId")
Exit Sub
End If

If txtText2.Text = "" Then
MsgBox ("请输入要修改的PO号")
Exit Sub
End If
 
'先查询到waferId对应LOT号的ID
txtText1.Text = Trim(txtText1.Text)
strSql = "select a.filename from mappingdatatest a where a.substrateid= '" & txtText1.Text & "' and a.customershortname = '37'"

 If Cnn.State = 0 Then '如果数据库关闭则打开
    ConOracle
   End If
   
RS.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

If RS.RecordCount > 0 Then
 
 strValue = RS.fields(0).Value '取得waferId对应的lotId的ID
 Else
  MsgBox "查询不到任何信息"
End If

strSql = "update  customeroitbl_test set po_num =  '" & Trim(txtText2.Text) & "' where id= '" & strValue & "' and Customershortname = '37' "

 If Cnn.State = 0 Then '如果数据库关闭则打开
    ConOracle
   End If
   
cmd.ActiveConnection = Cnn
cmd.CommandText = strSql
cmd.CommandType = adCmdText
cmd.Execute

MsgBox "PO_NUM修改成功"

End Sub

