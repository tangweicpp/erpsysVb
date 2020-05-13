VERSION 5.00
Begin VB.Form FrmDoubleData 
   Caption         =   "抛工单 数据重复 异常处理"
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
      Caption         =   "提交"
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
      Caption         =   "工单号："
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
'判断　工单号
Dim workOrderTemp As String
Dim createdDateTemp As Date
Dim createdDateTemp2 As Date
Dim date1 As Date

 If UCase(Trim(Text1.Text)) = "" Then
      MsgBox "工单号不可以为空！"
     Exit Sub
 
 End If
 
 '工单号长度
  If Len(UCase(Trim(Text1.Text))) <> 12 Then
      MsgBox "工单号输入错误，请重新确认！"
     Exit Sub
 
 End If
 
 '工单号开立日期，超过一天的工单，不让删。
 
 workOrderTemp = UCase(Trim(Text1.Text))
 date1 = Now
  Set oiRS = GetWoCreatedDate(workOrderTemp)
If (oiRS.RecordCount > 0) Then
    createdDateTemp = oiRS.fields("erpcreationdate").Value
    
    If date1 - createdDateTemp > 1 Then
           MsgBox "工单号已开立超过一天，不允许删除！"
           Exit Sub
    End If
      
End If


 '工单号开立日期，超过一个小时的，不让删。
 
  Set oiRS = GetWoCreatedDate2(workOrderTemp)
If (oiRS.RecordCount > 0) Then
    createdDateTemp2 = oiRS.fields("txntimestamp").Value
    
    If (date1 - createdDateTemp2) * 24 > 1 Then
           MsgBox "工单已开立超过一个小时，不允许删除！"
           Exit Sub
    End If
      
End If

'执行删除

'Cnn.Execute del_Double_Data(workOrderTemp)

Dim Cmd As New ADODB.Command
Set Cmd = New ADODB.Command
Set Cmd.ActiveConnection = Cnn
 Cmd.CommandType = adCmdStoredProc
 Cmd.CommandText = "del_Double_Data"
 Cmd.Parameters.Append Cmd.CreateParameter("woTemp", adVarChar, adParamInput, workOrderTemp)
 
 Cmd.Execute

                  

End Sub
