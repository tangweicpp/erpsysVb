VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmSemtechWeiPi 
   Caption         =   "SemTech 仓库尾批料"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8595
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox TxtPT 
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox TxtWo 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin MSDataListLib.DataCombo weipiColl 
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label3 
      Caption         =   "客户机种号："
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "工单号："
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "仓库尾批："
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "FrmSemtechWeiPi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim woTemp As String
Dim ptTemp As String
Dim containerTemp As String
Dim qtyTemp As Long

woTemp = TxtWo.Text

ptTemp = TxtPT.Text

containerTemp = weipiColl.Text

qtyTemp = Mid(containerTemp, InStr(containerTemp, " "))

'qtyTemp = GetQty37WoWeiPiQty(containerTemp)



If woTemp <> "" And ptTemp <> "" And containerTemp <> "" Then

sqlTemp = " insert into  woLotWeiPi(wo,pt,containername,flag,qty) values  ('" & woTemp & "','" & ptTemp & "','" & containerTemp & "','Y'," & qtyTemp & ")"
AddSql (sqlTemp)


 MsgBox "添加成功!", vbInformation, "友情提示"

End If



FrmApplyWO.Show

Unload Me

End Sub

Private Sub Form_Load()
TxtWo.Text = FrmApplyWO.Text2.Text
TxtPT.Text = FrmApplyWO.TxtCustomerPT.Text

Call GetSemtechWeiPi(TxtPT.Text)

End Sub



Private Sub GetSemtechWeiPi(ptTemp As String)
'明细数据



Set childItemRS = GetWeiPiItem(ptTemp)

If childItemRS.RecordCount > 0 Then

Set weipiColl.RowSource = childItemRS
weipiColl.ListField = childItemRS("name").Name
weipiColl.BoundColumn = childItemRS("id").Name

End If




End Sub
