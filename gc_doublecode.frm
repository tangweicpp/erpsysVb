VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form gc_doublecode 
   Caption         =   "GC二级代码维护"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
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
   ScaleHeight     =   7080
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra 
      Caption         =   "操作中心"
      Height          =   6375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      Begin VB.TextBox txtText3 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   4080
         Width           =   2895
      End
      Begin VB.TextBox txtText2 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtText1 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除"
         Height          =   360
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   990
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "新增"
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   990
      End
      Begin VB.CommandButton cmdselect 
         Caption         =   "查询"
         Height          =   360
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   3720
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "二级代码最后一位"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "厂内机种名"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   900
      End
   End
   Begin FPSpreadADO.fpSpread Fps 
      Height          =   6495
      Left            =   4800
      TabIndex        =   4
      Top             =   480
      Width           =   6255
      _Version        =   524288
      _ExtentX        =   11033
      _ExtentY        =   11456
      _StockProps     =   64
      EditEnterAction =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   0
      MaxRows         =   0
      SpreadDesigner  =   "gc_doublecode.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "厂内机种名"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   9
      Top             =   3240
      Width           =   900
   End
End
Attribute VB_Name = "gc_doublecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
Dim strsql As String
Dim strsqllog As String
Dim sqlSeverSql As String
Dim Cmd As New ADODB.Command
Dim tempProduct As String
Dim Rs As New Recordset
Dim doubleCodeTemp As String


tempProduct = Trim(txtText1.Text)

  strsql = "select DOUBLECODE from GC_DOUBLECODE where PRODUCTNAME='" & tempProduct & "'"
   If Cnn.State = 0 Then
    ConOracle
   End If
   Rs.open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
   If Rs.RecordCount > 0 Then
   doubleCodeTemp = Rs.fields(0).Value
   Else
   MsgBox "要删除的料号不存在无法删除！"
   Exit Sub
   End If

 '删除ORACLE
    strsql = "delete from GC_DOUBLECODE where PRODUCTNAME='" & tempProduct & "'"

   If Cnn.State = 0 Then
    ConOracle
   End If

   Cmd.ActiveConnection = Cnn
   Cmd.CommandText = strsql
   Cmd.CommandType = adCmdText
   Cmd.Execute
   
   '删除SQL SEVER
   sqlSeverSql = "delete from GC_DOUBLECODE where PRODUCTNAME='" & tempProduct & "'"

   If Cnn.State = 0 Then
    ConOracle
   End If

   Cmd.ActiveConnection = INIadoCon2
   Cmd.CommandText = sqlSeverSql
   Cmd.CommandType = adCmdText
   Cmd.Execute
   
   
   '记录日志
   
   strsqllog = "insert into GC_DOUBLECODE_log (PRODUCTNAME,DOUBLECODE,NOTE1,NOTE2) values('" & tempProduct & "','" & doubleCodeTemp & "','DELETE','" & gUserName & "')"
   If Cnn.State = 0 Then
    ConOracle
   End If
   
   Cmd.ActiveConnection = Cnn
   Cmd.CommandText = strsqllog
   Cmd.CommandType = adCmdText
   Cmd.Execute
   
   MsgBox "删除成功 ！"


End Sub

Private Sub cmdInsert_Click()
Dim tempProduct As String
Dim tempDoubleCode As String
Dim tempNote1 As String
Dim strsql As String
Dim strsqllog As String
Dim sqlSeverSql As String
Dim Cmd As New ADODB.Command

tempProduct = Trim(txtText1.Text)
tempDoubleCode = Trim(txtText2.Text)
tempNote1 = Trim(txtText3.Text)

If tempProduct = "" Or tempDoubleCode = "" Then
   MsgBox "请录入厂内料号和二级代码！"
Else
   '插入ORACLE
   strsql = "insert into GC_DOUBLECODE values('" & tempProduct & "','" & tempDoubleCode & "','" & tempNote1 & "')"
   If Cnn.State = 0 Then
    ConOracle
   End If
   
   Cmd.ActiveConnection = Cnn
   Cmd.CommandText = strsql
   Cmd.CommandType = adCmdText
   Cmd.Execute
   '插入SQLSEVER
    sqlSeverSql = "insert into GC_DOUBLECODE values('" & tempProduct & "','" & tempDoubleCode & "','" & tempNote1 & "')"
   If Cnn.State = 0 Then
    ConOracle
   End If
   
   Cmd.ActiveConnection = INIadoCon2
   Cmd.CommandText = sqlSeverSql
   Cmd.CommandType = adCmdText
   Cmd.Execute
   MsgBox "录入成功 ！"
   
   
   '记录日志
   strsqllog = "insert into GC_DOUBLECODE_log (PRODUCTNAME,DOUBLECODE,NOTE1,NOTE2) values('" & tempProduct & "','" & tempDoubleCode & "','INSERT','" & gUserName & "')"
   If Cnn.State = 0 Then
    ConOracle
   End If
   
   Cmd.ActiveConnection = Cnn
   Cmd.CommandText = strsqllog
   Cmd.CommandType = adCmdText
   Cmd.Execute
   
   
End If

End Sub

Private Sub cmdselect_Click()
Dim strsql              As String
Dim Rs                  As New ADODB.Recordset

strsql = "select * from GC_DOUBLECODE"
If Cnn.State = 0 Then
  ConOracle
End If
    
Rs.open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

If Rs.RecordCount > 0 Then
 Set Fps.DataSource = Rs
 Fps.MaxRows = Rs.RecordCount
 
 Else
  MsgBox "查询不到任何信息"
End If




End Sub

Private Sub Form_Load()

If gUserName = "08240" Or gUserName = "07885" Then

 cmdDelete.Enabled = True
 cmdInsert.Enabled = True
Else
 cmdDelete.Enabled = False
 cmdInsert.Enabled = False
End If

End Sub
