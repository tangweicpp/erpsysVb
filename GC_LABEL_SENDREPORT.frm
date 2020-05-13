VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form GC_LABEL_SENDREPORT 
   Caption         =   "GC标签发货资料二级代码维护"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15495
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
   ScaleHeight     =   7575
   ScaleWidth      =   15495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra 
      Caption         =   "操作中心"
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtText5 
         Height          =   405
         Left            =   480
         TabIndex        =   15
         Top             =   5520
         Width           =   2775
      End
      Begin VB.TextBox txtText4 
         Height          =   405
         Left            =   480
         TabIndex        =   13
         Top             =   3840
         Width           =   2775
      End
      Begin VB.TextBox txtText3 
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   4680
         Width           =   2775
      End
      Begin VB.TextBox txtText2 
         Height          =   405
         Left            =   480
         TabIndex        =   5
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox txtText1 
         Height          =   405
         Left            =   480
         TabIndex        =   4
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CommandButton DELETEcmd 
         Caption         =   "删除"
         Height          =   360
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   990
      End
      Begin VB.CommandButton insertCmd 
         Caption         =   "新增"
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   990
      End
      Begin VB.CommandButton cmd 
         Caption         =   "查询"
         Height          =   360
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   990
      End
      Begin VB.Label lblWLAWLT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "维护为WLA/WLT类型"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   14
         Top             =   5160
         Width           =   2265
      End
      Begin VB.Label lblNORMAL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "维护为NORMAL类型"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   12
         Top             =   3480
         Width           =   2160
      End
      Begin VB.Label lblWLAWLT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WLA/WLT料号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   4320
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "normal最后定值二级代码"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   8
         Top             =   2520
         Width           =   2670
      End
      Begin VB.Label lblNORMAL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "normal料号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   1680
         Width           =   1050
      End
   End
   Begin FPSpreadADO.fpSpread Fps 
      Height          =   6495
      Left            =   4800
      TabIndex        =   10
      Top             =   120
      Width           =   10575
      _Version        =   524288
      _ExtentX        =   18653
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
      SpreadDesigner  =   "GC_LABEL_SENDREPORT.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label lblNORMAL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请维护为NORMAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   11
      Top             =   3600
      Width           =   2670
   End
End
Attribute VB_Name = "GC_LABEL_SENDREPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click()
Dim strsql              As String
Dim Rs                  As New ADODB.Recordset

strsql = "select * from gc_label_send_doublecode"
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

Private Sub DELETEcmd_Click()
Dim normalProduct As String
Dim normalProductDoubleCode As String
Dim normalType As String
Dim wlaWltProduct As String
Dim wlawltType As String
Dim Rs As New Recordset
Dim strsql As String
Dim sqlSwverSql As String
Dim cmd As New ADODB.Command

normalProduct = Trim(txtText1.Text)
normalProductDoubleCode = Trim(txtText2.Text)
normalType = Trim(txtText4.Text)
wlaWltProduct = Trim(txtText3.Text)
wlawltType = Trim(txtText5.Text)

If normalProduct = "" And wlaWltProduct = "" Then
   MsgBox "请输入wlaWltProduct！料号"
Else
  strsql = "select NORMALPRODUCT,NORMALDOUBLECODE,NORMALTYPE,WLAWLTPRODUCT,WLAWLTTYPE from gc_label_send_doublecode where WLAWLTPRODUCT='" & wlaWltProduct & "'"
  If Cnn.State = 0 Then
  ConOracle
  End If
  Rs.open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
   
  If Rs.RecordCount > 0 Then
  wlawltType = Rs.fields(4).Value
  normalProduct = Rs.fields(0).Value
  normalProductDoubleCode = Rs.fields(1).Value
  normalType = Rs.fields(2).Value
  wlaWltProduct = Rs.fields(3).Value
  Else
   MsgBox "料号不存在无法删除！"
   Exit Sub
  End If
   '删除ORACLE
   strsql = "delete gc_label_send_doublecode  where WLAWLTPRODUCT='" & wlaWltProduct & "'"
   If Cnn.State = 0 Then
    ConOracle
   End If
   
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strsql
   cmd.CommandType = adCmdText
   cmd.Execute
   '删除SQLsever一份发货报表使用
   sqlSwverSql = "delete gc_label_send_doublecode  where WLAWLTPRODUCT='" & wlaWltProduct & "'"
   cmd.ActiveConnection = INIadoCon2
   cmd.CommandText = sqlSwverSql
   cmd.CommandType = adCmdText
   cmd.Execute
   '记录日志
   strsql = "insert into gc_label_send_doublecode_log (NORMALPRODUCT,NORMALDOUBLECODE,NORMALTYPE,WLAWLTPRODUCT,WLAWLTTYPE,NOTE1,NOTE2) " & _
   " values('" & normalProduct & "','" & normalProductDoubleCode & "','" & normalType & "','" & wlaWltProduct & "','" & wlawltType & "','delete','" & gUserName & "')"
   If Cnn.State = 0 Then
    ConOracle
   End If
   
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strsql
   cmd.CommandType = adCmdText
   cmd.Execute
   
   MsgBox "删除成功 "
End If


End Sub

Private Sub Form_Load()

If gUserName = "08240" Or gUserName = "07885" Then

 DELETEcmd.Enabled = True
 insertCmd.Enabled = True
Else
 DELETEcmd.Enabled = False
 insertCmd.Enabled = False
End If

End Sub

Private Sub insertCmd_Click()
Dim normalProduct As String
Dim normalProductDoubleCode As String
Dim normalType As String
Dim wlaWltProduct As String
Dim wlawltType As String

Dim strsql As String
Dim sqlSwverSql As String
Dim cmd As New ADODB.Command

normalProduct = Trim(txtText1.Text)
normalProductDoubleCode = Trim(txtText2.Text)
normalType = Trim(txtText4.Text)
wlaWltProduct = Trim(txtText3.Text)
wlawltType = Trim(txtText5.Text)

If normalProduct = "" Or normalProductDoubleCode = "" Or normalType = "" Or wlaWltProduct = "" Or wlawltType = "" Then
   MsgBox "信息都为必填！"
Else
   '信息插入ORACLE
   strsql = "insert into gc_label_send_doublecode (NORMALPRODUCT,NORMALDOUBLECODE,NORMALTYPE,WLAWLTPRODUCT,WLAWLTTYPE) values('" & normalProduct & "','" & normalProductDoubleCode & "','" & normalType & "','" & wlaWltProduct & "','" & wlawltType & "')"
   If Cnn.State = 0 Then
    ConOracle
   End If
   
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strsql
   cmd.CommandType = adCmdText
   cmd.Execute
   '信息插入SQLsever一份发货报表使用
   sqlSwverSql = "insert into gc_label_send_doublecode (NORMALPRODUCT,NORMALDOUBLECODE,NORMALTYPE,WLAWLTPRODUCT,WLAWLTTYPE) values('" & normalProduct & "','" & normalProductDoubleCode & "','" & normalType & "','" & wlaWltProduct & "','" & wlawltType & "')"
    cmd.ActiveConnection = INIadoCon2
   cmd.CommandText = sqlSwverSql
   cmd.CommandType = adCmdText
   cmd.Execute
   '记录日志
   strsql = "insert into gc_label_send_doublecode_log (NORMALPRODUCT,NORMALDOUBLECODE,NORMALTYPE,WLAWLTPRODUCT,WLAWLTTYPE,NOTE1,NOTE2) values('" & normalProduct & "','" & normalProductDoubleCode & "','" & normalType & "','" & wlaWltProduct & "','" & wlawltType & "','insert','" & gUserName & "')"
   If Cnn.State = 0 Then
    ConOracle
   End If
   
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strsql
   cmd.CommandType = adCmdText
   cmd.Execute
   
   MsgBox "录入成功 "
End If
   

End Sub

