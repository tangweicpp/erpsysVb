VERSION 5.00
Begin VB.Form Form_OrderHistory 
   Caption         =   "开工单记录"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16635
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
   ScaleHeight     =   10125
   ScaleWidth      =   16635
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   720
      TabIndex        =   10
      Top             =   4920
      Width           =   7935
      Begin VB.CommandButton Command1 
         Caption         =   "导出WAFERID信息"
         Height          =   840
         Left            =   5040
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtSBoxNO 
         BackColor       =   &H00FFC0FF&
         Height          =   3735
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "小箱号"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.TextBox txtLotID 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4058
      Width           =   1815
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3053
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00FF80FF&
      Caption         =   "查询"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txtOrderID 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOTID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   8
      Top             =   240
      Width           =   795
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工号:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   6
      Top             =   4080
      Width           =   555
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开工单日期:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   5
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工单号:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   795
   End
End
Attribute VB_Name = "Form_OrderHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdQuery_Click()

Dim sOrder As String
Dim sOra As String
Dim rs As New ADODB.Recordset
Dim strLotID As String

If txtLotID.Text <> "" Then
    strLotID = Trim$(UCase$(txtLotID.Text))
    txtOrderID.Text = Get_OracleStr("select distinct ordername from ib_waferlist where waferlot = '" & strLotID & "'")
    Exit Sub
End If

If txtOrderID.Text = "" Then
    MsgBox "请输入工单号", vbInformation, "友情提示!!!"
    Exit Sub
End If

sOrder = Trim$(UCase$(txtOrderID.Text))

sOra = "select nvl(great_date, '无数据') as create_date, nvl(creat_by, '无数据') as create_by from PJ_WO_PRI where wo =  '" & sOrder & "'"
Set rs = Get_OracleRs(sOra)

If rs.RecordCount = 0 Then
    MsgBox "无法查到该工单的记录请确认", vbInformation, "友情提示!!!"
    Exit Sub
End If

txtDate.Text = rs!CREATE_DATE
txtName.Text = rs!CREATE_BY

End Sub

Private Sub Command1_Click()
Dim strArr() As String
Dim strSql As String
Dim strBoxID As String
Dim i As Integer

If txtSBoxNO.Text = "" Then
    MsgBox "请输入小箱号", vbInformation, "提示"
    Exit Sub
End If

strArr() = Split(UCase(Trim$(txtSBoxNO.Text)), vbCrLf)

For i = 0 To UBound(strArr)
    If strArr(i) <> "" Then
        strBoxID = strBoxID & Trim(strArr(i)) & "','"
    End If
Next

strBoxID = Left(strBoxID, Len(strBoxID) - 3)

strSql = "select * from  erpdata..tblPackMainInfSub where 箱号 in ('" & strBoxID & "')"
SqlServer2ExporToExcel (strSql)

End Sub
