VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "用户登陆"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmLogin.frx":030A
   ScaleHeight     =   5280
   ScaleWidth      =   8790
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Combo1 
      ForeColor       =   &H00000000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5280
      TabIndex        =   14
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   1200
      Top             =   1320
   End
   Begin VB.CommandButton cmdTC 
      Caption         =   "退出(&X)"
      Height          =   345
      Left            =   7320
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdDL 
      Caption         =   "登陆(&O)"
      Height          =   345
      Left            =   7320
      TabIndex        =   0
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtPass 
      ForeColor       =   &H00000000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5280
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "3812"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "联系电话："
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "  维护人："
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Jiayun.Zhang"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   " 版  本："
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "华天科技（昆山） TSV生产管理系统"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   8175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "V-2016"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   2040
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "密 码"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   4725
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "用户名"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   4245
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************

'****************************************************************************

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'
Public gRS, iRS As New ADODB.Recordset
Public gsql, iSQL As String
Dim mDate1, mDate2, mDate3, mDate4 As Date
Public oSQL1 As String
Public iWAFERID, iNO, iSPCNotes, iSPCTYPE As String
Dim Alpha As Integer '声明变量




Private Sub cmdDL_Click()
If Me.Combo1.Text = "" Then
  MsgBox "请选择登陆用户！", 48, "错误提示"
  Exit Sub
End If


gsql = "select * from tblOperatorData r where  r.状态标记=1  and r.用户号='" & Me.Combo1.Text & "'and r.密码='" & Replace(Trim(txtPass.Text), "'", "") & "'"
If INIadoCon.State = 0 Then
INIConnectSTART
End If

If INIadoCon2.State = 0 Then
INIConnectSTART2
End If

Set gRS = INIadoCon.Execute(gsql)
 If Not gRS.EOF Then
        gUserName = Trim(gRS!用户号)
        gUserType = Trim(gRS!权限级别)
'
'        '2012-02-20 jiayunzhang add 把用户名，登录时间 放到数据库历史表中，用于以后查询用户用系统的频率
'        CnnSPC.Execute "insert into SPC_Login_History(username) values('" & gUserName & "')"
'
       MDIForm1.Show

       Unload Me
    Else
       MsgBox "密码错误！", 48, "错误提示"
    End If
End Sub

Private Sub cmdTC_Click()
End
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPass.SetFocus
End If

End Sub

Private Sub Form_Activate()
Me.Combo1.SetFocus
End Sub

Private Sub Form_Initialize()
Call InitCommonControls 'XP效果
End Sub

Private Sub Form_Load()
GetOracleConnection


Me.Shape1.Top = Me.Top '外边框
Me.Shape1.Left = Me.Left
Me.Shape1.Width = Me.ScaleWidth
Me.Shape1.Height = Me.ScaleHeight
'---------------------------------------------
Dim Ret As Long
    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    Timer1.Interval = 20


'  Combo1.Clear
' gsql = "select USERNAME from TSVSysUSER where flag='Y' order by usertype,username"
' Set gRS = CnnSPC.Execute(gsql)
' If Not gRS.EOF Then
'      gRS.MoveFirst
'    While Not gRS.EOF
'        With Combo1
'            .AddItem Trim(gRS!UserName)
'
'        End With
'        gRS.MoveNext
'       Wend
'
' End If


End Sub

Private Sub Timer1_Timer()
Alpha = Alpha + 5
If Alpha > 255 Then
   Timer1.Enabled = False
Exit Sub
End If
    SetLayeredWindowAttributes Me.hWnd, 0, Alpha, LWA_ALPHA
End Sub
Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdDL_Click
End If

End Sub
