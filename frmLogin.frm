VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "用户登陆"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11730
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmLogin.frx":030A
   ScaleHeight     =   6075
   ScaleWidth      =   11730
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdDL 
      Caption         =   "登陆(Enter)"
      Default         =   -1  'True
      Height          =   345
      Left            =   9840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Combo1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   7920
      TabIndex        =   4
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   1200
      Top             =   1320
   End
   Begin VB.CommandButton cmdTC 
      Caption         =   "退出(Esc)"
      Height          =   345
      Left            =   9840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   7920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   7200
      TabIndex        =   6
      Top             =   4965
      Width           =   675
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
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   7320
      TabIndex        =   1
      Top             =   5325
      Width           =   615
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

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal crKey As Long, _
                              ByVal bAlpha As Byte, _
                              ByVal dwFlags As Long) As Long

'
Public gRS, iRS As New ADODB.Recordset

Public gsql, iSQL As String

Dim mDate1, mDate2, mDate3, mDate4 As Date

Public oSQL1 As String

Public iWAFERID, iNO, iSPCNotes, iSPCTYPE As String

Dim Alpha As Integer '声明变量

Private Sub cmdDL_Click()

    If Me.Combo1.text = "" Then
        MsgBox "请选择登陆用户！", 48, "错误提示"
        Exit Sub

    End If

    gsql = "select * from tblOperatorData r where  r.状态标记=1  and r.用户号='" & Me.Combo1.text & "'and r.密码='" & Replace(Trim(txtPass.text), "'", "") & "'"

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
        gUserRealName = Get_SqlStr2("select EmpName from XTW..employee where empno ='" & gUserName & "' ")
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
    ConnectOracle
    ConnectSql
    ConnectSql2

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        End

    End If

End Sub

Private Sub Form_Load()
    GetOracleConnection
    InitWorkPath
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

Private Sub InitWorkPath()

Dim strPath As String

strPath = "C:\test\"
If Dir(strPath, vbDirectory) = "" Then
    MkDir strPath

End If

strPath = "C:\TSVWoLog\"
If Dir(strPath, vbDirectory) = "" Then
    MkDir strPath

End If

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
