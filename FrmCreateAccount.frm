VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCreateAccount 
   Caption         =   "开通账户"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14445
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
   ScaleHeight     =   8610
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   19711
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "账号权限"
      TabPicture(0)   =   "FrmCreateAccount.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtUser"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "账号开通"
      TabPicture(1)   =   "FrmCreateAccount.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btnCopyRight"
      Tab(1).Control(1)=   "btnRemoveAccount"
      Tab(1).Control(2)=   "btnUpdatePassword"
      Tab(1).Control(3)=   "txtUserName2"
      Tab(1).Control(4)=   "btnCreateAccount"
      Tab(1).Control(5)=   "btnGetPassword"
      Tab(1).Control(6)=   "txtPassword"
      Tab(1).Control(7)=   "txtUserName"
      Tab(1).Control(8)=   "lblUserRealName(1)"
      Tab(1).Control(9)=   "lblUserRealName(0)"
      Tab(1).Control(10)=   "lblUserName2"
      Tab(1).Control(11)=   "lblPassword"
      Tab(1).Control(12)=   "lblUserName"
      Tab(1).ControlCount=   13
      Begin VB.Frame Frame3 
         Caption         =   "密码相关"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   7560
         TabIndex        =   26
         Top             =   1440
         Width           =   3015
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "查看密码"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtNewPasswd 
            BackColor       =   &H00FFC0FF&
            Height          =   285
            Left            =   960
            TabIndex        =   29
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "修改密码"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtPasswd 
            BackColor       =   &H00FFC0FF&
            Height          =   285
            Left            =   960
            TabIndex        =   27
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lbl23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "新密码"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   32
            Top             =   720
            Width           =   675
         End
         Begin VB.Label lbl2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "原密码"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "账号开通& 权限复制"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   360
         TabIndex        =   20
         Top             =   1440
         Width           =   6855
         Begin VB.TextBox txtLike 
            BackColor       =   &H00FFC0FF&
            Height          =   285
            Left            =   720
            TabIndex        =   25
            Top             =   480
            Width           =   1935
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "开通账号"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   960
            Width           =   2055
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "复制ERP权限"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   960
            Width           =   2055
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "复制生产管理权限"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   0
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label lbl3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "参考:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   525
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ERP功能单项开通"
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Left            =   360
         TabIndex        =   16
         Top             =   3360
         Width           =   3135
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "开通ERP功能"
            Height          =   480
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFC0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   17
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "功能模块代码"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1350
         End
      End
      Begin VB.CommandButton btnCopyRight 
         Caption         =   "复制erp权限"
         Height          =   360
         Left            =   -64440
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnRemoveAccount 
         Caption         =   "删除账号"
         Height          =   360
         Left            =   -73200
         TabIndex        =   12
         Top             =   3240
         Width           =   990
      End
      Begin VB.CommandButton btnUpdatePassword 
         Caption         =   "修改密码"
         Height          =   360
         Left            =   -73200
         TabIndex        =   11
         Top             =   3720
         Width           =   990
      End
      Begin VB.TextBox txtUserName2 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   -67680
         TabIndex        =   10
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton btnCreateAccount 
         Caption         =   "开通账号"
         Height          =   360
         Left            =   -74280
         TabIndex        =   8
         Top             =   3240
         Width           =   990
      End
      Begin VB.CommandButton btnGetPassword 
         Caption         =   "查询密码"
         Height          =   360
         Left            =   -74280
         TabIndex        =   7
         Top             =   3720
         Width           =   990
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   -73920
         TabIndex        =   6
         Top             =   2100
         Width           =   2175
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   -73920
         TabIndex        =   4
         Top             =   1230
         Width           =   2175
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblUserRealName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "员工姓名"
         Height          =   195
         Index           =   1
         Left            =   -65400
         TabIndex        =   14
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label lblUserRealName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "员工姓名"
         Height          =   195
         Index           =   0
         Left            =   -71520
         TabIndex        =   13
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label lblUserName2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参考账号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -68640
         TabIndex        =   9
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74400
         TabIndex        =   5
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "账号"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74400
         TabIndex        =   3
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工号:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   525
      End
   End
End
Attribute VB_Name = "FrmCreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Private Sub btnCopyRight_Click()
Dim sOra  As String
Dim sUser As String
Dim sLike As String

If txtUserName.text = "" Then
    MsgBox "请输入工号", vbInformation, "提示"
    Exit Sub

End If

If txtUserName2.text = "" Then
    MsgBox "请输入参考工号", vbCritical, "提示"
    Exit Sub

End If

sUser = Trim$(txtUserName.text)
sLike = Trim$(txtUserName2.text)

If MsgBox("是否确定" + sUser + "复制" + sLike + "的ERP账号权限?", vbYesNo, "提示") = vbNo Then
    Exit Sub
End If

AddSql2 ("delete from tblOperatorright where 用户号= '" & sUser & "'")

sOra = "insert into tblOperatorright select '" & sUser & "', 权限明细简码,'1' from tblOperatorright where 用户号 = '" & sLike & "'"
Exec_Sql (sOra)
MsgBox "ERP权限已复制成功", vbInformation, "提示"
End Sub

Private Sub cmd_Click(Index As Integer)

Select Case Index

    Case 0  ' 开通权限
        CreateAccount

    Case 1  ' 修改密码
        ChangePasswd

End Select

End Sub

Sub CreateAccount()
Dim sOra  As String
Dim sUser As String
Dim sLike As String

If txtUser.text = "" Then
    MsgBox "请输入工号", vbInformation, "提示"
    Exit Sub

End If

If txtLike.text = "" Then
    MsgBox "请输入参考工号", vbCritical, "提示"
    Exit Sub

End If

sUser = Trim$(txtUser.text)
sLike = Trim$(txtLike.text)

If MsgBox("是否确定" + sUser + "复制" + sLike + "的生产管理账号权限?", vbYesNo, "提示") = vbNo Then
    Exit Sub
End If

AddSql ("delete from tblGrantSysMenu where username = '" & sUser & "'")
sOra = "insert into tblGrantSysMenu select id, '" & sUser & "',blnadd, blnmodify, blndelete, blnsave, blnconfirm, sysdate,''  from tblGrantSysMenu where username = '" & sLike & "'"
Exec_Ora (sOra)
MsgBox "生产管理权限已复制成功", vbInformation, "提示"

End Sub

Private Sub Command2_Click()
Dim sOra  As String
Dim sUser As String
Dim sLike As String

If txtUser.text = "" Then
    MsgBox "请输入工号", vbInformation, "提示"
    Exit Sub

End If

If txtLike.text = "" Then
    MsgBox "请输入参考工号", vbCritical, "提示"
    Exit Sub

End If

sUser = Trim$(txtUser.text)
sLike = Trim$(txtLike.text)

If MsgBox("是否确定" + sUser + "复制" + sLike + "的ERP账号权限?", vbYesNo, "提示") = vbNo Then
    Exit Sub
End If

AddSql2 ("delete from tblOperatorright where 用户号= '" & sUser & "'")

sOra = "insert into tblOperatorright select '" & sUser & "', 权限明细简码,'1' from tblOperatorright where 用户号 = '" & sLike & "'"
Exec_Sql (sOra)
MsgBox "ERP权限已复制", vbInformation, "提示"

End Sub

Sub ChangePasswd()
Dim sSql    As String
Dim sUser   As String
Dim sPasswd As String
Dim sNew    As String

If txtUser.text = "" Then
    MsgBox "请输入工号", vbInformation, "提示"
    Exit Sub

End If

If txtPasswd.text = "" Then
    MsgBox "请输入原密码", vbInformation, "提示"
    Exit Sub

End If

sUser = Trim$(txtUser.text)
sPasswd = Trim$(txtPasswd.text)
sSql = "select 密码 from tblOperatorData where 用户号 = '" & sUser & "'"
If Get_SqlStr(sSql) <> sPasswd Then
    MsgBox "原密码不正确, 请输入正确的原密码", vbInformation, "提示"
    Exit Sub

End If

If txtNewPasswd.text = "" Then
    MsgBox "请输入更改后的密码", vbInformation, "提示"
    Exit Sub

End If

sNew = Trim$(txtNewPasswd.text)
sSql = "update tblOperatorData set 密码 = '" & sNew & "' where 用户名 = '" & sUser & "'"
Exec_Sql (sSql)
MsgBox "密码更改为:" & sNew, vbInformation, "提示"

End Sub

Private Sub Command1_Click()
If txtUser.text = "" Then
    MsgBox "请输入工号", vbInformation, "提示"
    Exit Sub

End If

Dim sUser As String

sUser = Trim$(txtUser.text)
txtPasswd.text = Get_SqlStr("select 密码 from tblOperatorData where 用户号 = '" & sUser & "'")

End Sub

Private Sub Command3_Click()
Dim sSql As String
Dim sMod As String

If Text1.text = "" Then
    MsgBox "请输入模块代码", vbInformation, "提示"
    Exit Sub

End If

sMod = UCase(Trim$(Text1.text))
If txtUser.text = "" Then
    MsgBox "请输入工号", vbInformation, "提示"
    Exit Sub

End If

sUser = Trim$(txtUser.text)
sSql = "insert into tblOperatorright values('" & sUser & "', '" & sMod & "', '1')"
Exec_Sql (sSql)
MsgBox "模块已经更新权限", vbInformation, "提示"

End Sub

Private Sub Command4_Click()
Dim sUser As String
Dim sOra  As String

If txtUser.text = "" Then
    MsgBox "请输入工号", vbInformation, "提示"
    Exit Sub

End If

sUser = Trim(txtUser.text)
If MsgBox("是否确定开通" + sUser + "的生产管理和ERP账号?", vbYesNo, "提示") = vbNo Then
    Exit Sub
End If

AddSql2 ("delete from erpbase.dbo.tblOperatorright where 用户号 = '" & sUser & "'")
AddSql2 ("delete from erpbase..tblOperatorData where 用户号 = '" & sUser & "' ")
sOra = "insert into erpbase..tblOperatorData values('" & sUser & "','" & sUser & "', '3', '1','','0') "
Exec_Sql (sOra)

MsgBox "已开通账号", vbInformation, Me.Caption

End Sub

Private Sub btnCreateAccount_Click()
Dim userName      As String
Dim userName2     As String
Dim userRealName  As String
Dim userRealName2 As String
Dim password      As String
Dim sql           As String

If txtUserName.text = "" Then
    MsgBox "请输入账号", vbInformation, "提示"
    Exit Sub

End If

userName = Trim$(txtUserName.text)
If txtUserName2.text = "" Then
    MsgBox "请输入参考账号", vbInformation, "提示"
    Exit Sub

End If

userName2 = Trim$(txtUserName2.text)
' 开通账号
sql = "select EmpName from XTW..employee where empno = '" & userName & "'"
userRealName = Get_SqlStr2(sql)
If userRealName = "" Then
    MsgBox "请输入正确的工号作为登录账号", vbInformation, "提示"
    Exit Sub

End If

lblUserRealName(0).Caption = userRealName
sql = "select * from erpbase..tblOperatorData where 用户号 = '" & userName & "'"
If Get_SqlStr(sql) <> "" Then
    MsgBox "该账号已经开通,不可重复开通", vbCritical, "警告"
    Exit Sub

End If

If txtPassword.text = "" Then
    password = userName
    sql = "insert into erpbase..tblOperatorData(用户号,密码,权限级别,状态标记,说明,禁用权限) values('" & userName & "','" & password & "', '3', '1','','0')"
Else
    password = Trim(txtPassword.text)
    sql = "insert into erpbase..tblOperatorData(用户号,密码,权限级别,状态标记,说明,禁用权限) values('" & userName & "','" & password & "', '3', '1','','0')"

End If

If AddSql2(sql) > 0 Then
    MsgBox "ERP系统(包括生产管理系统)账号已经成功开通" & vbCrLf & "用户名: " & userName & vbCrLf & "密码: " & password, vbInformation, "提示"

End If

' 开通生产管理系统权限
sql = "select EmpName from XTW..employee where charindex(empno, '" & userName2 & "') > 0 "
userRealName2 = Get_SqlStr2(sql)
If userRealName2 = "" Then
    MsgBox "请输入正确的工号作为参考登录账号", vbInformation, "提示"
    Exit Sub

End If

lblUserRealName(1).Caption = userRealName2
sql = "select * from erpbase..tblOperatorData where 用户号 = '" & userName2 & "'"
If Get_SqlStr(sql) = "" Then
    MsgBox "该参考账号不存在,请确认是否填写错误", vbCritical, "警告"
    Exit Sub

End If

AddSql ("delete from tblGrantSysMenu where username = '" & userName & "'")
AddSql ("insert into tblGrantSysMenu select id, '" & userName & "',blnadd, blnmodify, blndelete, blnsave, blnconfirm, sysdate,''  from tblGrantSysMenu where username = '" & userName2 & "'")
MsgBox "生产管理系统权限已经复制完成", vbInformation, "提示"
' ERP
AddSql2 ("delete from erpbase.dbo.tblOperatorright where 用户号 = '" & userName & "'")
AddSql2 ("insert into erpbase.dbo.tblOperatorright select '" & userName & "',权限明细简码,'1' from erpbase.dbo.tblOperatorright where 用户号 = '" & userName2 & "'")

If Get_SqlStr("select * from erpbase.dbo.tblpersonneldata where 员工编号 = '" & userName & "' ") = "" Then
    AddSql2 ("insert into erpbase.dbo.tblpersonneldata(员工编号, 员工姓名,部门编号,性别,区域号,工序号,是否禁用,标记,考勤组编号,是否计件类,产线标记) select '" & userName & "', '" & userRealName & "',部门编号,性别,区域号,工序号,是否禁用,标记,考勤组编号,是否计件类,产线标记 from erpbase.dbo.tblpersonneldata where 员工编号 = '" & userName2 & "'")
End If

MsgBox "ERP系统权限已经复制完成", vbInformation, "提示"

End Sub

Private Sub btnRemoveAccount_Click()
If txtUserName.text = "" Then
    MsgBox "请输入要删除的账号", vbInformation, "提示"
    Exit Sub
End If
AddSql2 ("delete from erpbase..tblOperatorright where 用户号 = '" & Trim(txtUserName.text) & "' ")
AddSql2 ("delete from erpbase..tblOperatorData where 用户号 = '" & Trim(txtUserName.text) & "'")
MsgBox "账号已经删除完成", vbInformation, "提示"

End Sub

Private Sub btnUpdatePassword_Click()
If txtUserName.text = "" Then
    MsgBox "请输入要修改的账号", vbInformation, "提示"
    Exit Sub
End If

If txtPassword.text = "" Then
    MsgBox "请输入要修改账户的新密码", vbInformation, "提示"
    Exit Sub
End If

AddSql2 ("update erpbase..tblOperatorData  set 密码 = '" & Trim(txtPassword.text) & "' where 用户号 = '" & Trim(txtUserName.text) & "'")

MsgBox "账号已经修改完成" & vbCrLf & "用户名: " & Trim(txtUserName.text) & vbCrLf & "新的密码: " & Trim(txtPassword.text), vbInformation, "提示"

End Sub

Private Sub btnGetPassword_Click()
txtPassword.text = Get_SqlStr("select 密码 from erpbase..tblOperatorData where 用户号 = '" & Trim(txtUserName.text) & "'")
End Sub

