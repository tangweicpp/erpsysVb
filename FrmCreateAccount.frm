VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCreateAccount 
   Caption         =   "��ͨ�˻�"
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
      TabCaption(0)   =   "�˺�Ȩ��"
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
      TabCaption(1)   =   "�˺ſ�ͨ"
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
         Caption         =   "�������"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   7560
         TabIndex        =   26
         Top             =   1440
         Width           =   3015
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "�鿴����"
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
            Caption         =   "�޸�����"
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
            Caption         =   "������"
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
            Caption         =   "ԭ����"
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
         Caption         =   "�˺ſ�ͨ& Ȩ�޸���"
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
            Caption         =   "��ͨ�˺�"
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
            Caption         =   "����ERPȨ��"
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
            Caption         =   "������������Ȩ��"
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
            Caption         =   "�ο�:"
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
         Caption         =   "ERP���ܵ��ͨ"
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Left            =   360
         TabIndex        =   16
         Top             =   3360
         Width           =   3135
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "��ͨERP����"
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
            Caption         =   "����ģ�����"
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
         Caption         =   "����erpȨ��"
         Height          =   360
         Left            =   -64440
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnRemoveAccount 
         Caption         =   "ɾ���˺�"
         Height          =   360
         Left            =   -73200
         TabIndex        =   12
         Top             =   3240
         Width           =   990
      End
      Begin VB.CommandButton btnUpdatePassword 
         Caption         =   "�޸�����"
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
         Caption         =   "��ͨ�˺�"
         Height          =   360
         Left            =   -74280
         TabIndex        =   8
         Top             =   3240
         Width           =   990
      End
      Begin VB.CommandButton btnGetPassword 
         Caption         =   "��ѯ����"
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
         Caption         =   "Ա������"
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
         Caption         =   "Ա������"
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
         Caption         =   "�ο��˺�"
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
         Caption         =   "����"
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
         Caption         =   "�˺�"
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
         Caption         =   "����:"
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
    MsgBox "�����빤��", vbInformation, "��ʾ"
    Exit Sub

End If

If txtUserName2.text = "" Then
    MsgBox "������ο�����", vbCritical, "��ʾ"
    Exit Sub

End If

sUser = Trim$(txtUserName.text)
sLike = Trim$(txtUserName2.text)

If MsgBox("�Ƿ�ȷ��" + sUser + "����" + sLike + "��ERP�˺�Ȩ��?", vbYesNo, "��ʾ") = vbNo Then
    Exit Sub
End If

AddSql2 ("delete from tblOperatorright where �û���= '" & sUser & "'")

sOra = "insert into tblOperatorright select '" & sUser & "', Ȩ����ϸ����,'1' from tblOperatorright where �û��� = '" & sLike & "'"
Exec_Sql (sOra)
MsgBox "ERPȨ���Ѹ��Ƴɹ�", vbInformation, "��ʾ"
End Sub

Private Sub cmd_Click(Index As Integer)

Select Case Index

    Case 0  ' ��ͨȨ��
        CreateAccount

    Case 1  ' �޸�����
        ChangePasswd

End Select

End Sub

Sub CreateAccount()
Dim sOra  As String
Dim sUser As String
Dim sLike As String

If txtUser.text = "" Then
    MsgBox "�����빤��", vbInformation, "��ʾ"
    Exit Sub

End If

If txtLike.text = "" Then
    MsgBox "������ο�����", vbCritical, "��ʾ"
    Exit Sub

End If

sUser = Trim$(txtUser.text)
sLike = Trim$(txtLike.text)

If MsgBox("�Ƿ�ȷ��" + sUser + "����" + sLike + "�����������˺�Ȩ��?", vbYesNo, "��ʾ") = vbNo Then
    Exit Sub
End If

AddSql ("delete from tblGrantSysMenu where username = '" & sUser & "'")
sOra = "insert into tblGrantSysMenu select id, '" & sUser & "',blnadd, blnmodify, blndelete, blnsave, blnconfirm, sysdate,''  from tblGrantSysMenu where username = '" & sLike & "'"
Exec_Ora (sOra)
MsgBox "��������Ȩ���Ѹ��Ƴɹ�", vbInformation, "��ʾ"

End Sub

Private Sub Command2_Click()
Dim sOra  As String
Dim sUser As String
Dim sLike As String

If txtUser.text = "" Then
    MsgBox "�����빤��", vbInformation, "��ʾ"
    Exit Sub

End If

If txtLike.text = "" Then
    MsgBox "������ο�����", vbCritical, "��ʾ"
    Exit Sub

End If

sUser = Trim$(txtUser.text)
sLike = Trim$(txtLike.text)

If MsgBox("�Ƿ�ȷ��" + sUser + "����" + sLike + "��ERP�˺�Ȩ��?", vbYesNo, "��ʾ") = vbNo Then
    Exit Sub
End If

AddSql2 ("delete from tblOperatorright where �û���= '" & sUser & "'")

sOra = "insert into tblOperatorright select '" & sUser & "', Ȩ����ϸ����,'1' from tblOperatorright where �û��� = '" & sLike & "'"
Exec_Sql (sOra)
MsgBox "ERPȨ���Ѹ���", vbInformation, "��ʾ"

End Sub

Sub ChangePasswd()
Dim sSql    As String
Dim sUser   As String
Dim sPasswd As String
Dim sNew    As String

If txtUser.text = "" Then
    MsgBox "�����빤��", vbInformation, "��ʾ"
    Exit Sub

End If

If txtPasswd.text = "" Then
    MsgBox "������ԭ����", vbInformation, "��ʾ"
    Exit Sub

End If

sUser = Trim$(txtUser.text)
sPasswd = Trim$(txtPasswd.text)
sSql = "select ���� from tblOperatorData where �û��� = '" & sUser & "'"
If Get_SqlStr(sSql) <> sPasswd Then
    MsgBox "ԭ���벻��ȷ, ��������ȷ��ԭ����", vbInformation, "��ʾ"
    Exit Sub

End If

If txtNewPasswd.text = "" Then
    MsgBox "��������ĺ������", vbInformation, "��ʾ"
    Exit Sub

End If

sNew = Trim$(txtNewPasswd.text)
sSql = "update tblOperatorData set ���� = '" & sNew & "' where �û��� = '" & sUser & "'"
Exec_Sql (sSql)
MsgBox "�������Ϊ:" & sNew, vbInformation, "��ʾ"

End Sub

Private Sub Command1_Click()
If txtUser.text = "" Then
    MsgBox "�����빤��", vbInformation, "��ʾ"
    Exit Sub

End If

Dim sUser As String

sUser = Trim$(txtUser.text)
txtPasswd.text = Get_SqlStr("select ���� from tblOperatorData where �û��� = '" & sUser & "'")

End Sub

Private Sub Command3_Click()
Dim sSql As String
Dim sMod As String

If Text1.text = "" Then
    MsgBox "������ģ�����", vbInformation, "��ʾ"
    Exit Sub

End If

sMod = UCase(Trim$(Text1.text))
If txtUser.text = "" Then
    MsgBox "�����빤��", vbInformation, "��ʾ"
    Exit Sub

End If

sUser = Trim$(txtUser.text)
sSql = "insert into tblOperatorright values('" & sUser & "', '" & sMod & "', '1')"
Exec_Sql (sSql)
MsgBox "ģ���Ѿ�����Ȩ��", vbInformation, "��ʾ"

End Sub

Private Sub Command4_Click()
Dim sUser As String
Dim sOra  As String

If txtUser.text = "" Then
    MsgBox "�����빤��", vbInformation, "��ʾ"
    Exit Sub

End If

sUser = Trim(txtUser.text)
If MsgBox("�Ƿ�ȷ����ͨ" + sUser + "�����������ERP�˺�?", vbYesNo, "��ʾ") = vbNo Then
    Exit Sub
End If

AddSql2 ("delete from erpbase.dbo.tblOperatorright where �û��� = '" & sUser & "'")
AddSql2 ("delete from erpbase..tblOperatorData where �û��� = '" & sUser & "' ")
sOra = "insert into erpbase..tblOperatorData values('" & sUser & "','" & sUser & "', '3', '1','','0') "
Exec_Sql (sOra)

MsgBox "�ѿ�ͨ�˺�", vbInformation, Me.Caption

End Sub

Private Sub btnCreateAccount_Click()
Dim userName      As String
Dim userName2     As String
Dim userRealName  As String
Dim userRealName2 As String
Dim password      As String
Dim sql           As String

If txtUserName.text = "" Then
    MsgBox "�������˺�", vbInformation, "��ʾ"
    Exit Sub

End If

userName = Trim$(txtUserName.text)
If txtUserName2.text = "" Then
    MsgBox "������ο��˺�", vbInformation, "��ʾ"
    Exit Sub

End If

userName2 = Trim$(txtUserName2.text)
' ��ͨ�˺�
sql = "select EmpName from XTW..employee where empno = '" & userName & "'"
userRealName = Get_SqlStr2(sql)
If userRealName = "" Then
    MsgBox "��������ȷ�Ĺ�����Ϊ��¼�˺�", vbInformation, "��ʾ"
    Exit Sub

End If

lblUserRealName(0).Caption = userRealName
sql = "select * from erpbase..tblOperatorData where �û��� = '" & userName & "'"
If Get_SqlStr(sql) <> "" Then
    MsgBox "���˺��Ѿ���ͨ,�����ظ���ͨ", vbCritical, "����"
    Exit Sub

End If

If txtPassword.text = "" Then
    password = userName
    sql = "insert into erpbase..tblOperatorData(�û���,����,Ȩ�޼���,״̬���,˵��,����Ȩ��) values('" & userName & "','" & password & "', '3', '1','','0')"
Else
    password = Trim(txtPassword.text)
    sql = "insert into erpbase..tblOperatorData(�û���,����,Ȩ�޼���,״̬���,˵��,����Ȩ��) values('" & userName & "','" & password & "', '3', '1','','0')"

End If

If AddSql2(sql) > 0 Then
    MsgBox "ERPϵͳ(������������ϵͳ)�˺��Ѿ��ɹ���ͨ" & vbCrLf & "�û���: " & userName & vbCrLf & "����: " & password, vbInformation, "��ʾ"

End If

' ��ͨ��������ϵͳȨ��
sql = "select EmpName from XTW..employee where charindex(empno, '" & userName2 & "') > 0 "
userRealName2 = Get_SqlStr2(sql)
If userRealName2 = "" Then
    MsgBox "��������ȷ�Ĺ�����Ϊ�ο���¼�˺�", vbInformation, "��ʾ"
    Exit Sub

End If

lblUserRealName(1).Caption = userRealName2
sql = "select * from erpbase..tblOperatorData where �û��� = '" & userName2 & "'"
If Get_SqlStr(sql) = "" Then
    MsgBox "�òο��˺Ų�����,��ȷ���Ƿ���д����", vbCritical, "����"
    Exit Sub

End If

AddSql ("delete from tblGrantSysMenu where username = '" & userName & "'")
AddSql ("insert into tblGrantSysMenu select id, '" & userName & "',blnadd, blnmodify, blndelete, blnsave, blnconfirm, sysdate,''  from tblGrantSysMenu where username = '" & userName2 & "'")
MsgBox "��������ϵͳȨ���Ѿ��������", vbInformation, "��ʾ"
' ERP
AddSql2 ("delete from erpbase.dbo.tblOperatorright where �û��� = '" & userName & "'")
AddSql2 ("insert into erpbase.dbo.tblOperatorright select '" & userName & "',Ȩ����ϸ����,'1' from erpbase.dbo.tblOperatorright where �û��� = '" & userName2 & "'")

If Get_SqlStr("select * from erpbase.dbo.tblpersonneldata where Ա����� = '" & userName & "' ") = "" Then
    AddSql2 ("insert into erpbase.dbo.tblpersonneldata(Ա�����, Ա������,���ű��,�Ա�,�����,�����,�Ƿ����,���,��������,�Ƿ�Ƽ���,���߱��) select '" & userName & "', '" & userRealName & "',���ű��,�Ա�,�����,�����,�Ƿ����,���,��������,�Ƿ�Ƽ���,���߱�� from erpbase.dbo.tblpersonneldata where Ա����� = '" & userName2 & "'")
End If

MsgBox "ERPϵͳȨ���Ѿ��������", vbInformation, "��ʾ"

End Sub

Private Sub btnRemoveAccount_Click()
If txtUserName.text = "" Then
    MsgBox "������Ҫɾ�����˺�", vbInformation, "��ʾ"
    Exit Sub
End If
AddSql2 ("delete from erpbase..tblOperatorright where �û��� = '" & Trim(txtUserName.text) & "' ")
AddSql2 ("delete from erpbase..tblOperatorData where �û��� = '" & Trim(txtUserName.text) & "'")
MsgBox "�˺��Ѿ�ɾ�����", vbInformation, "��ʾ"

End Sub

Private Sub btnUpdatePassword_Click()
If txtUserName.text = "" Then
    MsgBox "������Ҫ�޸ĵ��˺�", vbInformation, "��ʾ"
    Exit Sub
End If

If txtPassword.text = "" Then
    MsgBox "������Ҫ�޸��˻���������", vbInformation, "��ʾ"
    Exit Sub
End If

AddSql2 ("update erpbase..tblOperatorData  set ���� = '" & Trim(txtPassword.text) & "' where �û��� = '" & Trim(txtUserName.text) & "'")

MsgBox "�˺��Ѿ��޸����" & vbCrLf & "�û���: " & Trim(txtUserName.text) & vbCrLf & "�µ�����: " & Trim(txtPassword.text), vbInformation, "��ʾ"

End Sub

Private Sub btnGetPassword_Click()
txtPassword.text = Get_SqlStr("select ���� from erpbase..tblOperatorData where �û��� = '" & Trim(txtUserName.text) & "'")
End Sub

