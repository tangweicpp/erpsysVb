VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_GC_DIE_SEC_CODE 
   Caption         =   "GC�����ͻ�����DIE����/��������ά����WO��"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
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
   ScaleHeight     =   9045
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra 
      Caption         =   "ά����ϸ"
      ForeColor       =   &H00800000&
      Height          =   7455
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   11895
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5535
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   10695
         _Version        =   524288
         _ExtentX        =   18865
         _ExtentY        =   9763
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   0
         SpreadDesigner  =   "Frm_GC_DIE_SEC_CODE.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "ά��ѡ��"
      ForeColor       =   &H00800000&
      Height          =   2055
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton cmdRead 
         Caption         =   "�鿴"
         Height          =   360
         Left            =   360
         TabIndex        =   14
         Top             =   1440
         Width           =   990
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�˳�"
         Height          =   360
         Left            =   6000
         TabIndex        =   13
         Top             =   1440
         Width           =   990
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��"
         Height          =   360
         Left            =   4590
         TabIndex        =   12
         Top             =   1440
         Width           =   990
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "����"
         Height          =   360
         Left            =   3180
         TabIndex        =   11
         Top             =   1440
         Width           =   990
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "����"
         Height          =   360
         Left            =   1770
         TabIndex        =   10
         Top             =   1440
         Width           =   990
      End
      Begin VB.TextBox txtSecCode 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   5040
         TabIndex        =   9
         Top             =   383
         Width           =   855
      End
      Begin VB.TextBox txtDies 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   4680
         TabIndex        =   7
         Top             =   743
         Width           =   1215
      End
      Begin VB.ComboBox cbCustCode 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         Text            =   "GC"
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cbCustPN 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblSecCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����������λ"
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
         Index           =   3
         Left            =   3360
         TabIndex        =   8
         Top             =   405
         Width           =   1620
      End
      Begin VB.Label lblGROSSDIES 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GROSSDIES"
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
         Index           =   2
         Left            =   3600
         TabIndex        =   6
         Top             =   765
         Width           =   900
      End
      Begin VB.Label lblGROSSDIES 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����"
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
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   405
         Width           =   900
      End
      Begin VB.Label lblGROSSDIES 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����"
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
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   765
         Width           =   900
      End
   End
End
Attribute VB_Name = "Frm_GC_DIE_SEC_CODE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
Dim strSql As String

strSql = "delete from tblCustomerDieQty where CustomerPT = '" & cbCustPN.Text & "'"

If AddSql(strSql) Then
    MsgBox "ɾ���ɹ�", vbInformation, "��ʾ"

End If

End Sub

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdNew_Click()
Dim strSql      As String
Dim strCustCode As String
Dim strCustPN   As String
Dim strDies     As String
Dim strSecCode  As String
Dim rs          As New ADODB.Recordset

strCustCode = UCase(Trim(cbCustCode.Text))
strCustPN = UCase$(Trim$(cbCustPN.Text))
strDies = Trim$(txtDies.Text)
strSecCode = UCase(Trim$(txtSecCode.Text))

If strCustPN = "" Then
    MsgBox "��������Ҫ�����Ŀͻ�����", vbInformation, "��ʾ" '
    Exit Sub

End If

If strDies = "" Then
    MsgBox "������GROSSDIE����", vbInformation, "��ʾ" '
    Exit Sub

End If

If strSecCode = "" Then
    MsgBox "�������������", vbInformation, "��ʾ" '
    Exit Sub

End If

strSql = "select customername as �ͻ�����,CustomerPT as �ͻ�����,DIEQTY as GROSSDIES, GCVERSION as �����������λ, CREATEDBY as �½���Ա,CREATEDDATE as �½�ʱ��,LASTUPDATEBY as ������Ա,LASTUPDATEDATE as ����ʱ��  from tblCustomerDieQty where customername = 'GC' and CustomerPT = '" & strCustPN & "'  "
Set rs = Get_OracleRs(strSql)

If Not rs.EOF Then

    With fps
        .MaxRows = 0
        Set .DataSource = rs

    End With

    MsgBox "�ÿͻ������Ѿ���ά����¼,��������;ֻ���޸Ļ���ɾ��", vbInformation, "��ʾ"
Else
    strSql = "insert into tblCustomerDieQty(customername,CustomerPT,DIEQTY,GCVERSION,CREATEDBY,CREATEDDATE)  values('GC','" & strCustPN & "','" & strDies & "','" & strSecCode & "','" & gUserName & "',sysdate)"
    AddSql (strSql)
    MsgBox "�����ɹ�", vbInformation, "��ʾ"
    Call cmdRead_Click

End If

Set rs = Nothing

End Sub

Private Sub cmdRead_Click()
Dim strSql      As String
Dim strCustCode As String
Dim strCustPN   As String
Dim rs          As New ADODB.Recordset

strCustCode = UCase(Trim(cbCustCode.Text))
strCustPN = UCase$(Trim$(cbCustPN.Text))

If strCustPN = "" Then
    strSql = "select customername as �ͻ�����,CustomerPT as �ͻ�����,DIEQTY as GROSSDIES, GCVERSION as �����������λ, CREATEDBY as �½���Ա,CREATEDDATE as �½�ʱ��,LASTUPDATEBY as ������Ա,LASTUPDATEDATE as ����ʱ��  from tblCustomerDieQty where customername = 'GC'   "
Else
    strSql = "select customername as �ͻ�����,CustomerPT as �ͻ�����,DIEQTY as GROSSDIES, GCVERSION as �����������λ, CREATEDBY as �½���Ա,CREATEDDATE as �½�ʱ��,LASTUPDATEBY as ������Ա,LASTUPDATEDATE as ����ʱ��  from tblCustomerDieQty where customername = 'GC' and CustomerPT = '" & strCustPN & "'  "

End If

Set rs = Get_OracleRs(strSql)

With fps
    .MaxRows = 0
    Set .DataSource = rs

End With

Set rs = Nothing

End Sub

Private Sub cmdUpdate_Click()
Dim strSql      As String
Dim strCustCode As String
Dim strCustPN   As String
Dim strDies     As String
Dim strSecCode  As String
Dim rs          As New ADODB.Recordset

strCustCode = UCase(Trim(cbCustCode.Text))
strCustPN = UCase$(Trim$(cbCustPN.Text))
strDies = Trim$(txtDies.Text)
strSecCode = UCase(Trim$(txtSecCode.Text))

If strCustPN = "" Then
    MsgBox "��������Ҫ�����Ŀͻ�����", vbInformation, "��ʾ" '
    Exit Sub

End If

If strDies = "" Then
    MsgBox "������GROSSDIE����", vbInformation, "��ʾ" '
    Exit Sub

End If

If strSecCode = "" Then
    MsgBox "�������������", vbInformation, "��ʾ" '
    Exit Sub

End If

strSql = "select customername as �ͻ�����,CustomerPT as �ͻ�����,DIEQTY as GROSSDIES, GCVERSION as �����������λ, CREATEDBY as �½���Ա,CREATEDDATE as �½�ʱ��,LASTUPDATEBY as ������Ա,LASTUPDATEDATE as ����ʱ��  from tblCustomerDieQty where customername = 'GC' and CustomerPT = '" & strCustPN & "'  "
Set rs = Get_OracleRs(strSql)

If rs.EOF Then

    With fps
        .MaxRows = 0
        Set .DataSource = rs

    End With

    MsgBox "�ÿͻ�����û��ά����¼,�����޸�;ֻ������", vbInformation, "��ʾ"
Else
    strSql = "update tblCustomerDieQty set DIEQTY = '" & strDies & "', GCVERSION= '" & strSecCode & "',LASTUPDATEBY = '" & gUserName & "',LASTUPDATEDATE = sysdate where  CustomerPT = '" & strCustPN & "' "
    AddSql (strSql)
    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"
    Call cmdRead_Click

End If

Set rs = Nothing

End Sub

Private Sub Form_Load()
Call initCB_CustCode
cbCustCode.Text = "GC"

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       initCB_CustCode
' Description:       ��ʼ���ͻ�����
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-7-2-11:50:54
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub initCB_CustCode()
Dim rs     As New ADODB.Recordset
Dim strSql As String

strSql = "select distinct �ͻ����� from erpdata..tblxcustomer where �ͻ����� is not null"
Set rs = Get_SqlserveRs(strSql)
cbCustCode.Clear

If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbCustCode.AddItem Trim(rs("�ͻ�����"))
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       cbCustCode_LostFocus
' Description:       �ͻ�����ת��д
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-10:27:48
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbCustCode_LostFocus()
cbCustCode.Text = UCase(cbCustCode.Text)

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       cbCustCode_Change
' Description:       �ͻ�����ı�����ͻ�����/���ڻ����б�,���lstLotID
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-10:22:50
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbCustCode_Change()
Dim rs          As New ADODB.Recordset
Dim strSql      As String
Dim strCustCode As String

strCustCode = UCase(Trim$(cbCustCode.Text))
strSql = "select distinct customerptno1  from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and customerptno1 is not null"
Set rs = Get_OracleRs(strSql)
cbCustPN.Clear

If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbCustPN.AddItem Trim(rs("customerptno1"))
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       cbCustCode_DropDown
' Description:       �ͻ������������ͻ�����/���ڻ����б�,���lstLotID
' Created by :       Project Administrator
' Machine    :       1-DAC5D958B04B4
' Date-Time  :       2019-6-20-10:23:42
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cbCustCode_Click()
Dim rs          As New ADODB.Recordset
Dim strSql      As String
Dim strCustCode As String

strCustCode = UCase(Trim$(cbCustCode.Text))
strSql = "select distinct customerptno1  from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and customerptno1 is not null"
Set rs = Get_OracleRs(strSql)
cbCustPN.Clear

If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        cbCustPN.AddItem Trim(rs("customerptno1"))
        rs.MoveNext
    Loop

End If

Set rs = Nothing

End Sub
