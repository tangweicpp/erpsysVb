VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmMarkCodeRep 
   BackColor       =   &H8000000B&
   Caption         =   "���ִ�������ά��ϵͳ"
   ClientHeight    =   12270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17415
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
   ScaleHeight     =   12270
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   13095
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   17415
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5280
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "��������(NPI)"
         Height          =   195
         Left            =   2160
         TabIndex        =   20
         Top             =   2640
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��������(IT)"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   13920
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton btnUpload 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4680
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":0000
         Picture         =   "FrmMarkCodeRep.frx":3072
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "����PDSԭ�����ͼ"
         Top             =   1800
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton btnExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�˳�"
         Height          =   615
         Left            =   3600
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":5814
         Picture         =   "FrmMarkCodeRep.frx":8886
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "����"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnExport 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3000
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":B028
         Picture         =   "FrmMarkCodeRep.frx":E09A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "����"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnImport 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2400
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":101C4
         Picture         =   "FrmMarkCodeRep.frx":13236
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "����"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnDelete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ɾ��"
         Height          =   615
         Left            =   1800
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":14F30
         Picture         =   "FrmMarkCodeRep.frx":17FA2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "����"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnModify 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�޸�"
         Height          =   615
         Left            =   1200
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":1A744
         Picture         =   "FrmMarkCodeRep.frx":1D7B6
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "����"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnNew 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   615
         Left            =   600
         MaskColor       =   &H008080FF&
         MouseIcon       =   "FrmMarkCodeRep.frx":20828
         Picture         =   "FrmMarkCodeRep.frx":2389A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "����"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtPDS 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtHTPN 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton btnQuery 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ѯ"
         Height          =   615
         Left            =   0
         MaskColor       =   &H008080FF&
         Picture         =   "FrmMarkCodeRep.frx":259C4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "��ѯ�������볧�ڻ��֣����ѯ���л���"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "��̬����****��ʾ,���з���\\��ʾ,�̶��벻��"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00FFC0FF&
         Height          =   1005
         Left            =   6600
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         ToolTipText     =   "�����������������,���һ�е�ABȡֵ��Դ,DATECODE�Ķ������Ϣ"
         Top             =   1080
         Width           =   9495
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   8535
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   16095
         _Version        =   524288
         _ExtentX        =   28390
         _ExtentY        =   15055
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
         MaxCols         =   10
         MaxRows         =   0
         SpreadDesigner  =   "FrmMarkCodeRep.frx":28A36
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڻ���"
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
         Left            =   600
         TabIndex        =   9
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʶ����"
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
         Index           =   1
         Left            =   810
         TabIndex        =   8
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
         Index           =   2
         Left            =   6120
         TabIndex        =   7
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PDS"
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
         Index           =   3
         Left            =   960
         TabIndex        =   6
         Top             =   1800
         Width           =   435
      End
   End
End
Attribute VB_Name = "FrmMarkCodeRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_MK

    E_ID = 1
    E_HTPN
    E_CREATE_DATE
    E_CREATE_BY
    E_ESTABLISH_FLAG
    E_BUY_OFF_FLAG
    E_REMARK
    E_DESCRIBE
    E_END

End Enum

Dim strTitle   As String
Dim strContent As String
Dim strPDSPath As String
'
'Private Sub btnExit_Click()
'Unload Me
'End Sub
'
'Private Sub btnExport_Click()
'Call ExporToExcel("select a.cust_code �ͻ�����,a.HT_PN  ���ڻ���,a.DEFINED_FLAG �Ƿ���,a.CARD_CONTROL_FLAG �Ƿ�����, " & _
'"a.CARD_CONTROL_FLAG �Ƿ���,a.BUY_OFF_FLAG �Ƿ�����,a.CARD_CONTROL_FLAG �Ƿ񿨿�, " & _
'"a.DESCRIBE ����������,a.LINES_CNT ���������,a.CHAR_CNT �����λ��,a.REMARK ��ע from tbl_markingcode_rep a  " & _
'" order by a.cust_code,a.HT_PN ")
'End Sub

Private Sub btnDelete_Click()
Dim strHTPN As String
Dim strID   As String
Dim strSql  As String

'���ڻ���
If Len(Trim(txtHTPN.Text)) = 0 Then
    MsgBox "��������Ҫɾ�����ĳ��ڻ���", vbInformation, "��ʾ"
    Exit Sub

End If

strHTPN = Trim$(txtHTPN.Text)

If txtID.Text = "" Then
    MsgBox "����������˫����Ҫɾ�����ֵ���", vbInformation, "��ʾ"
    Exit Sub

End If

If MsgBox("���ٴ�ȷ����Ҫɾ���Ĵ�����" & vbCrLf & "ɾ������ ��(Y),�������� ��(N) ", vbYesNo, "ɾ��ȷ��???") = vbNo Then
    Exit Sub

End If

strID = Trim$(txtID.Text)
strSql = "delete from tbl_markingcode_rep where HT_PN = '" & strHTPN & "' and id = " & strID & "  and  ESTABLISH_FLAG = 'N' and BUY_OFF_FLAG = 'N' "

If AddSql(strSql) > 0 Then
    strTitle = strHTPN & "����������ɾ��"
    strContent = "NPI:" & gUserRealName & " " & strHTPN & "�Ĵ���������ɾ��"
    Call SentEml(strTitle, strContent)
    MsgBox "����:" & strHTPN & "����������ɾ��,��֪Ϥ", vbInformation, "��ʾ"
Else
    MsgBox "ɾ��ʧ��", vbInformation, "��ʾ"

End If

Call ClearTextBox

End Sub

Private Sub btnExit_Click()
Unload Me

End Sub

Private Sub btnExport_Click()
ExporToExcel ("select ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE from tbl_markingcode_rep  order by HT_PN")
End Sub

Private Sub btnModify_Click()
Dim strHTPN   As String
Dim strRemark As String
Dim strDesc   As String
Dim strPDS    As String
Dim strSql    As String
Dim lID       As Long

'���ڻ���
If Len(Trim(txtHTPN.Text)) = 0 Then
    MsgBox "��������Ҫ�������ĳ��ڻ���", vbInformation, "��ʾ"
    Exit Sub

End If

strHTPN = Trim$(txtHTPN.Text)

If Get_OracleCnt("select * from tbltsvnpiproduct where qtechptno = '" & strHTPN & "'") = 0 Then
    MsgBox "NPI���ձ�δά���ó��ڻ���: " & strHTPN & vbCrLf & "��������ȷ�ĳ��ڻ���", vbCritical, "����"
    Exit Sub

End If

If Get_OracleCnt("select * from tbl_markingcode_rep where HT_PN = '" & strHTPN & "'") = 0 Then
    MsgBox "�ó��ڻ���: " & strHTPN & "δ�����������,�޷��޸Ĺ���" & vbCrLf & "����������,��ִ����������", vbCritical, "����"
    Exit Sub

End If

'ʶ����
If Len(Trim(txtRemark.Text)) = 0 Then
    MsgBox "������������ʶ����", vbInformation, "��ʾ"
    Exit Sub

End If

strRemark = Trim$(txtRemark.Text)

'PDS
If Len(Trim(txtPDS.Text)) = 0 Then
    MsgBox "���ϴ��������PDSԭ�����ͼ", vbInformation, "��ʾ"
    Exit Sub

End If

strPDS = Trim$(txtPDS.Text)

'����
If Len(Trim(txtDesc.Text)) = 0 Then
    MsgBox "�������������������", vbInformation, "��ʾ"
    Exit Sub

End If

strDesc = Trim$(txtDesc.Text)

If MsgBox("���ٴ�ȷ���޸ĺ�Ĵ�������Ϣ�Ƿ�׼ȷ" & vbCrLf & "�޸����� ��(Y),�������� ��(N) ", vbYesNo, "�޸�ȷ��???") = vbNo Then
    Exit Sub

End If

strSql = "update tbl_markingcode_rep set UPDATE_DATE = sysdate,ESTABLISH_FLAG = 'N',BUY_OFF_FLAG= 'N',UPDATE_BY = '" & gUserName & "' || '" & gUserRealName & "' ,REMARK = '" & strRemark & "',DESCRIBE = '" & strDesc & "' where HT_PN = '" & strHTPN & "' and instr(CREATE_BY,'" & gUserName & "') > 0 "

If AddSql(strSql) > 0 Then
    strTitle = strHTPN & "�����������޸�"
    strContent = "NPI:" & gUserRealName & " " & strHTPN & "�Ĵ����������޸�,��IT�������½����ù���"
    Call SentEml(strTitle, strContent)
    MsgBox "���ڻ���:" & strHTPN & "�Ĵ��������Ѿ��޸�,����ʼ��ѷ�����IT����,��ȴ����½���", vbInformation, "�����޸ĳɹ�"
    strPDSPath = "\\10.160.1.84\public\FileServer\36.���ִ����PDS�ֿ�\" & strHTPN

    If Dir(strPDSPath, vbDirectory) = "" Then
        MkDir strPDSPath

    End If

    Call CopyFileToFtp(txtPDS.Text, strPDSPath & "\")
Else
    MsgBox "�޸�ʧ��", vbInformation, "��ʾ"

End If

Call QueryData
Call ClearTextBox

End Sub

Private Sub btnNew_Click()
Dim strHTPN   As String
Dim strRemark As String
Dim strDesc   As String
Dim strPDS    As String
Dim strSql    As String
Dim lID       As Long

'���ڻ���
If Len(Trim(txtHTPN.Text)) = 0 Then
    MsgBox "��������Ҫ�������ĳ��ڻ���", vbInformation, "��ʾ"
    Exit Sub

End If

strHTPN = Trim$(txtHTPN.Text)

If Get_OracleCnt("select * from tbltsvnpiproduct where qtechptno = '" & strHTPN & "'") = 0 Then
    MsgBox "NPI���ձ�δά���ó��ڻ���: " & strHTPN & vbCrLf & "��������ȷ�ĳ��ڻ���", vbCritical, "����"
    Exit Sub

End If

If Get_OracleCnt("select * from tbl_markingcode_rep where HT_PN = '" & strHTPN & "'") > 0 Then
    MsgBox "�ó��ڻ���: " & strHTPN & "�������������,�޷���������" & vbCrLf & "��������,��ִ���޸Ĺ���", vbCritical, "����"
    Exit Sub

End If

'ʶ����
If Len(Trim(txtRemark.Text)) = 0 Then
    MsgBox "������������ʶ����", vbInformation, "��ʾ"
    Exit Sub

End If

strRemark = Trim$(txtRemark.Text)

'PDS
If Len(Trim(txtPDS.Text)) = 0 Then
    MsgBox "���ϴ��������PDSԭ�����ͼ", vbInformation, "��ʾ"
    Exit Sub

End If

strPDS = Trim$(txtPDS.Text)

'����
If Len(Trim(txtDesc.Text)) = 0 Then
    MsgBox "�������������������", vbInformation, "��ʾ"
    Exit Sub

End If

strDesc = Trim$(txtDesc.Text)

If MsgBox("���ٴ�ȷ�ϴ�������Ϣ�Ƿ�׼ȷ" & vbCrLf & "�������� ��(Y),�������� ��(N) ", vbYesNo, "����ȷ��???") = vbNo Then
    Exit Sub

End If

lID = Get_OracleNo("select max(ID) + 1 from tbl_markingcode_rep")
strSql = "insert into tbl_markingcode_rep(ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE) values(" & lID & ",'" & strHTPN & "',sysdate,'" & gUserName & "' || '" & gUserRealName & "','N','N','" & strRemark & "','" & strDesc & "')  "

If AddSql(strSql) > 0 Then
    strTitle = strHTPN & "��������������"
    strContent = "NPI:" & gUserRealName & " " & strHTPN & "�Ĵ�������������,��IT���콨���ù���"
    Call SentEml(strTitle, strContent)
    MsgBox "���ڻ���:" & strHTPN & "�Ĵ��������Ѿ�����,����ʼ��ѷ�����IT����,��ȴ�����", vbInformation, "���뽨���ɹ�"
    strPDSPath = "\\10.160.1.84\public\FileServer\36.���ִ����PDS�ֿ�\" & strHTPN

    If Dir(strPDSPath, vbDirectory) = "" Then
        MkDir strPDSPath

    End If

    Call CopyFileToFtp(txtPDS.Text, strPDSPath & "\")

End If

Call QueryData
Call ClearTextBox

End Sub

Private Sub btnQuery_Click()
Call QueryData

End Sub

Private Sub QueryData()
Dim strHTPN As String
Dim strSql  As String
Dim rs      As New ADODB.Recordset
strHTPN = Trim$("" & txtHTPN.Text)
fps.MaxRows = 0

If strHTPN <> "" Then
    strSql = "select ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE from tbl_markingcode_rep  where HT_PN = '" & strHTPN & "' "
Else
    strSql = "select ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE from tbl_markingcode_rep  order by HT_PN"

End If

Set rs = Get_OracleRs(strSql)

If Not rs.EOF Then

    With fps
        Set .DataSource = rs

    End With

Else
    MsgBox "��ѯ������Ϣ", vbInformation
    Exit Sub

End If

End Sub

Private Sub btnUpload_Click()
CommonDialog1.Filter = "�����ļ�(*.*)|*.*|Excel�ļ�(*.xls;*.xlsx)|*.xls;*.xlsx"
CommonDialog1.ShowOpen

If CommonDialog1.filename = "" Then
    Exit Sub

End If

txtPDS.Text = CommonDialog1.filename
CommonDialog1.filename = ""

If txtPDS.Text = "" Then
    MsgBox "��ѡ��Ҫ�ϴ����ļ�", vbInformation, "��ʾ"
    Exit Sub

End If

End Sub

Private Sub Form_Load()
Call InitCtrls

End Sub

''--------------------------------------------------------------------------------
'' Project    :       ��ʽ����1
'' Procedure  :       InitCtrls
'' Description:       ��ʼ���ؼ�
'' Created by :       Project Administrator
'' Machine    :       0-354AD8C194ED4
'' Date-Time  :       2019-11-29-11:40:10
''
'' Parameters :
''--------------------------------------------------------------------------------
Private Sub InitCtrls()
'cbIsDefined.AddItem ("Y")
'cbIsDefined.AddItem ("N")
'cbIsApply.AddItem ("Y")
'cbIsApply.AddItem ("N")
'cbIsEstablish.AddItem ("Y")
'cbIsEstablish.AddItem ("N")
'cbIsBuyOff.AddItem ("Y")
'cbIsBuyOff.AddItem ("N")
'cbIsChecked.AddItem ("Y")
'cbIsChecked.AddItem ("N")
Call InitFps

End Sub

'
''--------------------------------------------------------------------------------
'' Project    :       ��ʽ����1
'' Procedure  :       InitFps
'' Description:       ��ʼ��fps
'' Created by :       Project Administrator
'' Machine    :       0-354AD8C194ED4
'' Date-Time  :       2019-11-29-11:55:48
''
'' Parameters :
''--------------------------------------------------------------------------------
Private Sub InitFps()

With fps
    .MaxCols = E_MK.E_END - 1
    .Col = -1
    .Row = -1
    .TypeMaxEditLen = 500
    .Lock = True
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
    .SetText E_MK.E_ID, 0, "���"
    .ColWidth(E_MK.E_ID) = 4
    .SetText E_MK.E_HTPN, 0, "���ڻ���"
    .ColWidth(E_MK.E_HTPN) = 8
    .SetText E_MK.E_CREATE_DATE, 0, "��������"
    .ColWidth(E_MK.E_CREATE_DATE) = 8
    .SetText E_MK.E_CREATE_BY, 0, "������"
    .ColWidth(E_MK.E_CREATE_BY) = 8
    .SetText E_MK.E_ESTABLISH_FLAG, 0, "�Ƿ���"
    .ColWidth(E_MK.E_ESTABLISH_FLAG) = 8
    .SetText E_MK.E_BUY_OFF_FLAG, 0, "�Ƿ�����"
    .ColWidth(E_MK.E_BUY_OFF_FLAG) = 8
    .SetText E_MK.E_REMARK, 0, "ʶ����"
    .ColWidth(E_MK.E_REMARK) = 15
    .SetText E_MK.E_DESCRIBE, 0, "����"
    .ColWidth(E_MK.E_DESCRIBE) = 60

End With

End Sub

Private Sub fps_Click(ByVal Col As Long, ByVal Row As Long)
Dim strID  As String
Dim strSql As String
Dim rs     As New ADODB.Recordset

If Row = 0 Then Exit Sub

With fps
    .Row = Row
    .Col = 1
    strID = .Text

End With

txtID.Text = strID
strSql = "select HT_PN,REMARK,DESCRIBE from tbl_markingcode_rep where ID = " & strID & " "
Set rs = Get_OracleRs(strSql)
txtHTPN.Text = Trim$("" & rs!HT_PN)
txtRemark.Text = Trim$("" & rs!REMARK)
txtDesc.Text = Trim$("" & rs!DESCRIBE)

End Sub

Private Sub Fps_RightClick(ByVal ClickType As Integer, _
                           ByVal Col As Long, _
                           ByVal Row As Long, _
                           ByVal MouseX As Long, _
                           ByVal MouseY As Long)
Dim strHTPN     As String
Dim strCreateBy As String

With fps
    .Row = Row
    .Col = 2
    strHTPN = Trim$("" & .Text)
    .Col = 4
    strCreateBy = Trim$("" & .Text)

End With

If gUserName = "07885" And Option1.Value = True Then
    If MsgBox("�Ƿ��Ѿ�������" & strHTPN & "�Ĵ�������,ѡ����(Y)������֪ͨ�ʼ�����ӦNPI����,��(N)�˳�", vbYesNo, "�ѽ���???") = vbNo Then
        Exit Sub

    End If

    If AddSql("update tbl_markingcode_rep set ESTABLISH_FLAG = 'Y' where HT_PN = '" & strHTPN & "' and ESTABLISH_FLAG = 'N' ") > 0 Then
        strTitle = strHTPN & "���������ѽ���"
        strContent = "IT:" & gUserRealName & " " & strHTPN & "�Ĵ��������Ѿ�����,��NPI�����������"
        Call SentEml(strTitle, strContent)
        MsgBox "�ʼ���֪ͨ��ӦNPI����,���ڻ���" & strHTPN & "���������ѽ���", vbInformation, "��ʾ"

    End If

    Call Option1_Click
ElseIf gUserName <> "07885" And Option2.Value = True Then

    If InStr(strCreateBy, gUserName) = 0 Then
        MsgBox "����˺�ֻ��BuyOff�Լ�����Ĵ����", vbCritical, "����"
        Exit Sub

    End If

    If MsgBox("�Ƿ��Ѿ�����" & strHTPN & "�Ĵ�������,ѡ����(Y)������֪ͨ�ʼ�����ӦIT����,��(N)�˳�", vbYesNo, "������???") = vbNo Then
        Exit Sub

    End If

    If AddSql("update tbl_markingcode_rep set BUY_OFF_FLAG = 'Y' where HT_PN = '" & strHTPN & "' and ESTABLISH_FLAG = 'Y' and  BUY_OFF_FLAG = 'N'") > 0 Then
        strTitle = strHTPN & "��������������"
        strContent = "NPI:" & gUserRealName & " " & strHTPN & "�Ĵ��������Ѿ�����,��IT֪Ϥ"
        Call SentEml(strTitle, strContent)
        MsgBox "�ʼ���֪ͨ��ӦIT����,���ڻ���" & strHTPN & "��������������", vbInformation, "��ʾ"

    End If

    Call Option2_Click

End If

End Sub

Private Sub Option1_Click()
Dim strSql As String
Dim rs     As New ADODB.Recordset
fps.MaxRows = 0

If Option1.Value = True Then
    strSql = "select ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE from tbl_markingcode_rep  where ESTABLISH_FLAG = 'N' and BUY_OFF_FLAG = 'N' "
    Set rs = Get_OracleRs(strSql)

    If Not rs.EOF Then

        With fps
            Set .DataSource = rs

        End With

    Else
        MsgBox "��ѯ������������", vbInformation, "��ʾ"
        Exit Sub

    End If

End If

End Sub

Private Sub Option2_Click()
Dim strSql As String
Dim rs     As New ADODB.Recordset
fps.MaxRows = 0

If Option2.Value = True Then
    strSql = "select ID,HT_PN,CREATE_DATE,CREATE_BY,ESTABLISH_FLAG,BUY_OFF_FLAG,REMARK,DESCRIBE from tbl_markingcode_rep  where ESTABLISH_FLAG = 'Y' and BUY_OFF_FLAG = 'N'  "
    Set rs = Get_OracleRs(strSql)

    If Not rs.EOF Then

        With fps
            Set .DataSource = rs

        End With

    Else
        MsgBox "��ѯ������������", vbInformation, "��ʾ"
        Exit Sub

    End If

End If

End Sub

Private Sub ClearTextBox()
Dim obj As Object

For Each obj In Me.Controls

    If TYPENAME(obj) = "TextBox" Then
        obj.Text = ""

    End If

Next

End Sub

Private Function SentEml(strSentTitle As String, strSentText As String) As Boolean
Dim i             As Integer
Dim dirtemp       As String
Dim strSentTo(50) As String
Dim strSentCC(10) As String
i = 0
dirtemp = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\�ʼ�����\SentTo_MarkingCode.cfg"
Open dirtemp For Input As #1

Do While Not EOF(1)
    Line Input #1, strTemp
    strSentTo(i) = Trim$(strTemp)
    i = i + 1
Loop
Close #1

If SentMes(strSentTitle, strSentText, strSentTo, txtPDS.Text, strSentCC) = True Then
    SentEml = True
Else
    SentEml = False

End If

End Function

Private Function SentMes(Subject As String, _
                         SentText As String, _
                         Recipient() As String, _
                         Attachment As String, _
                         RecipientCC() As String) As Boolean
Dim JM As Object
Set JM = CreateObject("JMAIL.Message")
Dim Recipients()   As String
Dim RecipientCCs() As String
Dim strBodyinfo    As String
Dim i              As Integer
Dim strSql         As String
Dim j              As Integer
Dim rs             As New ADODB.Recordset
Dim RsD            As New ADODB.Recordset

On Error GoTo ErrHandler

SentMes = False
JM.Charset = "gb2312"
JM.Silent = False
JM.Priority = 1
JM.MailServerUserName = "sqladmin" '�ʺ�
JM.MailServerPassWord = "ksitadmin" '����
JM.From = "sqladmin@ht-tech.com"    '����
JM.FromName = "sqladmin"  '����������

'�ռ���
For i = 0 To UBound(Recipient) - 1

    If Recipient(i) <> "" Then
        JM.AddRecipient Recipient(i)

    End If

Next

'������
For i = 0 To UBound(RecipientCC) - 1

    If RecipientCC(i) <> "" Then
        JM.AddRecipientCC RecipientCC(i)

    End If

Next

'����
If Attachment <> "" Then
    If Dir(Attachment, vbNormal Or vbArchive) = "" Then
        Exit Function
    Else
        JM.AddAttachment Attachment

    End If

End If

JM.Subject = Subject
JM.AppendText SentText
SentMes = JM.Send("mail.ht-tech.com")
ErrHandler:
Set JM = Nothing
Exit Function

End Function
