VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmWaferRK 
   Caption         =   "�Ǳ�˰��Բ����¼ά��"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13425
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
   ScaleHeight     =   10395
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Height          =   10335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      Begin VB.TextBox txtLotID 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   1020
         Width           =   2055
      End
      Begin VB.CommandButton btnDel 
         Caption         =   "ɾ��"
         Height          =   360
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   990
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   7935
         Left            =   480
         TabIndex        =   2
         Top             =   1560
         Width           =   11055
         _Version        =   524288
         _ExtentX        =   19500
         _ExtentY        =   13996
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
         MaxRows         =   0
         SpreadDesigner  =   "FrmWaferRK.frx":0000
         AppearanceStyle =   0
      End
      Begin VB.CommandButton btnQuery 
         Caption         =   "��ѯ"
         Height          =   360
         Left            =   480
         TabIndex        =   1
         ToolTipText     =   "��ѯ"
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label1 
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
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   525
      End
   End
End
Attribute VB_Name = "FrmWaferRK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnDel_Click()
Dim strInRecID As String
Dim strLotID   As String
Dim strSql     As String
Dim strEventDesc  As String
Dim strEventKey As String
Dim strEventID As String
Dim strNewID As String

With fps
    If .MaxRows = 0 Then
        MsgBox "û������¼���Ա�ɾ��", vbCritical, "����"
        Exit Sub

    End If

    .Row = 1
    .Col = 1
    strInRecID = Trim$("" & .Text)
    .Col = 2
    strLotID = Trim$("" & .Text)
    '����
    If MsgBox("�Ƿ�ȷ��ɾ������:" & strLotID & " ��ⵥ���:" & strInRecID & "������¼", vbYesNo, "����") = vbNo Then
        MsgBox "δɾ���κμ�¼", vbInformation, "��ʾ"
        Exit Sub

    End If

    '��д����
    strGReason = ""
    DlgEventReason.Show 1

    If strGReason = "" Then
        MsgBox "������д����,���򲻿�ɾ��", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    '��¼
    strEventKey = strInRecID
    strEventDesc = "ɾ���Ǳ�˰��Բ����¼"
    
    strEventID = SaveTblEventRec(E_TBL_EVENT.E_DELETE, strEventKey, strEventDesc, strGReason, "")
    If strEventID = "" Then
        MsgBox "�޸�/ɾ���¼�δ��¼,�޷�ɾ��", vbCritical, "��ʾ"
        Exit Sub
    End If
    
    '����
    strSql = "insert into ERPBASE..tbltoinrec_wafer_bak select * from ERPBASE..tbltoinrec_wafer where ��ⵥ��� = '" & strInRecID & "'  "
    AddSql2 (strSql)
    strNewID = strInRecID & "|" & strEventID
    strSql = "update ERPBASE..tbltoinrec_wafer_bak set ��ⵥ��� = '" & strNewID & "' where ��ⵥ��� = '" & strInRecID & "'  "
    
    If AddSql2(strSql) > 0 Then
        MsgBox "�����ѱ���", vbInformation, "��ʾ"

    End If

    'ɾ��
    strSql = "delete from ERPBASE..tbltoinrec_wafer where ��ⵥ��� = '" & strInRecID & "'"
    If AddSql2(strSql) > 0 Then
        MsgBox "������ɾ��", vbInformation, "��ʾ"
        
    End If
    
    Call btnQuery_Click
    
End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       SaveTblEventRec
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/11/18-11:15:16
'
' Parameters :       enumEventType (E_TBL_EVENT)
'
'E_INSERT ��
'E_DELETE ɾ
'E_UPDATE ��
'E_QUERY  ��
'                    strEventKey (String)  �ؼ���,�絥�ݱ��,����,�Ϻŵ�
'                    strEventDesc (String) �¼�����
'                    strEventReason (String) �¼�����
'                    strEventRemark (String) �¼���ע
'--------------------------------------------------------------------------------
Private Function SaveTblEventRec(enumEventType As E_TBL_EVENT, _
                                 strEventKey As String, _
                                 strEventDesc As String, _
                                 strEventReason As String, _
                                 strEventRemark As String) As String
Dim strEventID   As String
Dim strSql       As String
Dim strUserName  As String
Dim strEventType As String

Select Case enumEventType

    Case E_INSERT
        strEventType = "INSERT"

    Case E_DELETE
        strEventType = "DELETE"

    Case E_UPDATE
        strEventType = "UPDATE"

    Case E_QUERY
        strEventType = "QUERY"

End Select

strSql = "select EmpName from XTW..employee where empno = '" & gUserName & "'"
strUserName = Get_SqlStr2(strSql)
strEventID = Right("00" & Year(Now), 2) & Right("00" & Month(Now), 2) & Right$("00" & Day(Now), 2)
strEventID = strEventID & Right$("00" & Get_OracleStr("select nvl(max(EVENT_ID),0) + 1 from TBL_EVENT_RECORD where  instr(EVENT_ID,'" & strEventID & "') > 0 "), 2)
strSql = "insert into TBL_EVENT_RECORD(EVENT_ID,EVENT_TYPE,EVENT_KEY,EVENT_DESC,EVENT_REASON,USER_ID,USER_NAME,DATETIME,REMARK) values('" & strEventID & "','" & strEventType & "','" & strEventKey & "','" & strEventDesc & "','" & strEventReason & "','" & gUserName & "','" & strUserName & "',sysdate,'" & strEventRemark & "') "
If AddSql(strSql) > 0 Then
    MsgBox "�¼��Ѽ�¼", vbInformation, "��ʾ"
    SaveTblEventRec = strEventID
Else
    MsgBox "�¼�δ��¼", vbCritical, "��ʾ"
    SaveTblEventRec = ""

End If

End Function

Private Sub btnQuery_Click()
Dim strSql   As String
Dim strLotID As String
Dim rs       As New ADODB.Recordset

fps.MaxRows = 0

If txtLotID.Text = "" Then
    MsgBox "�����뾧Բ����", vbInformation, "��ʾ"
    Exit Sub

End If

strLotID = UCase(Trim$(txtLotID.Text))
strSql = "select * from ERPBASE..tblToInRec_Wafer where ���� = '" & strLotID & "'"
Set rs = Get_SqlserveRs(strSql)
If rs.RecordCount = 0 Then
    MsgBox "��ѯ�����þ�Բ���ŵ�����¼", vbCritical, "����"
    Exit Sub

End If

With fps
    .MaxRows = 0
    Set .DataSource = rs

End With

End Sub
