VERSION 5.00
Begin VB.Form DialogDNDel 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ɾ��DN��¼"
   ClientHeight    =   1635
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtDNDel 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DN"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   3
      Top             =   450
      Width           =   210
   End
End
Attribute VB_Name = "DialogDNDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub OKButton_Click()

Dim strDN  As String
Dim strSql As String
Dim rs     As New ADODB.Recordset
Dim i      As Integer

strDN = UCase(Trim(txtDNDel.Text))
If Len(strDN) = 0 Then
    MsgBox "������Ҫɾ����DN", vbInformation, "��ʾ"
    txtDNDel.Text = ""
    Exit Sub

End If

strDN = Replace$(strDN, "I", "")

strSql = "select * from packing_detailed where dn_num = '" & strDN & "'"
If Get_OracleCnt(strSql) = 0 Then
    MsgBox "��ѯ������DN����Ϣ, ����ɾ��", vbInformation, "��ʾ"
    txtDNDel.Text = ""
    Exit Sub

End If

If MsgBox("ȷ��Ҫɾ����?", vbYesNoCancel, "��ʾ") = vbNo Then
    Exit Sub

End If

' 1.ɾ����ż�¼
Call DelToErp(strDN)
' 2.ɾ��PACKING_DETAILED
strSql = "insert into packing_detailed_bak select * from packing_detailed where dn_num = '" & strDN & "'"
If AddSql(strSql) Then
    MsgBox "�ѱ���DN����", vbInformation, "��ʾ"

End If

strSql = "delete from packing_detailed where dn_num = '" & strDN & "'"
AddSql (strSql)
strSql = "delete from PKGIDSEQ_37 where dn = '" & strDN & "'  "
AddSql (strSql)
MsgBox "��ɾ��DN����", vbInformation, "��ʾ"
strSql = "update PRINT_37FLAG set printed = '0', combined = '0', scaned = '0' where dn = '" & strDN & "'"
AddSql (strSql)

Unload Me

End Sub
