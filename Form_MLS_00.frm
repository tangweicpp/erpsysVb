VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form_MLS_00 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "WLP���"
   ClientHeight    =   11160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17865
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
   ScaleHeight     =   11160
   ScaleWidth      =   17865
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame ��ǰ���� 
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17895
      Begin VB.TextBox txtLBD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   12240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   5280
         Width           =   2895
      End
      Begin VB.CheckBox chkCheck1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   13320
         TabIndex        =   21
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�˳�"
         Height          =   600
         Left            =   3840
         TabIndex        =   20
         Top             =   9720
         Width           =   1335
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "�����ȶԼ�¼"
         Height          =   600
         Left            =   1440
         TabIndex        =   19
         Top             =   9720
         Width           =   1335
      End
      Begin VB.TextBox txtBoxIDHistory 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   8280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   5280
         Width           =   2895
      End
      Begin VB.TextBox txtBoxCode 
         Height          =   2295
         Left            =   8280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txtBoxQty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   11400
         TabIndex        =   14
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtBoxID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   8280
         TabIndex        =   13
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtCartonQty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtCartonCode 
         Height          =   2295
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2400
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form_MLS_00.frx":0000
         Left            =   1440
         List            =   "Form_MLS_00.frx":000D
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtCartonID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtScan 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   15015
      End
      Begin VB.TextBox txtCartonIDHistory 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   5400
         Width           =   2895
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   15360
         TabIndex        =   18
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   2143
         _cy             =   873
      End
      Begin VB.Label lbl12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʷ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7080
         TabIndex        =   16
         Top             =   6600
         Width           =   960
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7080
         TabIndex        =   12
         Top             =   1987
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʷ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   6600
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   1987
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ɨ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   690
         Width           =   720
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʷ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   7680
      TabIndex        =   7
      Top             =   3480
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   7680
      TabIndex        =   6
      Top             =   1980
      Width           =   1080
   End
End
Attribute VB_Name = "Form_MLS_00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlag As Long

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdExport_Click()
ExporToExcel ("select TYPE  ""���"" ,id  ""���"", res  ""�ȶԽ��"",qty  ""����"", create_by  ""�˶���Ա"", create_date from MLS_00 order by create_date desc")

End Sub

Private Sub Form_Activate()
txtScan.SetFocus

End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0

End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then
    Exit Sub

End If

If txtScan.Text = "" Then
    Exit Sub

End If

Select Case Combo1.ListIndex

    Case 0  ' ����
        If Me.Caption = "GD108" Then
            Call DoCarton_GD108(UCase$(Trim$(txtScan.Text)))
        Else
            Call DoCarton(UCase$(Trim$(txtScan.Text)))

        End If

    Case 1  ' ����
        If Me.Caption = "GD108" Then
            Call DoBox_GD108(UCase$(Trim$(txtScan.Text)))
        Else
            Call DoBox(UCase$(Trim$(txtScan.Text)))

        End If

End Select

txtScan.Text = ""

End Sub

Private Sub DoCarton(strCode As String)
Dim iCnt As Integer
Dim i    As Integer
Dim sCode

strCode = "<" & strCode & ">"
If InStr(txtCartonCode, strCode) Then
    Play ("�����ظ�ɨ��")
    Exit Sub

End If

Play ("������ɨ��")
txtCartonCode.Text = txtCartonCode.Text & strCode & vbCrLf
' �ж��Ƿ�ɨ��6������
sCode = Split(txtCartonCode.Text, vbCrLf)
iCnt = UBound(sCode)

For i = 0 To iCnt - 1
    ' ���������
    If InStr(sCode(i), "-C") Then
        If InStr(txtCartonIDHistory.Text, Replace(Replace$(sCode(i), "<", ""), ">", "")) Then
            Play ("�������Ѿ��˶Թ�,ȷ���Ƿ�������")
            txtCartonCode.Text = Replace$(txtCartonCode.Text, sCode(i) & vbCrLf, "")
            Exit Sub
        Else
            If txtCartonID.Text <> "" Then
                Play ("������Ѿ�ȷ��, ����ɨ�������������")
                txtCartonCode.Text = Replace$(txtCartonCode.Text, sCode(i) & vbCrLf, "")
                Exit Sub
            Else
                txtCartonID.Text = Replace(Replace$(sCode(i), "<", ""), ">", "")
                txtCartonCode.Text = Replace$(txtCartonCode.Text, sCode(i), "")
                Play ("�������ȷ��")

            End If

        End If

        ' ����������
    ElseIf IsNumeric(Replace(Replace$(sCode(i), "<", ""), ">", "")) = True And Left$(sCode(i), 3) <> "<19" Then
        If txtCartonQty.Text <> "" Then
            Play ("���������Ѿ�ȷ��, ����ɨ������������")
            txtCartonCode.Text = Replace$(txtCartonCode.Text, sCode(i) & vbCrLf, "")
            Exit Sub
        Else
            txtCartonQty.Text = Replace(Replace$(sCode(i), "<", ""), ">", "")
            txtCartonCode.Text = Replace$(txtCartonCode.Text, sCode(i), "")
            Play ("����������ȷ��")

        End If

    End If

Next
If iCnt = 6 Then
    If txtCartonID.Text <> "" And txtCartonQty.Text <> "" Then
        Play ("������ɨ��, ��ɨ������")
        lFlag = Get_OracleNo("select MLSSEQ_00.NEXTVAL��from dual")
        AddSql ("insert into MLS_00(ID, QTY,FLAG_ID, CREATE_BY, CREATE_DATE, TYPE) values('" & txtCartonID.Text & "', '" & CLng(Trim(txtCartonQty.Text)) & "', '" & lFlag & "', '" & gUserName & "', sysdate, 'CARTON')")
        Combo1.ListIndex = 1
        Exit Sub
    Else
        MsgBox "��ǩɨ�����"

    End If

End If

End Sub

Private Sub DoCarton_GD108(strCode As String)
Dim iCnt As Integer
Dim i    As Integer
Dim sCode

strCode = "<" & strCode & ">"
If InStr(txtCartonCode, strCode) Then
    Play ("�����ظ�ɨ��")
    Exit Sub

End If

Play ("������ɨ��")
If InStr(strCode, "-R") Then
    MsgBox "����Ψһ��ɨ�����", vbInformation, "��ʾ"
    Exit Sub

End If

txtCartonCode.Text = txtCartonCode.Text & strCode & vbCrLf
' �ж��Ƿ�ɨ��6������
sCode = Split(txtCartonCode.Text, vbCrLf)
iCnt = UBound(sCode)

For i = 0 To iCnt - 1
    ' ���������
    If InStr(sCode(i), "-C") Then
        If InStr(txtCartonIDHistory.Text, Replace(Replace$(sCode(i), "<", ""), ">", "")) Then
            Play ("�������Ѿ��˶Թ�,ȷ���Ƿ�������")
            txtCartonCode.Text = Replace$(txtCartonCode.Text, sCode(i) & vbCrLf, "")
            Exit Sub
        Else
            If txtCartonID.Text <> "" Then
                Play ("������Ѿ�ȷ��, ����ɨ�������������")
                txtCartonCode.Text = Replace$(txtCartonCode.Text, sCode(i) & vbCrLf, "")
                Exit Sub
            Else
                txtCartonID.Text = Replace(Replace$(sCode(i), "<", ""), ">", "")
                txtCartonCode.Text = Replace$(txtCartonCode.Text, sCode(i), "")
                Play ("�������ȷ��")

            End If

        End If

        ' ����������
    ElseIf IsNumeric(Replace(Replace$(sCode(i), "<", ""), ">", "")) = True And Left$(sCode(i), 3) <> "<19" Then
        If txtCartonQty.Text = "" Then
            ' Play ("���������Ѿ�ȷ��, ����ɨ������������")
            ' txtCartonCode.Text = Replace$(txtCartonCode.Text, sCode(i) & vbCrLf, "")
            ' Exit Sub
            txtCartonQty.Text = Replace(Replace$(sCode(i), "<", ""), ">", "")
            txtCartonCode.Text = Replace$(txtCartonCode.Text, sCode(i), "")
            Play ("����������ȷ��")

        End If

    End If

Next
If iCnt = 5 Then
    If txtCartonID.Text <> "" And txtCartonQty.Text <> "" Then
        Play ("������ɨ��, ��ɨ������")
        lFlag = Get_OracleNo("select MLSSEQ_00.NEXTVAL��from dual")
        AddSql ("insert into MLS_00(ID, QTY,FLAG_ID, CREATE_BY, CREATE_DATE, TYPE) values('" & txtCartonID.Text & "', '" & CLng(Trim(txtCartonQty.Text)) & "', '" & lFlag & "', '" & gUserName & "', sysdate, 'CARTON')")
        Combo1.ListIndex = 1
        Exit Sub
    Else
        MsgBox "��ǩɨ�����"

    End If

End If

End Sub

Private Sub DoBox(strCode As String)
Dim iCnt As Integer
Dim i    As Integer
Dim sCode

strCode = "<" & strCode & ">"
If InStr(txtBoxCode, strCode) Then
    Play ("�����ظ�ɨ��")
    Exit Sub

End If

txtBoxCode.Text = txtBoxCode.Text & strCode & vbCrLf
' �ж��Ƿ�ɨ��6������
sCode = Split(txtBoxCode.Text, vbCrLf)
iCnt = UBound(sCode)

For i = 0 To iCnt - 1
    ' ���������
    If InStr(sCode(i), "-B") Then
        If InStr(txtBoxIDHistory.Text, Replace(Replace$(sCode(i), "<", ""), ">", "")) Then
            Play ("�������Ѿ��˶Թ�,ȷ���Ƿ�������")
            txtBoxCode.Text = Replace$(txtBoxCode.Text, sCode(i) & vbCrLf, "")
            Exit Sub
        Else
            If txtBoxID.Text <> "" Then
                Play ("�������ȷ��, ����ɨ�������������")
                txtBoxCode.Text = Replace$(txtBoxCode.Text, sCode(i) & vbCrLf, "")
                Exit Sub
            Else
                txtBoxID.Text = Replace(Replace$(sCode(i), "<", ""), ">", "")
                txtBoxCode.Text = Replace$(txtBoxCode.Text, sCode(i), "")

            End If

        End If

        ' ����������
    ElseIf (IsNumeric(Replace(Replace$(sCode(i), "<", ""), ">", "")) = True And Left$(sCode(i), 3) <> "<19") Then
        If txtBoxQty.Text <> "" Then
            Play ("����������ȷ��, ����ɨ������������")
            txtBoxCode.Text = Replace$(txtBoxCode.Text, sCode(i) & vbCrLf, "")
            Exit Sub
        Else
            txtBoxQty.Text = Replace(Replace$(sCode(i), "<", ""), ">", "")
            txtBoxCode.Text = Replace$(txtBoxCode.Text, sCode(i), "")
            Play ("����������ȷ��")

        End If

    Else
        If InStr(txtCartonCode.Text, sCode(i)) = 0 Then
            Play ("�����ǩ����")
            txtBoxCode.Text = Replace$(txtBoxCode.Text, sCode(i) & vbCrLf, "")
            Exit Sub

        End If

    End If

Next
' ����˶���ȷ
'MsgBox "��ǩ�Ѻ˶�"
Play ("right")
Dim lCartonQty As Long
Dim lBoxQty    As Long

If iCnt = 6 Then
    ' ���Ӽ�¼
    AddSql ("insert into MLS_00(ID, QTY,FLAG_ID, CREATE_BY, CREATE_DATE,TYPE, RES) values('" & Trim(txtBoxID.Text) & "', '" & CLng(Trim(txtBoxQty.Text)) & "', '" & lFlag & "', '" & gUserName & "', sysdate,'BOX', 'Y')")
    txtBoxIDHistory.Text = txtBoxIDHistory.Text & txtBoxID.Text & vbCrLf
    ' ��������
    lCartonQty = Get_OracleNo("select qty from MLS_00 where flag_id = '" & lFlag & "' and TYPE = 'CARTON' ")
    lBoxQty = Get_OracleNo("select sum(qty) from MLS_00 where flag_id =  '" & lFlag & "' and TYPE = 'BOX' and RES = 'Y' ")
    If lBoxQty > lCartonQty Then
        AddSql ("update MLS_00 set RES = 'N' where FLAG_ID = '" & lFlag & "' and ID = '" & Trim(txtBoxID.Text) & "'  ")
        Play ("���������ܺʹ�������, �˶Դ���, ��ȷ��")
        Combo1.ListIndex = 1
    ElseIf lBoxQty = lCartonQty Then
        AddSql ("update MLS_00 set RES = 'Y' where FLAG_ID = '" & lFlag & "' and TYPE = 'CARTON'  ")
        Play ("��������ȫ���˶�, ����������һ��, ��˶���һ������")
        txtCartonIDHistory.Text = txtCartonIDHistory.Text & txtCartonID.Text & vbCrLf
        txtCartonID.Text = ""
        txtCartonQty.Text = ""
        txtCartonCode.Text = ""
        txtBoxIDHistory.Text = ""
        Combo1.ListIndex = 0
    Else
        Play ("��������ɨ��, ��ɨ����һ������")
        Combo1.ListIndex = 1

    End If

    txtBoxID.Text = ""
    txtBoxQty.Text = ""
    txtBoxCode.Text = ""
    Exit Sub

End If

End Sub

Private Sub DoBox_GD108(strCode As String)
Dim iCnt As Integer
Dim i    As Integer
Dim sCode

If chkCheck1.Value = 1 Then
    If InStr(strCode, "-R") = 0 Or (Replace$(Trim(txtBoxID.Text), "-B", "") <> Replace$(strCode, "-R", "")) Then
        MsgBox "������Ψһ�����", vbInformation, "��ʾ"
        Exit Sub
    ElseIf InStr(txtLBD.Text, strCode) > 0 Then
        MsgBox "������Ψһ���Ѿ��˶Թ�, ��ȷ���Ƿ����", vbInformation, "��ʾ"
        Exit Sub
    Else
        txtLBD.Text = txtLBD.Text & strCode & vbCrLf
        GoTo CHECK

    End If

End If

strCode = "<" & strCode & ">"
If InStr(txtBoxCode, strCode) Then
    Play ("�����ظ�ɨ��")
    Exit Sub

End If

txtBoxCode.Text = txtBoxCode.Text & strCode & vbCrLf
' �ж��Ƿ�ɨ��6������
sCode = Split(txtBoxCode.Text, vbCrLf)
iCnt = UBound(sCode)

For i = 0 To iCnt - 1
    ' ���������
    If InStr(sCode(i), "-B") Then
        If InStr(txtBoxIDHistory.Text, Replace(Replace$(sCode(i), "<", ""), ">", "")) Then
            Play ("�������Ѿ��˶Թ�,ȷ���Ƿ�������")
            txtBoxCode.Text = Replace$(txtBoxCode.Text, sCode(i) & vbCrLf, "")
            Exit Sub
        Else
            If txtBoxID.Text <> "" Then
                Play ("�������ȷ��, ����ɨ�������������")
                txtBoxCode.Text = Replace$(txtBoxCode.Text, sCode(i) & vbCrLf, "")
                Exit Sub
            Else
                txtBoxID.Text = Replace(Replace$(sCode(i), "<", ""), ">", "")
                txtBoxCode.Text = Replace$(txtBoxCode.Text, sCode(i), "")

            End If

        End If

        ' ����������
    ElseIf (IsNumeric(Replace(Replace$(sCode(i), "<", ""), ">", "")) = True And Left$(sCode(i), 3) <> "<19") Then
        If txtBoxQty.Text = "" Then
            txtBoxQty.Text = Replace(Replace$(sCode(i), "<", ""), ">", "")
            txtBoxCode.Text = Replace$(txtBoxCode.Text, sCode(i), "")
            Play ("����������ȷ��")

        End If

    Else
        If InStr(txtCartonCode.Text, sCode(i)) = 0 Then
            Play ("�����ǩ����")
            txtBoxCode.Text = Replace$(txtBoxCode.Text, sCode(i) & vbCrLf, "")
            Exit Sub

        End If

    End If

Next
' ����˶���ȷ
'MsgBox "��ǩ�Ѻ˶�"
Play ("right")
Dim lCartonQty As Long
Dim lBoxQty    As Long

If iCnt = 5 Then
    If chkCheck1.Value = 0 Then
        Play ("������ɨ��,��ɨ������Ψһ��")
        chkCheck1.Value = 1
    Else
        ' ���Ӽ�¼
CHECK:
        AddSql ("insert into MLS_00(ID, QTY,FLAG_ID, CREATE_BY, CREATE_DATE,TYPE, RES) values('" & Trim(txtBoxID.Text) & "', '" & CLng(Trim(txtBoxQty.Text)) & "', '" & lFlag & "', '" & gUserName & "', sysdate,'BOX', 'Y')")
        txtBoxIDHistory.Text = txtBoxIDHistory.Text & txtBoxID.Text & vbCrLf
        ' ��������
        lCartonQty = Get_OracleNo("select qty from MLS_00 where flag_id = '" & lFlag & "' and TYPE = 'CARTON' ")
        lBoxQty = Get_OracleNo("select sum(qty) from MLS_00 where flag_id =  '" & lFlag & "' and TYPE = 'BOX' and RES = 'Y' ")
        If lBoxQty > lCartonQty Then
            AddSql ("update MLS_00 set RES = 'N' where FLAG_ID = '" & lFlag & "' and ID = '" & Trim(txtBoxID.Text) & "'  ")
            Play ("���������ܺʹ�������, �˶Դ���, ��ȷ��")
            Combo1.ListIndex = 1
        ElseIf lBoxQty = lCartonQty Then
            AddSql ("update MLS_00 set RES = 'Y' where FLAG_ID = '" & lFlag & "' and TYPE = 'CARTON'  ")
            Play ("��������ȫ���˶�, ����������һ��, ��˶���һ������")
            txtCartonIDHistory.Text = txtCartonIDHistory.Text & txtCartonID.Text & vbCrLf
            txtCartonID.Text = ""
            txtCartonQty.Text = ""
            txtCartonCode.Text = ""
            'txtBoxIDHistory.Text = ""
            Combo1.ListIndex = 0
        Else
            Play ("��������ɨ��, ��ɨ����һ������")
            Combo1.ListIndex = 1

        End If

        chkCheck1.Value = 0
        txtBoxID.Text = ""
        txtBoxQty.Text = ""
        txtBoxCode.Text = ""
        Exit Sub

    End If

End If

End Sub

Rem: ������Ƶ����
Private Sub Play(sFileName As String)
Dim sPath   As String
Dim sSuffix As String

sPath = "\\10.160.1.84\public\media_source\"
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix

End Sub
