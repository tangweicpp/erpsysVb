VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmSemtechLPS 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Semtech��ǩ��ӡϵͳ"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   10935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   19455
      Begin VB.CheckBox chk 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   9960
         Width           =   1455
      End
      Begin VB.TextBox txtHTLot 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   6135
         Left            =   7680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton cmdOuterPrinter 
         BackColor       =   &H00FF80FF&
         Caption         =   "����  ��ӡ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   8760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdInnerTrayPrinter 
         BackColor       =   &H0080C0FF&
         Caption         =   "�ں�/���� ��ӡ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   8760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "�˳�"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8880
         Width           =   2175
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H008080FF&
         Caption         =   "��ʼ��"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8880
         Width           =   2175
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   6135
      End
      Begin VB.TextBox txtScan 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtDN 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1155
         Width           =   2655
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   6135
         Left            =   10560
         TabIndex        =   12
         Top             =   1800
         Width           =   6135
         _Version        =   524288
         _ExtentX        =   10821
         _ExtentY        =   10821
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
         SpreadDesigner  =   "frmSemtechLPS.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lblSelMatch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ϻ˶�:"
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
         Left            =   10560
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "״̬��:"
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
         TabIndex        =   5
         Top             =   3960
         Width           =   750
      End
      Begin VB.Label lblScan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ɨ���"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   645
         Width           =   540
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN"
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   1200
         Width           =   210
      End
   End
End
Attribute VB_Name = "frmSemtechLPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Enum E_LPS
E_JobNO = 1
E_QTY
E_JobNO2
E_Qty2
E_End

End Enum

Private Sub chk_Click()

If chk.Visible = True Then
    cmdOuterPrinter.Visible = True
    cmdInnerTrayPrinter.Visible = True
End If

End Sub

Private Sub CmdExit_Click()

Unload Me
End Sub

Private Sub cmdInnerTrayPrinter_Click()

Dim rs As New ADODB.Recordset
Dim iDeviceQty As Integer
Dim sOra As String
Dim sDevice As String

If txtDN.Text = "" Then
    MsgBox "DN������Ϊ��", vbInformation
    Exit Sub
End If

sOra = "select * from ST_TR_SEQ order by seqtime, dev"
Set rs = Get_OracleRs(sOra)

If rs.BOF Then
'    PrintOPLable
    Exit Sub
End If

rs.MoveFirst
Do While Not rs.EOF

    Call InsertDB(rs)
    rs.MoveNext
Loop

Sleep (2000)

' ͬʱ��ӡ�ںо��̱�ǩ
PrintInLable

cmdOuterPrinter.Visible = True

End Sub

Private Sub cmdOuterPrinter_Click()

Dim sOra As String
Dim iMax As Integer
Dim i As Integer
Dim rs As New ADODB.Recordset

sOra = "select max(outbox_num) from PACKING_DETAILED where dn_num = '" & txtDN.Text & "'"

iMax = Get_OracleNo(sOra)

Frm_37_QboxLabel.Show
        
Frm_37_QboxLabel.Hide

For i = 1 To iMax

    sOra = "select distinct INBOX_NUM  from PACKING_DETAILED where outbox_num = '" & i & "' and dn_num = '" & txtDN.Text & "'"

    Set rs = Get_OracleRs(sOra)
    
    If Not rs.BOF Then
      
            rs.MoveFirst
            Do While Not rs.EOF
                Frm_37_QboxLabel.ComDN.Text = txtDN.Text
        
                Frm_37_QboxLabel.TxtWaferIDOut.Text = Frm_37_QboxLabel.TxtWaferIDOut.Text & rs.fields(0).Value & vbCrLf
        
                rs.MoveNext
            Loop
                
    End If
        
    Call Frm_37_QboxLabel.CmdOKOut_Click
    Sleep (2000)
    
Next

' ��ӡ�ܱ�ǩ
FrmSemtech_LablePrint.Show
FrmSemtech_LablePrint.Hide

FrmSemtech_LablePrint.cmbDN.Text = txtDN.Text

FrmSemtech_LablePrint.Opt(2).Value = True
    ' ��ѯ
    Call FrmSemtech_LablePrint.cmd_Click(0)
    
    With FrmSemtech_LablePrint.fps(0)

        For j = 1 To .MaxRows
            .Row = j
            
            .Col = 1
            .Text = 1
        Next

    End With
        
    ' ��ӡ
    Sleep (2000)
    Call FrmSemtech_LablePrint.cmd_Click(2)
    
End Sub

Private Sub PrintOutLable()








End Sub


Private Sub cmdReset_Click()
Unload Me
frmSemtechLPS.Show
End Sub

Private Sub Form_Activate()
txtScan.SetFocus
End Sub

Private Sub Form_Load()

' ��ʼ��fps
InitFps
InitTxtStatus

End Sub

Private Sub InitFps()

With fps
    .ReDraw = False
    .MaxCols = E_F_PO.E_End - 1
    .MaxRows = 0
        
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
        
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
        
    ' �����ͷ��
    .SetText E_LPS.E_JobNO, 0, "Ӧ��JOB"
    .SetText E_LPS.E_QTY, 0, "Ӧ������"
    .SetText E_LPS.E_JobNO2, 0, "����JOB"
    .SetText E_LPS.E_Qty2, 0, "��������"
          
    ' ������
    .ColWidth(1) = 12
    .ColWidth(2) = 10
    .ColWidth(3) = 12
    .ColWidth(4) = 10
    
    ' ����߶�
    .RowHeight(0) = 20
    .RowHeight(-1) = 15

    .ReDraw = True
End With

End Sub

Private Sub InitTxtStatus()

Dim iLotLen As Integer

iLotLen = (Len(txtHTLot.Text) - Len(Replace$(txtHTLot.Text, vbCrLf, ""))) / 2

txtStatus.Text = vbCrLf & iLotLen

End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)

' ɨ���������
If KeyAscii <> 13 Then
    Exit Sub
End If

txtStatus.ForeColor = vbBlue

' ץȡDN�жϸ�ֵ
InitDN

' Job�����˶�
MatchJobQty

' ���
ClearTxtScan

End Sub

Private Sub InitDN()

Dim sFsChar As String
Dim sDbChar As String
Dim sSelFsChar As String
Dim sSelDbChar As String

sFsChar = Left$(Trim(txtScan.Text), 1)
sDbChar = Left$(Trim(txtScan.Text), 2)
sSelFsChar = Mid$(Trim(txtScan.Text), 2)
sSelDbChar = Mid$(Trim(txtScan.Text), 3)
sSelFuChar = Trim$(txtScan.Text)

' DN��ȡ
If sFsChar = "I" Then
    txtDN.Text = sSelFsChar
    
    ' �ж��Ƿ�Ϸ�
    If Get_OracleCnt("select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & sSelFsChar & "'") = 0 Then
        
        MsgBox "ɨ���DN����ȷ, ��ȷ��", vbInformation
        Exit Sub
    End If
    
    ' �Ϸ�: ��ʼ��fps
    AssignFps (sSelFsChar)
    
    ' ���: ST_TR_SEQ
    ClearST_TR_SEQ
End If

End Sub

Private Sub MatchJobQty()

Dim sFsChar As String
Dim sDbChar As String
Dim sSelFsChar As String
Dim sSelDbChar As String
Dim sSelFuChar As String
Dim sSql As String
Dim rs As New ADODB.Recordset

sFsChar = Left$(Trim(txtScan.Text), 1)
sDbChar = Left$(Trim(txtScan.Text), 2)
sSelFsChar = Mid$(Trim(txtScan.Text), 2)
sSelDbChar = Mid$(Trim(txtScan.Text), 3)
sSelFuChar = Trim$(txtScan.Text)

' Job��ȡ
If sFsChar = "S" Then
    If InStr(txtHTLot.Text, sSelFsChar) Then
       ' MsgBox "�벻Ҫɨ��ͬһ�����̺�", vbInformation
        
        txtStatus.ForeColor = vbRed
        Exit Sub
    End If

    If Get_SqlserverCnt("select * from [erpdata].[dbo].TblTSV_Tray_details where TRAYQBOXNUMBER = '" & sSelFuChar & "'") = 0 Then
        MsgBox "ɨ���LotID����ȷ, ��ȷ��", vbInformation
        Exit Sub
    End If

    ' �ж��Ƿ�Lot��DN�Ƿ�ҹ�
    sSql = "select * from [erpdata].[dbo].TblTSV_Tray_details where TRAYQBOXNUMBER = '" & sSelFuChar & "'"
    Set rs = Get_SqlserveRs(sSql)
    
    If Get_OracleCnt("select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "' and batchnumber = '" & rs.fields("Customerlotid").Value & "' ") = 0 Then
        MsgBox "ɨ���LotID��DN��ƥ��, ��ȷ��", vbInformation
        Exit Sub
    Else
        txtHTLot = txtHTLot & sSelFsChar & vbCrLf
        
        ' ɨ1�����̲�1������
        sOra = "insert into ST_TR_SEQ values('" & txtDN.Text & "', '" & rs.fields("Customerlotid").Value & "', '" & rs.fields("CUSTOMERPT").Value & "', '" & rs.fields("QTY").Value & "', sysdate,'" & sSelFuChar & "' )"
        Exec_Ora (sOra)
        
    End If
    
    ' ����Fps
    AssignFps (sSelFuChar)
End If

' ����״̬���
InitTxtStatus

End Sub

Private Sub AssignFps(sSel As String)

Dim sOra As String
Dim rs As New ADODB.Recordset
Dim iLotLen As Integer
Dim irow As Integer
Dim sJobNo As String
Dim sLotQty As Long
Dim sRightQty As String
Dim sPreQty As String
Dim bPrintCheck As Boolean

iSum = 0
iSumPre = 0
bPrintCheck = True

iLotLen = (Len(txtHTLot.Text) - Len(Replace$(txtHTLot.Text, vbCrLf, ""))) / 2

If iLotLen = 0 Then
    ' ��ʼ��
    sOra = "select batchnumber, sum(quantity), '', '0' from CUSTOMERSHIPPINGUPTBL where delivery = '" & sSel & "' group by batchnumber"

    Set rs = Get_OracleRs(sOra)

    With fps
        .MaxRows = 0
        If rs.RecordCount > 0 Then
            Set .DataSource = rs
        End If
    End With

Else
    ' ����
    sOra = "select * from [erpdata].[dbo].TblTSV_Tray_details where TRAYQBOXNUMBER = '" & sSel & "'"
    
    Set rs = Get_SqlserveRs(sOra)
    
    With fps
        For irow = 1 To .MaxRows
            .Row = irow
            .Col = 1
            
            If .Text = rs.fields("Customerlotid").Value Then
                .Col = 3
                .Text = rs.fields("Customerlotid").Value
                
                .Col = 4
                .Text = Str(Val(.Text) + rs.fields("Qty").Value)
            End If
        Next
    End With

    ' �ж�����
    With fps
        For irow = 1 To .MaxRows
              .Row = irow
              .Col = 2
              
              sRightQty = .Text
              
              .Col = 4
              sPreQty = .Text
              
              If Val(sPreQty) <> Val(sRightQty) Then
                bPrintCheck = False
              End If
        Next
    End With
    
    If bPrintCheck Then
        cmdInnerTrayPrinter.Visible = True
        'cmdOuterPrinter.Visible = True
        txtScan.Locked = True
    End If
End If

End Sub

' �����ַ����������
Private Sub InsertDB(rs As ADODB.Recordset)

Dim sOra As String
Dim tData As tSTData
Dim lCnt As Long

' ����ID
sOra = "select count(outbox_num) from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "' " & _
" and outbox_num in (select nvl(max(outbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "')"
lCnt = Get_OracleNo(sOra)

If lCnt <= 107 Then
    sOra = "select * from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "'"
    If Get_OracleCnt(sOra) > 0 Then
        sOra = "select nvl(max(outbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "' "
        tData.OUTBOX_NUM = Get_OracleNo(sOra)
    Else
        sOra = " select nvl(max(outbox_num), '0') +1 from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "'"
        tData.OUTBOX_NUM = Get_OracleNo(sOra)
    End If
    
Else
    sOra = "select (nvl(max(outbox_num), '1') + 1) from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "' "
    tData.OUTBOX_NUM = Get_OracleStr(sOra)
End If

' �ں�ID
sOra = "select count(inbox_num) from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "' " & _
"and inbox_num in (select nvl(max(inbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "')"
lCnt = Get_OracleNo(sOra)

If lCnt <= 8 Then
    sOra = "select nvl(max(inbox_num), '1') from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "'"
    tData.INBOX_NUM = Get_OracleStr(sOra)
Else
    sOra = "select nvl(max(inbox_num), '1') + 1 from PACKING_DETAILED where dn_num = '" & rs.fields("DN") & "' and customer_device = '" & rs.fields("DEV") & "'  and outbox_num = '" & tData.OUTBOX_NUM & "'"
    tData.INBOX_NUM = Get_OracleStr(sOra)
End If

' ����ID
tData.TRAYID = rs.fields("LOTID")
tData.CREATE_BY = gUserName
tData.DN_NUM = rs.fields("DN")
tData.JOB_ID = rs.fields("JOB")
tData.CUSTOMER_DEVICE = rs.fields("DEV")
tData.qty = 15000

' ��������
Call insertToSql(tData)

' ��ӡ����
'PrintOPLable

End Sub


Private Sub insertToSql(tData As tSTData)

Dim sOra As String

sOra = "insert into PACKING_DETAILED values('" & tData.TRAYID & "','" & tData.INBOX_NUM & "','" & tData.OUTBOX_NUM & "','" & tData.DN_NUM & "','" & tData.JOB_ID & "','" & tData.qty & "','" & tData.CUSTOMER_DEVICE & "',sysdate,'" & tData.CREATE_BY & "','0','0','') "

Exec_Ora (sOra)

End Sub

Private Sub ClearST_TR_SEQ()

Dim sOra As String

sOra = "delete from ST_TR_SEQ"

Exec_Ora (sOra)

End Sub

Private Sub ClearTxtScan()

txtScan.Text = ""

End Sub

Private Sub PrintInLable()

' ������
Dim rs As New ADODB.Recordset
Dim sInfo As String
Dim sOra As String
Dim sAppend As String
Dim iOp As Integer
Dim iIp As Integer
Dim iOpMax As Integer
Dim iIpMax As Integer
Dim FName As String
Dim dirtemp As String

sInfo = ""
iOp = 1
iIp = 1

sOra = "select max(outbox_num) from PACKING_DETAILED where dn_num = '" & txtDN.Text & "'  "
iOpMax = Get_OracleNo(sOra)
iIpMax = 9

For iOp = 1 To iOpMax
    For iIp = 1 To iIpMax
        sAppend = "select * from PACKING_DETAILED where dn_num = '" & txtDN.Text & "' and outbox_num = '" & iOp & "' and inbox_num = '" & iIp & "' "
        Set rs = Get_OracleRs(sAppend)
        
        If Not rs.BOF Then
        
            rs.MoveFirst
            Do While Not rs.EOF
                sInfo = sInfo & rs.fields("trayid") & vbCrLf
                rs.MoveNext
            Loop
                
        End If
        
        Call PrintInbox(sInfo, txtDN.Text)
        
        ' ��һ���ںд�ӡ���
        sInfo = ""
        
        Sleep (8000)
    Next
    
Next

End Sub
