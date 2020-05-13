VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Frm_LblMatchSys 
   BackColor       =   &H00E0E0E0&
   Caption         =   "标签核对系统"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18195
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
   ScaleHeight     =   9165
   ScaleWidth      =   18195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "37标签比对"
      ForeColor       =   &H00FF0000&
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19935
      Begin VB.TextBox txtShipTo 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00FF8080&
         Caption         =   "导出核对记录"
         Height          =   360
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   6720
         Width           =   1455
      End
      Begin VB.OptionButton choose 
         BackColor       =   &H00C0C0C0&
         Caption         =   "二维码"
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   27
         Top             =   6120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton choose 
         BackColor       =   &H00C0C0C0&
         Caption         =   "条码枪"
         Height          =   495
         Index           =   0
         Left            =   2400
         TabIndex        =   26
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox txtCustReelCheckData 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   2295
         Left            =   5760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   6600
         Width           =   12255
      End
      Begin VB.TextBox txt37ReelID 
         BackColor       =   &H00FFC0FF&
         Height          =   2295
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtCur37ReelID 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   5280
         Width           =   2415
      End
      Begin VB.TextBox txtCurInnerBoxNum 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   4905
         Width           =   2415
      End
      Begin VB.TextBox txtMediaPath 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Text            =   "\\10.160.1.84\public\media_source\"
         Top             =   5760
         Width           =   2895
      End
      Begin VB.TextBox txtInnerBoxNum 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1545
         Width           =   4455
      End
      Begin VB.TextBox txtCurOuterBoxNum 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4560
         Width           =   2415
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FF80FF&
         Caption         =   "退出"
         Height          =   360
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6720
         Width           =   1455
      End
      Begin VB.TextBox txtInnerBoxCheckData 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   2295
         Left            =   5760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3960
         Width           =   12255
      End
      Begin VB.TextBox txtOuterBoxCheckData 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   3135
         Left            =   5760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   480
         Width           =   12255
      End
      Begin VB.TextBox txtOuterBoxNum 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1185
         Width           =   4455
      End
      Begin VB.TextBox txtDN 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   825
         Width           =   1455
      End
      Begin VB.TextBox txtScan 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lbl22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘待检:"
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
         TabIndex        =   25
         Top             =   2940
         Width           =   975
      End
      Begin VB.Label lblReel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘标签数据:"
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
         Left            =   5730
         TabIndex        =   22
         Top             =   6360
         Width           =   1425
      End
      Begin VB.Label lblReel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前卷盘:"
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
         TabIndex        =   20
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label lblIpkg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前内箱:"
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
         Top             =   4920
         Width           =   975
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   615
         Left            =   1200
         TabIndex        =   17
         Top             =   7920
         Visible         =   0   'False
         Width           =   615
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
         _cx             =   1085
         _cy             =   1085
      End
      Begin VB.Label lblPreOP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前外箱:"
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
         TabIndex        =   13
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label lblInPkgLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内箱标签数据:"
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
         Left            =   5730
         TabIndex        =   11
         Top             =   3720
         Width           =   1425
      End
      Begin VB.Label lblOutPkgLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱标签数据:"
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
         Left            =   5730
         TabIndex        =   10
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label lblInPkg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内箱待检:"
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
         TabIndex        =   7
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblOutPkg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱待检:"
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
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN:"
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
         Left            =   795
         TabIndex        =   3
         Top             =   840
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫码框:"
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
         TabIndex        =   1
         Top             =   495
         Width           =   750
      End
   End
End
Attribute VB_Name = "Frm_LblMatchSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim gOuterBoxCheckData As String
Dim gInnerBoxCheckData As String
Dim gCustReelCheckData As String
Dim gShipTo            As String

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdExport_Click()
Dim strSql As String
Dim strDN  As String

If txtDN.Text = "" Then
    MsgBox "请输入查找的DN", vbInformation, "提示"
    Exit Sub

End If

strDN = Trim$(txtDN.Text)
strSql = "select distinct a.dn_num, a.mpn, a.cpn, a.job_id, a.datecode, b.create_date as 核对时间, b.create_by as 核对人员,b.status as 核对状态,sum(a.qty) as 数量,b.ADDINFO as 备注 from lpstbl a, HWLBLMATCHHIS b where a.dn_num  = '" & strDN & "' and a.dn_num = b.dn " & "group by a.dn_num, a.mpn, a.cpn, a.job_id, a.datecode,b.create_date,b.create_by,b.status, b.ADDINFO "
Call ExporToExcel(strSql)

End Sub

Private Sub Form_Activate()
txtScan.SetFocus

End Sub

Private Sub Form_Load()
txtMediaPath.Text = "\\10.160.1.84\public\media_source\"

End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Or txtScan.Text = "" Then
    Exit Sub

End If

Dim strScan As String

strScan = UCase$(Trim$(txtScan.Text))
If txtDN.Text = "" Then
    Call InitDN(strScan)
ElseIf txtCurOuterBoxNum.Text = "" Then
    '外箱数据展开
    Call InitOP(strScan)
ElseIf txtOuterBoxCheckData.Text <> "" Then
    '匹配外箱数据
    Call MatchOP(strScan)
ElseIf txtCurInnerBoxNum.Text = "" Then
    '内盒数据展开
    Call InitIP(strScan)
ElseIf Replace(txtInnerBoxCheckData.Text, vbCrLf, "") <> "" Then
    '匹配内盒数据
    Call MatchIP(strScan)
ElseIf txtCur37ReelID.Text = "" Then
    '卷盘数据展开
    Call InitReel(strScan)
ElseIf txt37ReelID.Text <> "" Then
    '匹配卷盘数据
    Call MatchReel(strScan)
Else
    Play ("noScan")

End If

txtScan.Text = ""
txtScan.SetFocus

End Sub

Private Sub InitDN(strScan As String)
Dim strSql As String
Dim rs     As New ADODB.Recordset
Dim strDN  As String

strDN = Right$(strScan, 8)
strSql = "select dn_num from packing_detailed where dn_num = '" & strDN & "' "
txtDN.Text = Get_OracleStr(strSql)
If txtDN.Text = "" Then
    Play ("wrongDN")
    Exit Sub

End If

' 判断DN类型
gShipTo = GetDNType(txtDN.Text)
txtShipTo.Text = gShipTo
strSql = "select distinct outbox_num from packing_detailed where dn_num = '" & strDN & "' order by outbox_num "
Set rs = Get_OracleRs(strSql)
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        txtOuterBoxNum.Text = txtOuterBoxNum.Text & "<" & rs!OUTBOX_NUM & ">"
        rs.MoveNext
    Loop

End If

Play ("rightDN")
rs.Close
Set rs = Nothing

End Sub

Private Sub InitOP(strScan As String)
Dim strSql    As String
Dim rs        As New ADODB.Recordset
Dim strQrCode As String

'检查KID,QID
txtCurOuterBoxNum.Text = Get_OracleStr(" select distinct outbox_num from packing_detailed where dn_num = '" & txtDN.Text & "' and (kid = '" & strScan & "' or carton = '" & strScan & "') ")
If txtCurOuterBoxNum.Text = "" Then
    Play ("wrongCode")
    Exit Sub

End If

'检查该外箱是否已经扫描过
If InStr(txtOuterBoxNum.Text, "<" & txtCurOuterBoxNum.Text & ">") = 0 Then
    Play ("otherOP")
    txtCurOuterBoxNum.Text = ""
    Exit Sub

End If

Play ("rightOP")
'展开内箱NUM
strSql = "select distinct inbox_num from packing_detailed where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "'  order by inbox_num "
Set rs = Get_OracleRs(strSql)
If Not rs.EOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        txtInnerBoxNum.Text = txtInnerBoxNum.Text & "<" & rs!INBOX_NUM & ">"
        rs.MoveNext
    Loop

End If

'展开外箱标签数据
txtOuterBoxCheckData.Text = ""

'外箱大标签
Select Case gShipTo

    Case "HW", "ST"
        ' strSql = "select distinct '<I' || dn_num || '>' || '<K' || nvl(po, 'N/A') || '>' || '<P' || nvl(cpn, 'N/A') || '>' || '<Z' || mpn || '>' || '<Q' || sum(qty) || '>' from LPSTBL where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' group by dn_num,po,cpn,mpn "
        strSql = "select '<' || key_value || '>' from TBL37QRVALUE where key_name = '" & txtDN.Text & "' || '_K' || '" & CLng(txtCurOuterBoxNum.Text) & "'"
        strQrCode = UCase(Get_OracleStr(strSql))
        txtOuterBoxCheckData.Text = strQrCode & vbCrLf

    Case "SSE2", "SHORT"
        strSql = "select distinct '<I' || dn_num || '>' || '<K' || nvl(substr(po,1, 10), 'N/A') || '>' || '<P' || nvl(substr(cpn, 1, 11), 'N/A') || '>' || '<Z' || mpn || '>' || '<Q' || sum(qty) || '>' from LPSTBL where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' group by dn_num,po,cpn,mpn "
        txtOuterBoxCheckData.Text = UCase(Get_OracleStr(strSql)) & vbCrLf

End Select

'外箱C标签
If choose(1).Value = True Then  ' 二维码
    strSql = "select distinct cartonid from packing_detailed where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "'"
    Set rs = Get_OracleRs(strSql)
    If Not rs.EOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            strSql = "select '<' || key_value || '>' from TBL37QRVALUE where key_name = '" & txtDN.Text & "' || '_' || '" & rs!CARTONID & "'"
            strQrCode = UCase(Get_OracleStr(strSql))
            txtOuterBoxCheckData.Text = txtOuterBoxCheckData.Text & strQrCode & vbCrLf
            rs.MoveNext
        Loop

    End If

Else    '条码
    strSql = "select distinct cartonid from packing_detailed where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "'"
    Set rs = Get_OracleRs(strSql)
    If Not rs.EOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            txtOuterBoxCheckData.Text = txtOuterBoxCheckData.Text & "<" & rs!CARTONID & ">" & vbCrLf
            rs.MoveNext
        Loop

    End If

End If

gOuterBoxCheckData = txtOuterBoxCheckData.Text
rs.Close
Set rs = Nothing

End Sub

Private Sub InitIP(strScan As String)
Dim strSql As String
Dim strBID As String
Dim rs     As New ADODB.Recordset

'扫描37BID或二维码
If choose(1).Value = True Then  ' 二维码
    If Left$(strScan, 7) <> "[)>061T" Then
        Play ("wrongCode")
        Exit Sub
    End If
    
    strSql = "select key_name from TBL37QRVALUE where key_value = '" & strScan & "'"
    If InStr(Get_OracleStr(strSql), "_") = 0 Then
        Play ("wrongCode")
        Exit Sub
    End If
    
    strBID = Split(Get_OracleStr(strSql), "_")(1)
Else 'BID
    strBID = strScan

End If

'得到内箱序号
strSql = " select distinct inbox_num from LPSTBL where dn_num = '" & txtDN.Text & "' and boxid = '" & strBID & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' "
txtCurInnerBoxNum.Text = Get_OracleStr(strSql)
If txtCurInnerBoxNum.Text = "" Then
    Play ("wrongCode")
    Exit Sub

End If

Play ("rightCode")
If InStr(txtInnerBoxNum.Text, "<" & txtCurInnerBoxNum.Text & ">") = 0 Then
    Play ("otherIP")
    txtCurInnerBoxNum.Text = ""
    Exit Sub

End If

' 初始化ReelData
If choose(1).Value = True Then  ' 二维码
    strSql = "select distinct trayid  from LPSTBL where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "' "
Else
    strSql = "select distinct trayid from LPSTBL where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "' "

End If

Set rs = Get_OracleRs(strSql)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        txt37ReelID.Text = txt37ReelID.Text & rs.Fields(0) & vbCrLf
        rs.MoveNext
    Loop

End If

' 37内盒B标签
txtInnerBoxCheckData.Text = ""
If choose(1).Value = True Then  ' 二维码
    strSql = "select distinct boxid from packing_detailed where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "' "
    Set rs = Get_OracleRs(strSql)
    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            strSql = "select '<' || key_value || '>' from TBL37QRVALUE where key_name = '" & txtDN.Text & "' || '_' || '" & rs!BOXID & "'"
            strQrCode = UCase(Get_OracleStr(strSql))
            txtInnerBoxCheckData.Text = txtInnerBoxCheckData.Text & strQrCode & vbCrLf
            rs.MoveNext
        Loop

    End If

Else    '条码
    strSql = "select distinct boxid from packing_detailed where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "' "
    Set rs = Get_OracleRs(strSql)
    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            txtInnerBoxCheckData.Text = txtInnerBoxCheckData.Text & "<" & rs!BOXID & ">" & vbCrLf
            rs.MoveNext
        Loop

    End If

End If

'三星内盒标签
If gShipTo = "SSE2" Then
    strSql = "select distinct '<' ||cpn||'DPTKE2'||substr('000000' || sum(qty), -6, 6)||'>' from LPSTBL where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "' group by cpn"
    txtInnerBoxCheckData.Text = txtInnerBoxCheckData.Text & Get_OracleStr(strSql) & vbCrLf
    
ElseIf gShipTo = "SHORT" Then
    strSql = "select distinct '<' ||cpn||'DPTK'||substr('000000' || sum(qty), -6, 6)||'>' from LPSTBL where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "' group by cpn"
    txtInnerBoxCheckData.Text = txtInnerBoxCheckData.Text & Get_OracleStr(strSql) & vbCrLf

    '华为内盒标签
ElseIf gShipTo = "HW" Then
    ' 扫描二维码
    strSql = "select distinct '1P' || cpn || '1V601024' || '10D' || datecode || '1T' || job_ID || 'Q' || sum(qty) from lpstbl where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "' group by cpn,datecode,job_ID "
    Set rs = Get_OracleRs(strSql)
    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            txtInnerBoxCheckData.Text = txtInnerBoxCheckData.Text & "<" & UCase(Trim(rs.Fields(0))) & ">" & vbCrLf
            rs.MoveNext
        Loop

    End If

End If

gInnerBoxCheckData = txtInnerBoxCheckData.Text
txtInnerBoxCheckData.Text = Replace$(txtInnerBoxCheckData.Text, "<" & strScan & ">", "", , 1)
'三星卷盘标签
txtCustReelCheckData.Text = ""
If gShipTo = "SSE2" Then
    strSql = "select distinct cpn||'DPTKE2'||reelid||substr('000000' || qty, -6, 6) from LPSTBL where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "'"
    Set rs = Get_OracleRs(strSql)
    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            txtCustReelCheckData.Text = txtCustReelCheckData.Text & "<" & UCase(Trim(rs.Fields(0))) & ">" & vbCrLf
            rs.MoveNext
        Loop

    End If


ElseIf gShipTo = "SHORT" Then
     strSql = "select distinct cpn||'DPTK'||reelid||substr('000000' || qty, -6, 6) from LPSTBL where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "'"
    Set rs = Get_OracleRs(strSql)
    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            txtCustReelCheckData.Text = txtCustReelCheckData.Text & "<" & UCase(Trim(rs.Fields(0))) & ">" & vbCrLf
            rs.MoveNext
        Loop

    End If
    
    '华为卷盘标签
ElseIf gShipTo = "HW" Then
    ' 华为卷盘
    ' 二维码
    strSql = "select '52S' || reelid || '18VLEHWTF02010I' || '1P' || cpn || '1V601024' || '10D' || datecode || '1T' || job_ID || 'Q' || sum(qty) from lpstbl where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "' group by cpn,datecode,job_ID,reelid "
    Set rs = Get_OracleRs(strSql)
    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            txtCustReelCheckData.Text = txtCustReelCheckData.Text & "<" & rs.Fields(0) & ">" & vbCrLf
            rs.MoveNext
        Loop

    End If

End If

rs.Close
Set rs = Nothing

End Sub

Private Sub InitReel(strScan As String)
Dim sLbl1  As String
Dim rs     As ADODB.Recordset
Dim strRID As String
Dim strDC  As String

'二维码
If choose(1).Value = False Then
    If Left(strScan, 3) <> "[)>" Then
        Play ("wrongCode")
        Exit Sub

    End If

    strRID = Mid(strScan, InStr(strScan, "S"), InStr(strScan, "-R") - InStr(strScan, "S")) & Mid(strScan, InStr(strScan, "-R"), InStr(Mid(strScan, InStr(strScan, "-R")), "Q") - 1)
    strDC = Right$(strScan, 4)

    
    txtCur37ReelID.Text = Get_OracleStr("select trayid from packing_detailed where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "' and trayid = '" & strRID & "'")
    If txtCur37ReelID.Text = "" Then
        Play ("wrongCode")
        Exit Sub

    End If

    Play ("rightCode")
    
    If InStr(txt37ReelID.Text, strRID) = 0 Then
        Play ("该卷盘已核对, 请勿重复核对")
        txtCur37ReelID.Text = ""
        Exit Sub

    End If

Else
    txtCur37ReelID.Text = Get_OracleStr("select trayid from packing_detailed where dn_num = '" & txtDN.Text & "' and outbox_num = '" & txtCurOuterBoxNum.Text & "' and inbox_num = '" & txtCurInnerBoxNum.Text & "' and trayid = '" & strScan & "'")
    If txtCur37ReelID.Text = "" Then
        Play ("wrongCode")
        Exit Sub

    End If

    Play ("rightCode")
    If InStr(txt37ReelID.Text, txtCur37ReelID.Text) = 0 Then
        Play ("该卷盘已核对, 请勿重复核对")
        txtCur37ReelID.Text = ""
        Exit Sub

    End If

End If

txt37ReelID.Text = Replace$(txt37ReelID.Text, strScan, "", , 1)

If gShipTo = "ST" Then
    txtCur37ReelID.Text = ""
    If Replace(txt37ReelID.Text, vbCrLf, "") = "" Then
        Play ("otherIP")
        txtInnerBoxNum.Text = Replace$(txtInnerBoxNum.Text, "<" & txtCurInnerBoxNum.Text & ">", "", , 1)
        txtCurInnerBoxNum.Text = ""
        txt37ReelID.Text = ""

    End If

    If txtInnerBoxNum.Text = "" Then
        Play ("otherOP")
        txtOuterBoxNum.Text = Replace$(txtOuterBoxNum.Text, "<" & txtCurOuterBoxNum.Text & ">", "", , 1)
        txtCurOuterBoxNum.Text = ""

    End If

    If txtOuterBoxNum.Text = "" Then
        Play ("completeDN")
        ' 更新检验状态
        If txtDN.Text <> "" Then
            Call UpdateChkStatus(Trim$(txtDN.Text))

        End If

    End If

End If

End Sub

' 核对外箱标签
Private Sub MatchOP(strScan As String)
If strScan = "" Then
    Play ("wrongCode")
    Exit Sub

End If

strScan = "<" & strScan & ">"
If InStr(gOuterBoxCheckData, strScan) = 0 Then
    Play ("wrongCode")
    Exit Sub

End If

Play ("rightCode")
If InStr(txtOuterBoxCheckData.Text, strScan) = 0 Then
    Play ("repCode")
    Exit Sub

End If

txtOuterBoxCheckData.Text = Replace$(txtOuterBoxCheckData.Text, strScan, "", , 1)
If Replace(txtOuterBoxCheckData.Text, vbCrLf, "") = "" Then
    Play ("nextIP")
    txtOuterBoxCheckData.Text = ""
    gOuterBoxCheckData = ""

End If

End Sub

' 核对内盒标签
Private Sub MatchIP(strScan As String)
Dim strArr() As String

If strScan = "" Then
    Play ("wrongCode")
    Exit Sub

End If

If gShipTo = "HW" Then
    strScan = Replace$(strScan, "[)>06F01001P18VLEHWTF02010I", "")
End If

strScan = "<" & strScan & ">"
If InStr(gInnerBoxCheckData, strScan) = 0 Then
    Play ("wrongCode")
    Exit Sub

End If

Play ("rightCode")
If InStr(txtInnerBoxCheckData.Text, strScan) = 0 Then
    Play ("repCode")
    Exit Sub

End If

txtInnerBoxCheckData.Text = Replace$(txtInnerBoxCheckData.Text, strScan, "", , 1)
If Replace(txtInnerBoxCheckData.Text, vbCrLf, "") = "" Then
    Play ("nextReel")
    txtInnerBoxCheckData.Text = ""
    gInnerBoxCheckData = ""

End If

End Sub

Private Sub MatchReel(strScan As String)
Dim strJobID As String
Dim strLotID As String

If strScan = "" Then
    Play ("wrongCode")
    Exit Sub

End If

'三星卷盘
If gShipTo = "SSE2" Or gShipTo = "SHORT" Then
    '判断是否是一个内盒的
    strScan = "<" & strScan & ">"
    If InStr(txtCustReelCheckData, strScan) = 0 Then
        Play ("wrongCode")
        Exit Sub

    End If

    '判断JOB一致
    strJobID = Get_OracleStr("select job_id from lpstbl where dn_num = '" & txtDN.Text & "' and trayid = '" & txtCur37ReelID.Text & "' ")
    If Right(strJobID, 1) = "M" Then
        If InStr(strScan, strJobID) = 0 Then
            Play ("wrongJob")
            Exit Sub

        End If

    Else
        If InStr(strScan, strJobID) > 0 Then
            If InStr(strScan, strJobID & "M") > 0 Then
                Play ("wrongJob")
                Exit Sub

            End If

        Else
            Play ("wrongJob")
            Exit Sub

        End If

    End If

    Play ("otherReel")
    txtCustReelCheckData.Text = Replace$(txtCustReelCheckData.Text, strScan, "", , 1)
    '华为卷盘
ElseIf gShipTo = "HW" Then
    Dim strPSN   As String
    Dim strArr() As String

    strScan = Replace$(strScan, "[)>06F01001P", "")
    strPSN = strScan
    strScan = "<" & strScan & ">"
    If InStr(txtCustReelCheckData, strScan) = 0 Then
        Play ("wrongCode")
        Exit Sub

    End If

    'JOB一一对应
    strJobID = Get_OracleStr("select job_id from lpstbl where dn_num = '" & txtDN.Text & "'  and trayid = '" & txtCur37ReelID.Text & "' ")
    If Right(strJobID, 1) = "M" Then
        If InStr(strScan, strJobID) = 0 Then
            Play ("wrongJob")
            Exit Sub

        End If

    Else
        If InStr(strScan, strJobID) > 0 Then
            If InStr(strScan, strJobID & "M") > 0 Then
                Play ("wrongJob")
                Exit Sub

            End If

        Else
            Play ("wrongJob")
            Exit Sub

        End If

    End If

    ' Reel一一对应: PSN对应-R, 否则更新对照表
    Dim strReelIDOld As String, strReelIDNow As String, strTrayID As String
    Dim strSql       As String

    strReelIDOld = Get_OracleStr("select reelid from packing_detailed where dn_num = '" & txtDN.Text & "' and trayid = '" & txtCur37ReelID.Text & "' ")
    strReelIDNow = strPSN
    If InStr(strReelIDNow, strReelIDOld) = 0 Then
        Play ("卷盘标签不匹配")
        Exit Sub

    End If

    Play ("otherReel")
    txtCustReelCheckData.Text = Replace$(txtCustReelCheckData.Text, strScan, "", , 1)

End If

txt37ReelID.Text = Replace$(txt37ReelID.Text, txtCur37ReelID.Text, "", , 1)
txtCur37ReelID.Text = ""
If Replace(txt37ReelID.Text, vbCrLf, "") = "" Then
    Play ("otherIP")
    txtInnerBoxNum.Text = Replace$(txtInnerBoxNum.Text, "<" & txtCurInnerBoxNum.Text & ">", "", , 1)
    txtCurInnerBoxNum.Text = ""
    txt37ReelID.Text = ""

End If

If txtInnerBoxNum.Text = "" Then
    Play ("otherOP")
    txtOuterBoxNum.Text = Replace$(txtOuterBoxNum.Text, "<" & txtCurOuterBoxNum.Text & ">", "", , 1)
    txtCurOuterBoxNum.Text = ""

End If

If txtOuterBoxNum.Text = "" Then
    Play ("completeDN")
    ' 更新检验状态
    If txtDN.Text <> "" Then
        Call UpdateChkStatus(Trim$(txtDN.Text))

    End If

End If

End Sub

Private Sub UpdateChkStatus(strDN As String)
Dim sOra      As String
Dim i         As Integer
Dim rs        As New ADODB.Recordset
Dim strSql    As String
Dim strQboxNO As String, lQty As Long

sOra = "update PACKING_DETAILED set flag = '1' where dn_num = '" & strDN & "'"
AddSql (sOra)
MsgBox "自动核对完成, 可以出货", vbInformation, "友情提示:"
strSql = "insert into HWLBLMATCHHIS(DN,CREATE_DATE,CREATE_BY,STATUS) values('" & strDN & "', sysdate,'" & gUserName & "','PSN PASS') "
AddSql (strSql)

End Sub

Rem: 播放音频提醒
Private Sub Play(sFileName As String)
Dim sPath   As String
Dim sSuffix As String

sPath = txtMediaPath.Text
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix
Sleep (200)

End Sub
