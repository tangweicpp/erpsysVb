VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Frm_LblMatchSysNew 
   BackColor       =   &H00E0E0E0&
   Caption         =   "标签核对系统"
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16080
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
   ScaleHeight     =   10275
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   10815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   19935
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00FF8080&
         Caption         =   "导出核对记录"
         Height          =   360
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1102
         Width           =   1455
      End
      Begin VB.TextBox tRPLbl 
         Height          =   2295
         Left            =   6960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   7200
         Width           =   12255
      End
      Begin VB.TextBox tRPData 
         BackColor       =   &H00C0C0C0&
         Height          =   2295
         Left            =   1560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox tRPVal 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   7080
         Width           =   2415
      End
      Begin VB.TextBox tIPVal 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   6480
         Width           =   2415
      End
      Begin VB.TextBox txtMediaPath 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Text            =   "\\10.160.1.84\public\media_source\"
         Top             =   7800
         Width           =   4215
      End
      Begin VB.TextBox tIPData 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox tOPVal 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   5880
         Width           =   2415
      End
      Begin VB.CommandButton cExit 
         BackColor       =   &H00FF80FF&
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   8640
         Width           =   1695
      End
      Begin VB.TextBox tIPLbl 
         Height          =   2295
         Left            =   6960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   4320
         Width           =   12255
      End
      Begin VB.TextBox tOPLbl 
         Height          =   3135
         Left            =   6960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   600
         Width           =   12255
      End
      Begin VB.TextBox tOPData 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1680
         Width           =   5055
      End
      Begin VB.TextBox tDN 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1140
         Width           =   1455
      End
      Begin VB.TextBox tCode 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label lbl22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   25
         Top             =   3907
         Width           =   555
      End
      Begin VB.Label lblReel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘标签数据:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12120
         TabIndex        =   22
         Top             =   6960
         Width           =   1515
      End
      Begin VB.Label lblReel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前卷盘:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   20
         Top             =   7080
         Width           =   1035
      End
      Begin VB.Label lblIpkg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前内箱:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   19
         Top             =   6480
         Width           =   1035
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   720
         TabIndex        =   17
         Top             =   7680
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
         _cy             =   873
      End
      Begin VB.Label lblPreOP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前外箱:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   13
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Label lblInPkgLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内箱标签数据:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12090
         TabIndex        =   11
         Top             =   4080
         Width           =   1515
      End
      Begin VB.Label lblOutPkgLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱标签数据:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11850
         TabIndex        =   10
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label lblInPkg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内箱:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   7
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label lblOutPkg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   5
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1065
         TabIndex        =   3
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫码框:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   1
         Top             =   622
         Width           =   795
      End
   End
End
Attribute VB_Name = "Frm_LblMatchSysNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim sOpLbl   As String
Dim sIpLbl   As String
Dim sReelLbl As String
Dim sType    As String

Private Sub cExit_Click()
Unload Me

End Sub

Private Sub cmdExport_Click()
Dim strSql As String, strDN As String

If tDN.Text = "" Then
    MsgBox "请输入查找的DN", vbInformation, "提示"
    Exit Sub

End If

strDN = Trim$(tDN.Text)
strSql = "select dn, create_by as 核对人员, create_date as 核对时间, 'OQC PASS' as 核对状态  from HWLBLMATCHHIS where dn = '" & strDN & "' order by create_date desc"
ExporToExcel (strSql)

End Sub

Private Sub Form_Activate()
tCode.SetFocus

End Sub

Private Sub Form_Load()
sOpLbl = ""
sIpLbl = ""
sReelLbl = ""
If gUserName <> "07885" Then
    lblOutPkgLbl.Visible = False
    tOPLbl.Visible = False
    lblInPkgLbl.Visible = False
    tIPLbl.Visible = False
    lblReel2.Visible = False
    tRPLbl.Visible = False
    tRPData.Visible = False
    lbl22.Visible = False
Else
    txtMediaPath.Text = "\\10.160.1.84\public\media_source\"

    '   txtMediaPath.Text = App.Path & "\media_source\"
End If

End Sub

Private Sub tCode_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Or tCode.Text = "" Then
    Exit Sub

End If

Dim sCode As String

sCode = UCase$(Trim$(tCode.Text))
If tDN.Text = "" Then
    Call InitDN(sCode)
ElseIf tOPVal.Text = "" Then
    Call InitOP(sCode)
ElseIf tOPLbl.Text <> "" Then
    Call MatchOP(sCode)
ElseIf tIPVal.Text = "" Then
    Call InitIP(sCode)
ElseIf Replace(tIPLbl.Text, vbCrLf, "") <> "" Then
    Call MatchIP(sCode)
ElseIf tRPVal.Text = "" Then
    Call InitReel(sCode)
ElseIf tRPData.Text <> "" Then
    Call MatchReel(sCode)
Else
    Play ("noScan")

End If

tCode.Text = ""
tCode.SetFocus

End Sub

Private Sub InitDN(sCode As String)
Dim sOutBox As String
Dim rs      As ADODB.Recordset

tDN.Text = Get_OracleStr(" select dn_num from LPSTBL where 'I' || dn_num = '" & sCode & "' ")
If tDN.Text = "" Then
    Play ("wrongDN")
    Exit Sub

End If

' 判断DN类型
sType = GetDNType(tDN.Text)
sOutBox = "select distinct outbox_num from LPSTBL where dn_num = '" & tDN.Text & "' order by outbox_num "
Set rs = Get_OracleRs(sOutBox)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tOPData.Text = tOPData.Text & "<" & rs.Fields(0) & ">"
        rs.MoveNext
    Loop

End If

Play ("rightDN")

End Sub

Private Sub InitOP(sCode As String)
Dim sLbl1  As String
Dim sLbl2  As String
Dim sLbl3  As String
Dim sInbox As String
Dim rs     As ADODB.Recordset

tOPVal.Text = Get_OracleStr(" select distinct outbox_num from LPSTBL where dn_num = '" & tDN.Text & "' and (kid = '" & sCode & "' or carton = '" & sCode & "') ")
If tOPVal.Text = "" Then
    Play ("wrongCode")
    Exit Sub

End If

If InStr(tOPData.Text, "<" & tOPVal.Text & ">") = 0 Then
    Play ("otherOP")
    tOPVal.Text = ""
    Exit Sub

End If

Play ("rightOP")
sInbox = "select distinct inbox_num from LPSTBL where dn_num = '" & tDN.Text & "' and outbox_num = '" & tOPVal.Text & "'  order by inbox_num "
Set rs = Get_OracleRs(sInbox)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tIPData.Text = tIPData.Text & "<" & rs.Fields(0) & ">"
        rs.MoveNext
    Loop

End If

tOPLbl.Text = ""
' Sangsung外箱汇总标签
sLbl1 = "select distinct '<I' || dn_num || '>' || '<K' || nvl(substr(po,1, 10), 'N/A') || '>' || '<P' || nvl(substr(cpn, 1, 11), 'N/A') || '>' || '<Z' || mpn || '>' || '<Q' || sum(qty) || '>' from LPSTBL where dn_num = '" & tDN.Text & "' and outbox_num = '" & tOPVal.Text & "' group by dn_num,po,cpn,mpn "
tOPLbl.Text = UCase(Trim(Get_OracleStr(sLbl1))) & vbCrLf
' Semtech外箱C标签
sLbl3 = "select distinct cartonid from LPSTBL where dn_num = '" & tDN.Text & "' and outbox_num = '" & tOPVal.Text & "'"
Set rs = Get_OracleRs(sLbl3)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tOPLbl.Text = tOPLbl.Text & "<" & rs.Fields(0) & ">" & vbCrLf
        rs.MoveNext
    Loop

End If

sOpLbl = tOPLbl.Text

End Sub

Private Sub InitIP(sCode As String)
Dim sLbl1 As String
Dim sLbl2 As String
Dim sLbl3 As String
Dim sReel As String
Dim rs    As ADODB.Recordset

tIPVal.Text = Get_OracleStr(" select distinct inbox_num from LPSTBL where dn_num = '" & tDN.Text & "' and boxid = '" & sCode & "' and outbox_num = '" & tOPVal.Text & "' ")
If tIPVal.Text = "" Then
    Play ("wrongCode")
    Exit Sub

End If

Play ("rightCode")
If InStr(tIPData.Text, "<" & tIPVal.Text & ">") = 0 Then
    Play ("otherIP")
    tIPVal.Text = ""
    Exit Sub

End If

' 初始化ReelData
sReel = "select distinct trayid from LPSTBL where dn_num = '" & tDN.Text & "' and outbox_num = '" & tOPVal.Text & "' and inbox_num = '" & tIPVal.Text & "' "
Set rs = Get_OracleRs(sReel)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tRPData.Text = tRPData.Text & rs.Fields(0) & vbCrLf
        rs.MoveNext
    Loop

End If

' Semtech内盒B标签
tIPLbl.Text = ""
sLbl1 = "select distinct boxid from LPSTBL where dn_num = '" & tDN.Text & "' and outbox_num = '" & tOPVal.Text & "' and inbox_num = '" & tIPVal.Text & "' "
Set rs = Get_OracleRs(sLbl1)
If Not rs.BOF Then
    rs.MoveFirst

    Do While Not rs.EOF
        tIPLbl.Text = tIPLbl.Text & "<" & rs.Fields(0) & ">" & vbCrLf
        rs.MoveNext
    Loop

End If

' Sangsung内盒汇总标签
If sType = "SSE2" Then
    Dim sType2 As String

    sType2 = Get_OracleStr("select UPPER(labelrequirement) as type from CUSTOMERSHIPPINGUPTBL where delivery = '" & tDN.Text & "'")
    If InStr(sType2, "SHORT") Then
        sLbl2 = "select distinct '<' ||cpn||'DPTK'||substr('000000' || sum(qty), -6, 6)||'>' from LPSTBL where dn_num = '" & tDN.Text & "' and outbox_num = '" & tOPVal.Text & "' and inbox_num = '" & tIPVal.Text & "' group by cpn"
    Else
        sLbl2 = "select distinct '<' ||cpn||'DPTKE2'||substr('000000' || sum(qty), -6, 6)||'>' from LPSTBL where dn_num = '" & tDN.Text & "' and outbox_num = '" & tOPVal.Text & "' and inbox_num = '" & tIPVal.Text & "' group by cpn"

    End If

    tIPLbl.Text = tIPLbl.Text & Get_OracleStr(sLbl2) & vbCrLf

End If

sIpLbl = tIPLbl.Text
tIPLbl.Text = Replace$(tIPLbl.Text, "<" & sCode & ">", "", , 1)
' Sangsung卷盘标签
tRPLbl.Text = ""
If sType = "SSE2" Then
    sType2 = Get_OracleStr("select UPPER(labelrequirement) as type from CUSTOMERSHIPPINGUPTBL where delivery = '" & tDN.Text & "'")
    If InStr(sType2, "SHORT") Then
        sLbl3 = "select distinct cpn||'DPTK'||reelid||substr('000000' || qty, -6, 6) from LPSTBL where dn_num = '" & tDN.Text & "' and outbox_num = '" & tOPVal.Text & "' and inbox_num = '" & tIPVal.Text & "'"
    Else
        sLbl3 = "select distinct cpn||'DPTKE2'||reelid||substr('000000' || qty, -6, 6) from LPSTBL where dn_num = '" & tDN.Text & "' and outbox_num = '" & tOPVal.Text & "' and inbox_num = '" & tIPVal.Text & "'"

    End If

    Set rs = Get_OracleRs(sLbl3)
    If Not rs.BOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            tRPLbl.Text = tRPLbl.Text & "<" & rs.Fields(0) & ">" & vbCrLf
            rs.MoveNext
        Loop

    End If

End If

End Sub

Private Sub InitReel(sCode As String)
Dim sLbl1 As String
Dim rs    As ADODB.Recordset

tRPVal.Text = Get_OracleStr("select distinct trayid from LPSTBL where dn_num = '" & tDN.Text & "' and trayid = '" & sCode & "' and outbox_num = '" & tOPVal.Text & "' and inbox_num = '" & tIPVal.Text & "'")
If tRPVal.Text = "" Then
    Play ("wrongCode")
    Exit Sub

End If

Play ("rightCode")
If InStr(tRPData.Text, tRPVal.Text) = 0 Then
    Play ("该卷盘已核对, 请勿重复核对")
    tRPVal.Text = ""
    Exit Sub

End If

tRPData.Text = Replace$(tRPData.Text, sCode, "", , 1)
If sType <> "SSE2" Then
    tRPVal.Text = ""
    If Replace(tRPData.Text, vbCrLf, "") = "" Then
        Play ("otherIP")
        tIPData.Text = Replace$(tIPData.Text, "<" & tIPVal.Text & ">", "", , 1)
        tIPVal.Text = ""
        tRPData.Text = ""

    End If

    If tIPData.Text = "" Then
        Play ("otherOP")
        tOPData.Text = Replace$(tOPData.Text, "<" & tOPVal.Text & ">", "", , 1)
        tOPVal.Text = ""

    End If

    If tOPData.Text = "" Then
        Play ("completeDN")
        ' 更新检验状态
        If tDN.Text <> "" Then
            Call UpdateChkStatus(Trim$(tDN.Text))

        End If

    End If

End If

End Sub

' 核对外箱标签
Private Sub MatchOP(sCode As String)
If sCode = "" Then
    Play ("wrongCode")
    Exit Sub

End If

sCode = "<" & sCode & ">"
If InStr(sOpLbl, sCode) = 0 Then
    Play ("wrongCode")
    Exit Sub

End If

Play ("rightCode")
If InStr(tOPLbl.Text, sCode) = 0 Then
    Play ("repCode")
    Exit Sub

End If

tOPLbl.Text = Replace$(tOPLbl.Text, sCode, "", , 1)
If Replace(tOPLbl.Text, vbCrLf, "") = "" Then
    Play ("nextIP")
    tOPLbl.Text = ""
    sOpLbl = ""

End If

End Sub

' 核对内盒标签
Private Sub MatchIP(sCode As String)
If sCode = "" Then
    Play ("wrongCode")
    Exit Sub

End If

sCode = "<" & sCode & ">"
If InStr(sIpLbl, sCode) = 0 Then
    Play ("wrongCode")
    Exit Sub

End If

Play ("rightCode")
If InStr(tIPLbl.Text, sCode) = 0 Then
    Play ("repCode")
    Exit Sub

End If

tIPLbl.Text = Replace$(tIPLbl.Text, sCode, "", , 1)
If Replace(tIPLbl.Text, vbCrLf, "") = "" Then
    Play ("nextReel")
    tIPLbl.Text = ""
    sIpLbl = ""

End If

End Sub

Private Sub MatchReel(sCode As String)
Dim sJob As String
Dim sLot As String

If sCode = "" Then
    Play ("wrongCode")
    Exit Sub

End If

If sType = "SSE2" Then
    Dim sType2 As String

    sType2 = Get_OracleStr("select UPPER(labelrequirement) as type from CUSTOMERSHIPPINGUPTBL where delivery = '" & tDN.Text & "'")
    sCode = "<" & sCode & ">"
    If InStr(tRPLbl, sCode) = 0 Then
        Play ("wrongCode")
        Exit Sub

    End If

    sJob = Get_OracleStr("select job_id from lpstbl where dn_num = '" & tDN.Text & "'  and trayid = '" & tRPVal.Text & "' ")
    If InStr(sType2, "SHORT") Then
        If InStr(sCode, "M") > 0 Then
            sLot = Mid(sCode, InStr(sCode, "DPTK") + 4, 9)
        Else
            sLot = Mid$(sCode, InStr(sCode, "DPTK") + 4, 8)

        End If

    Else
        If InStr(sCode, "M") > 0 Then
            sLot = Mid(sCode, InStr(sCode, "E2") + 2, 9)
        Else
            sLot = Mid$(sCode, InStr(sCode, "E2") + 2, 8)

        End If

    End If

    If Left(sLot, 6) <> Left$(sJob, 6) Then
        Play ("wrongJob")
        Exit Sub

    End If

    Play ("otherReel")
    tRPLbl.Text = Replace$(tRPLbl.Text, sCode, "", , 1)

End If

tRPData.Text = Replace$(tRPData.Text, tRPVal.Text, "", , 1)
tRPVal.Text = ""
If Replace(tRPData.Text, vbCrLf, "") = "" Then
    Play ("otherIP")
    tIPData.Text = Replace$(tIPData.Text, "<" & tIPVal.Text & ">", "", , 1)
    tIPVal.Text = ""
    tRPData.Text = ""

End If

If tIPData.Text = "" Then
    Play ("otherOP")
    tOPData.Text = Replace$(tOPData.Text, "<" & tOPVal.Text & ">", "", , 1)
    tOPVal.Text = ""

End If

If tOPData.Text = "" Then
    Play ("completeDN")
    ' 更新检验状态
    If tDN.Text <> "" Then
        Call UpdateChkStatus(Trim$(tDN.Text))

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
strSql = "insert into HWLBLMATCHHIS(DN,CREATE_DATE,CREATE_BY,STATUS) values('" & strDN & "', sysdate,'" & gUserName & "','OQC PASS') "
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
