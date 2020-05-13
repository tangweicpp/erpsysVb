VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmAT71Match 
   Caption         =   "AT71内外箱标签比对"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13290
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
   ScaleHeight     =   9570
   ScaleWidth      =   13290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton btnMatch 
      Caption         =   "比对"
      Height          =   360
      Left            =   5640
      TabIndex        =   14
      Top             =   5640
      Width           =   735
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   4335
      Index           =   0
      Left            =   765
      TabIndex        =   12
      Top             =   3240
      Width           =   4215
      _Version        =   524288
      _ExtentX        =   7435
      _ExtentY        =   7646
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
      MaxCols         =   2
      MaxRows         =   0
      SpreadDesigner  =   "FrmAT71Match.frx":0000
      AppearanceStyle =   0
   End
   Begin VB.CommandButton btnSwitch 
      Caption         =   ">>"
      Height          =   360
      Left            =   5640
      TabIndex        =   9
      Top             =   4320
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "内箱"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7245
      TabIndex        =   7
      Top             =   2745
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "外箱"
      Enabled         =   0   'False
      Height          =   495
      Left            =   765
      TabIndex        =   6
      Top             =   2745
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      BackColor       =   &H00FFC0FF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1365
      TabIndex        =   3
      Top             =   1665
      Width           =   2055
   End
   Begin VB.TextBox txtLotID 
      BackColor       =   &H00FFC0FF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1365
      TabIndex        =   2
      Top             =   1185
      Width           =   2055
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "退出"
      Height          =   375
      Left            =   2325
      TabIndex        =   1
      Top             =   585
      Width           =   1095
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "开始"
      Height          =   375
      Left            =   765
      TabIndex        =   0
      Top             =   585
      Width           =   1095
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   4335
      Index           =   1
      Left            =   7245
      TabIndex        =   13
      Top             =   3240
      Width           =   4215
      _Version        =   524288
      _ExtentX        =   7435
      _ExtentY        =   7646
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
      MaxCols         =   2
      MaxRows         =   0
      SpreadDesigner  =   "FrmAT71Match.frx":044C
      AppearanceStyle =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "或扫描1111条码"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5640
      TabIndex        =   15
      Top             =   6000
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "或扫描0000条码"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5640
      TabIndex        =   11
      Top             =   5040
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "外箱切换内箱"
      Height          =   195
      Left            =   5640
      TabIndex        =   10
      Top             =   4800
      Width           =   1080
   End
   Begin WMPLibCtl.WindowsMediaPlayer media 
      Height          =   495
      Left            =   10320
      TabIndex        =   8
      Top             =   1080
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
      _cy             =   873
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数量："
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   765
      TabIndex        =   5
      Top             =   1725
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "批号："
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   765
      TabIndex        =   4
      Top             =   1245
      Width           =   735
   End
End
Attribute VB_Name = "FrmAT71Match"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private gPkgNO As String

Private Sub PlaySound(sFileName As String)
Dim sPath   As String
Dim sSuffix As String
sPath = "\\10.160.1.84\public\media_source\"
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix

End Sub

Private Sub btnExit_Click()
Unload Me

End Sub

Private Sub btnMatch_Click()
Call MatchData

End Sub

Private Sub MatchData()
Dim rs          As New ADODB.Recordset
Dim strSql      As String
Dim lOutSum     As Long
Dim lInSum      As Long
Dim strCurLotID As String
lOutSum = Get_OracleNo(" select sum(Qty) from tbl_AT71_MATCH where pkgno = '" & gPkgNO & "'|| '_OUT' ")
lInSum = Get_OracleNo(" select sum(Qty) from tbl_AT71_MATCH where pkgno = '" & gPkgNO & "'|| '_IN' ")

If lOutSum <> lInSum Then
    MsgBox "外箱总数量：" & lOutSum & "不等于内箱总数量：" & lInSum & "，标签错误", vbCritical, "警告"
    Exit Sub

End If

strSql = "select distinct lotid from tbl_AT71_MATCH where pkgno = '" & gPkgNO & "'|| '_OUT' "
Set rs = Get_OracleRs(strSql)

Do While Not rs.EOF
    strCurLotID = rs!LOTID
    lOutSum = Get_OracleNo(" select sum(Qty) from tbl_AT71_MATCH where pkgno = '" & gPkgNO & "'|| '_OUT' and lotid = '" & strCurLotID & "'   ")
    lInSum = Get_OracleNo(" select sum(Qty) from tbl_AT71_MATCH where pkgno = '" & gPkgNO & "'|| '_IN' and lotid = '" & strCurLotID & "'   ")

    If lOutSum <> lInSum Then
        MsgBox "Lot：" & strCurLotID & " 外箱数量：" & lOutSum & "不等于内箱数量：" & lInSum & "，标签错误", vbCritical, "警告"
        Exit Sub

    End If

    rs.MoveNext
Loop
Call PlaySound("内外箱标签批号及数量一致,请继续核对其他外箱")
fps(0).MaxRows = 0
fps(1).MaxRows = 0

gPkgNO = Get_OracleStr("select SEQ_AT71_MATCH.Nextval from dual")
Option1.Value = True
txtLotID.SetFocus
txtLotID.Text = ""
End Sub

Private Sub btnStart_Click()
Option1.Value = True
fps(0).MaxRows = 0
fps(1).MaxRows = 0
gPkgNO = Get_OracleStr("select SEQ_AT71_MATCH.Nextval from dual")

If Option1.Value = True Then
    Call PlaySound("请扫描外箱批号")
    txtLotID.Enabled = True
    txtLotID.SetFocus
    txtLotID.Text = ""

End If

End Sub

Private Sub btnSwitch_Click()
Call SwitchBox

End Sub

Private Sub SwitchBox()
Option2.Value = True
Call PlaySound("请扫描内箱批号")
txtLotID.Enabled = True
txtLotID.SetFocus
txtLotID.Text = ""

End Sub

Private Sub txtLotID_KeyPress(KeyAscii As Integer)
Dim strLotID As String

If KeyAscii <> vbKeyReturn Or Len(Trim(txtLotID.Text)) = 0 Then Exit Sub
If txtLotID.Text = "0000" Then
    Call SwitchBox
    Exit Sub

End If

If txtLotID.Text = "1111" Then
    Call MatchData
    Exit Sub

End If

strLotID = Trim$(txtLotID.Text)

If Get_OracleCnt("select * from mappingdatatest where lotid = '" & strLotID & "'") = 0 Then
    Call PlaySound("请扫描正确的批号")
    txtLotID.Text = ""
    Exit Sub

End If

If Option2.Value = True Then
    If Get_OracleCnt("select * from tbl_AT71_MATCH where PKGNO = '" & gPkgNO & "'|| '_OUT' and LOTID = '" & Trim$(txtLotID.Text) & "' ") = 0 Then
        MsgBox "外箱标签不包含该内箱批号，标签错误", vbCritical, "警告"
        txtLotID.Text = ""
        Exit Sub

    End If

End If

Call PlaySound("请扫描数量")
txtQty.Enabled = True
txtQty.SetFocus
txtQty.Text = ""

End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
Dim strQty As String

If KeyAscii <> vbKeyReturn Or Len(Trim(txtQty.Text)) = 0 Then Exit Sub
strQty = Trim$(txtQty.Text)

If IsNumeric(strQty) = False Then
    Call PlaySound("请扫描正确的数量")
    txtQty.Text = ""
    Exit Sub

End If

If Option1.Value = True Then
    AddSql ("insert into TBL_AT71_MATCH(LOTID,QTY,PKGNO,CREATE_DATE) values('" & Trim$(txtLotID.Text) & "'," & CLng(txtQty.Text) & ",'" & gPkgNO & "'|| '_OUT',sysdate) ")
Else
    AddSql ("insert into TBL_AT71_MATCH(LOTID,QTY,PKGNO,CREATE_DATE) values('" & Trim$(txtLotID.Text) & "'," & CLng(txtQty.Text) & ",'" & gPkgNO & "' || '_IN',sysdate) ")

End If

Call ShowData
Call PlaySound("数量已扫描请继续")
txtLotID.Enabled = True
txtLotID.SetFocus
txtLotID.Text = ""
txtQty.Text = ""

End Sub

Private Sub ShowData()
Dim rs     As New ADODB.Recordset
Dim strSql As String
strSql = " select LOTID 批号,sum(Qty) 数量 from tbl_AT71_MATCH where PKGNO = '" & gPkgNO & "'|| '_OUT' group by LOTID"
Set rs = Get_OracleRs(strSql)

With fps(0)
    .MaxRows = 0
    Set .DataSource = rs

End With

rs.Close
strSql = " select LOTID 批号,sum(Qty) 数量 from tbl_AT71_MATCH where PKGNO = '" & gPkgNO & "'|| '_IN' group by LOTID"
Set rs = Get_OracleRs(strSql)

With fps(1)
    .MaxRows = 0
    Set .DataSource = rs

End With

End Sub
