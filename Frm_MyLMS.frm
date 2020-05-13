VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Frm_MyLMS 
   BackColor       =   &H00C0C0C0&
   Caption         =   "标签核对系统"
   ClientHeight    =   10815
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18105
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
   ScaleHeight     =   10815
   ScaleWidth      =   18105
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtBID 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtKID 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtMediaPath 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Text            =   "C:\Users\tony\Desktop\media_source\"
      Top             =   3360
      Width           =   3855
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FF8080&
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8640
      Width           =   1455
   End
   Begin VB.TextBox txtDN 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtScan 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblBox 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Box:"
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
      Left            =   1575
      TabIndex        =   11
      Top             =   2640
      Width           =   420
   End
   Begin VB.Label lblCARTONID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KID:"
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
      Left            =   1590
      TabIndex        =   8
      Top             =   2160
      Width           =   405
   End
   Begin VB.Label lblMediaFilePath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MediaFile Path:"
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
      Left            =   540
      TabIndex        =   7
      Top             =   3360
      Width           =   1455
   End
   Begin WMPLibCtl.WindowsMediaPlayer media 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
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
      _cx             =   1720
      _cy             =   873
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
      Left            =   1665
      TabIndex        =   2
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "扫描框:"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   1215
      Width           =   795
   End
End
Attribute VB_Name = "Frm_MyLMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KID As String
Dim CartonID As String
Dim CartonIDSub As String
Dim CartonCUSID As String
Dim BoxID As String
Dim BoxCUSID As String
Dim BoxBID As String
Dim ReelID As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Activate()
txtScan.SetFocus
End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)
Dim ts As tScan

ts.sVal = UCase$(Trim$(txtScan.Text))
ts.sKey = Left$(ts.sVal, 1)
ts.sSel = Mid$(ts.sVal, 2)

If KeyAscii <> vbKeyReturn Then
    Exit Sub
End If

If ts.sVal = "" Then
    GoTo clear
End If

If ts.sSel = "" Then
    GoTo clear
End If

If txtDN.Text = "" Then
    Call DN_Now(ts)

ElseIf txtKID.Text = "" Then
    Call KID_Now(ts)

ElseIf CartonID <> "" Then
    Call Carton_Now(ts)

ElseIf CartonIDSub <> "" Then
    Call CartonSub_Now(ts)

ElseIf CartonCUSID <> "" Then
    Call CartonCUS_Now(ts)

ElseIf txtBID.Text = "" Then
    Call Box_Now(ts)

ElseIf BoxBID <> "" Then
    Call BoxBID_Now(ts)

ElseIf BoxCUSID <> "" Then
    Call BoxCUS_Now(ts)

ElseIf ReelID <> "" Then
    Call ReelID_Now(ts)
End If

clear:
txtScan.Text = ""
End Sub

Private Sub DN_Now(ts As tScan)
Dim sOra As String
Dim sReal As String

sOra = "select to_char(wm_concat(distinct dn_num))  from PACKING_DETAILED"
sReal = Get_OracleStr(sOra)
If InStr(sReal, ts.sSel) = 0 Then
    Play ("DN_ERR")
    Exit Sub
End If

txtDN.Text = ts.sSel

Play ("PASS")
Call InitKID
End Sub

Private Sub KID_Now(ts As tScan)
Dim sOra As String
Dim sReal As String

sOra = "select to_char(wm_concat(distinct kid))  from PACKING_DETAILED where dn_num = '" & txtDN.Text & "'"
sReal = Get_OracleStr(sOra)
If InStr(sReal, ts.sSel) = 0 Then
    Play ("KID_ERR")
    Exit Sub
End If
If InStr(KID, ts.sSel) = 0 Then
    Play ("KID_MATCHED")
    Exit Sub
End If

txtKID.Text = ts.sSel

Play ("PASS")
Call InitCartonID
End Sub

Private Sub Box_Now(ts As tScan)
Dim sOra As String
Dim sReal As String
Dim sBoxID As String

sBoxID = Get_OracleStr("select distinct inbox_num from PACKING_DETAILED where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "' and boxid = '" & ts.sVal & "'")
If sBoxID = "" Then
    Play ("BOX_ERR")
    Exit Sub
End If

sOra = "select replace(to_char(wm_concat(distinct inbox_num)), ',','')  from PACKING_DETAILED where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "'"
sReal = Get_OracleStr(sOra)
If InStr(sReal, sBoxID) = 0 Then
    Play ("BOX_ERR")
    Exit Sub
End If
If InStr(BoxID, sBoxID) = 0 Then
    Play ("BOX_MATCHED")
    Exit Sub
End If

txtBID.Text = ts.sVal

Play ("PASS")
Call InitBoxBID
End Sub

Private Sub Carton_Now(ts As tScan)
Dim sOra As String
Dim sReal As String

sOra = "select dn_num||po||cpn||mpn||sum(qty) from LPSTBL where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "' group by dn_num,po,cpn,mpn"
sReal = Get_OracleStr(sOra)
If InStr(sReal, ts.sSel) = 0 Then
    Play ("CARTON_ERR")
    Exit Sub
End If
If InStr(CartonID, ts.sSel) = 0 Then
    Play ("CARTON_MATCHED")
    Exit Sub
End If

Play ("PASS")

CartonID = Replace$(CartonID, ts.sSel, "", , 1)
If CartonID = "" Then
    Play ("CARTON_FIN")
    Call InitCartonIDSub
End If

End Sub

Private Sub CartonSub_Now(ts As tScan)
Dim sOra As String
Dim sReal As String
Dim rs As ADODB.Recordset

sReal = ""

sOra = "select dn_num||po||cpn||mpn||job_id||sum(qty) as label from LPSTBL where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "' group by dn_num,po,cpn,mpn,job_id"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        sReal = sReal & rs.Fields(0)
        rs.MoveNext
    Loop
End If

If InStr(sReal, ts.sSel) = 0 Then
    Play ("CARTONSUB_ERR")
    Exit Sub
End If
If InStr(CartonIDSub, ts.sSel) = 0 Then
    Play ("CARTONSUB_MATCHED")
    Exit Sub
End If

Play ("PASS")

CartonIDSub = Replace$(CartonIDSub, ts.sSel, "", , 1)
If CartonIDSub = "" Then
    Play ("CARTONSUB_FIN")
    Call InitCusID
End If
End Sub

Private Sub CartonCUS_Now(ts As tScan)
Dim sOra As String
Dim sReal As String
Dim rs As ADODB.Recordset

sReal = ""

sOra = "select to_char(wm_concat(distinct cartonid))  from PACKING_DETAILED where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "'"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        sReal = sReal & rs.Fields(0)
        rs.MoveNext
    Loop
End If

If InStr(sReal, ts.sVal) = 0 Then
    Play ("CARTONCUS_ERR")
    Exit Sub
End If
If InStr(CartonCUSID, ts.sVal) = 0 Then
    Play ("CARTONCUS_MATCHED")
    Exit Sub
End If
Play ("PASS")

CartonCUSID = Replace$(CartonCUSID, ts.sVal, "", , 1)
If CartonCUSID = "" Then
    Play ("CARTONCUS_FIN")
    Call InitBoxID
End If
End Sub

Private Sub BoxBID_Now(ts As tScan)
Dim sOra As String
Dim sReal As String
Dim rs As ADODB.Recordset

sReal = ""

sOra = "select to_char(wm_concat(distinct boxid))  from PACKING_DETAILED where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "' and inbox_num = '" & txtBID.Text & "'"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        sReal = sReal & rs.Fields(0)
        rs.MoveNext
    Loop
End If

If InStr(sReal, ts.sVal) = 0 Then
    Play ("BOXBID_ERR")
    Exit Sub
End If
If InStr(BoxBID, ts.sVal) = 0 Then
    Play ("BOXBID_MATCHED")
    Exit Sub
End If
Play ("PASS")

BoxBID = Replace$(BoxBID, ts.sVal, "", , 1)
If BoxBID = "" Then
    Play ("BOXBID_FIN")
    Call InitBoxCusID
End If
End Sub

Private Sub BoxCUS_Now(ts As tScan)
Dim sOra As String
Dim sReal As String
Dim rs As ADODB.Recordset

sReal = ""

sOra = sOra = "select cpn||'DPTKE2'||sum(qty) from LPSTBL where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "' and inbox_num = '" & txtBID.Text & "' group by cpn"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        sReal = sReal & rs.Fields(0)
        rs.MoveNext
    Loop
End If

If InStr(sReal, ts.sVal) = 0 Then
    Play ("BOXCUS_ERR")
    Exit Sub
End If
If InStr(BoxCUSID, ts.sVal) = 0 Then
    Play ("BOXCUS_REP")
    Exit Sub
End If
Play ("PASS")

BoxCUSID = Replace$(BoxCUSID, ts.sVal, "", , 1)
If BoxCUSID = "" Then
    Play ("BOX_FIN")
    Call InitReelID
End If
End Sub

Private Sub ReelID_Now(ts As tScan)
Dim sOra As String
Dim sReal As String
Dim rs As ADODB.Recordset

sReal = ""

sOra = "select trayid || cpn||'DPTKE2'||reelid||'0'||qty  from LPSTBL where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "' and inbox_num = '" & txtBID.Text & "'"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        sReal = sReal & rs.Fields(0)
        rs.MoveNext
    Loop
End If

If InStr(sReal, ts.sVal) = 0 Then
    Play ("REEL_ERR")
    Exit Sub
End If
If InStr(ReelID, ts.sVal) = 0 Then
    Play ("REEL_REP")
    Exit Sub
End If
Play ("PASS")

ReelID = Replace$(ReelID, ts.sVal, "", , 1)
If ReelID = "" Then
    Play ("REEL_FIN")
    BoxID = Replace(BoxID, txtBID.Text, "", , 1)
    txtBID.Text = ""
End If

If BoxID = "" Then
    Play ("CARTON_FULL")
    KID = Replace$(KID, txtKID.Text, "", , 1)
    txtKID.Text = ""
End If

If KID = "" Then
    Play ("DN_FULL")
End If

End Sub

Private Sub InitKID()
Dim sOra As String

sOra = "select replace(to_char(wm_concat(distinct kid)), ',','')  from PACKING_DETAILED where dn_num = '" & txtDN.Text & "'"
KID = Get_OracleStr(sOra)
End Sub

Private Sub InitCartonID()
Dim sOra As String

sOra = "select dn_num||po||cpn||mpn||sum(qty) from LPSTBL where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "' group by dn_num,po,cpn,mpn"
CartonID = Get_OracleStr(sOra)
End Sub

Private Sub InitCartonIDSub()
Dim sOra As String
Dim rs As ADODB.Recordset

sOra = "select dn_num||po||cpn||mpn||job_id||sum(qty) as label from LPSTBL where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "' group by dn_num,po,cpn,mpn,job_id"
Set rs = Get_OracleRs(sOra)

If Not rs.BOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        CartonIDSub = CartonIDSub & rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Sub

Private Sub InitCusID()
Dim sOra As String
Dim rs As ADODB.Recordset

sOra = "select replace(to_char(wm_concat(distinct cartonid)), ',','')  from PACKING_DETAILED where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "'"
Set rs = Get_OracleRs(sOra)

If Not rs.BOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        CartonCUSID = CartonCUSID & rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Sub

Private Sub InitBoxID()
Dim sOra As String
Dim rs As ADODB.Recordset

sOra = "select replace(to_char(wm_concat(distinct inbox_num)), ',','')  from PACKING_DETAILED where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "'"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        BoxID = BoxID & rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Sub

Private Sub InitBoxBID()
Dim sOra As String
Dim rs As ADODB.Recordset

sOra = "select replace(to_char(wm_concat(distinct boxid)), ',','')  from PACKING_DETAILED where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "' and inbox_num = '" & txtBID.Text & "'"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        BoxBID = BoxBID & rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Sub

Private Sub InitBoxCusID()
Dim sOra As String
Dim rs As ADODB.Recordset

sOra = "select cpn||'DPTKE2'||sum(qty) from LPSTBL where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "' and inbox_num = '" & txtBID.Text & "' group by cpn"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        BoxCUSID = BoxCUSID & rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Sub

Private Sub InitReelID()
Dim sOra As String
Dim rs As ADODB.Recordset

sOra = "select trayid || cpn||'DPTKE2'||reelid||'0'||qty  from LPSTBL where dn_num = '" & txtDN.Text & "' and kid = '" & txtKID.Text & "' and inbox_num = '" & txtBID.Text & "'"
Set rs = Get_OracleRs(sOra)
If Not rs.BOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        ReelID = ReelID & rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Sub

Rem: 播放音频提醒
Private Sub Play(sFileName As String)
Dim sPath As String
Dim sSuffix As String

sPath = txtMediaPath.Text
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix
End Sub
