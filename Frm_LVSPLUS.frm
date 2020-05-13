VERSION 5.00
Begin VB.Form Frm_LVSPLUS 
   BackColor       =   &H00C0C0C0&
   Caption         =   "标签核对系统LVS"
   ClientHeight    =   10395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10395
   ScaleWidth      =   15435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0080C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8880
      Width           =   2175
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H008080FF&
      Caption         =   "重置"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   2640
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8880
      Width           =   2175
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00C0C0C0&
      Caption         =   "内外箱防混料核对"
      Height          =   4935
      Left            =   360
      TabIndex        =   11
      Top             =   3480
      Width           =   14655
      Begin VB.OptionButton Opt7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "内盒录入"
         Height          =   315
         Left            =   3720
         TabIndex        =   36
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Opt6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "卷盘 VS 内盒"
         Height          =   195
         Left            =   3720
         TabIndex        =   35
         Top             =   2040
         Width           =   1935
      End
      Begin VB.OptionButton Opt5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "外箱 VS 外箱"
         Height          =   195
         Left            =   360
         TabIndex        =   34
         Top             =   1800
         Width           =   3255
      End
      Begin VB.OptionButton Opt4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "外箱:  Semtech分标签录入"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtInQty 
         Height          =   285
         Left            =   9360
         TabIndex        =   31
         Top             =   3435
         Width           =   2295
      End
      Begin VB.TextBox txtInLot 
         Height          =   285
         Left            =   6120
         TabIndex        =   29
         Top             =   3435
         Width           =   2295
      End
      Begin VB.TextBox txtPkgSeq 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3960
         Width           =   2295
      End
      Begin VB.TextBox txtLog2 
         ForeColor       =   &H00C00000&
         Height          =   2055
         Left            =   6120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   720
         Width           =   5655
      End
      Begin VB.OptionButton Opt3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "内盒 VS 内盒"
         Height          =   375
         Left            =   3720
         TabIndex        =   22
         Top             =   1680
         Width           =   3375
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "内盒 VS 外箱"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   2040
         Width           =   3015
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00C0C0C0&
         Caption         =   "外箱:  客户分标签录入"
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txtLot 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Timer Timer2 
         Interval        =   1200
         Left            =   3600
         Top             =   480
      End
      Begin VB.TextBox txtScan2 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内盒数量"
         Height          =   195
         Index           =   2
         Left            =   8640
         TabIndex        =   30
         Top             =   3480
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内盒Job"
         Height          =   195
         Index           =   1
         Left            =   5400
         TabIndex        =   28
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblPKG_NO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PKG_NO"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   3960
         Width           =   600
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   195
         Left            =   600
         TabIndex        =   18
         Top             =   3600
         Width           =   270
      End
      Begin VB.Label lblLOT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOT"
         Height          =   195
         Left            =   600
         TabIndex        =   16
         Top             =   3240
         Width           =   285
      End
      Begin VB.Label lblScan2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "条码扫描框"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   900
      End
   End
   Begin VB.Frame Frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DN信息匹配"
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   14655
      Begin VB.CommandButton cmdNoCus 
         BackColor       =   &H00FF00FF&
         Caption         =   "切换(无客户标签)"
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
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtOPQty 
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtLog 
         ForeColor       =   &H00FF0000&
         Height          =   2415
         Left            =   4320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   480
         Width           =   5535
      End
      Begin VB.CommandButton cmdCus 
         BackColor       =   &H00FF8080&
         Caption         =   "切换(含客户标签)"
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
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Timer Timer1 
         Interval        =   800
         Left            =   3600
         Top             =   360
      End
      Begin VB.TextBox txtMPN 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtCPN 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1728
         Width           =   2055
      End
      Begin VB.TextBox txtPO 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1297
         Width           =   2055
      End
      Begin VB.TextBox txtDN 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   866
         Width           =   2055
      End
      Begin VB.TextBox txtScan 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   435
         Width           =   2055
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱总数量"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   26
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label lblMPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MPN"
         Height          =   195
         Left            =   945
         TabIndex        =   10
         Top             =   2160
         Width           =   315
      End
      Begin VB.Label lblCPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPN"
         Height          =   195
         Left            =   960
         TabIndex        =   7
         Top             =   1680
         Width           =   300
      End
      Begin VB.Label lblPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO"
         Height          =   195
         Left            =   1050
         TabIndex        =   5
         Top             =   1320
         Width           =   210
      End
      Begin VB.Label lblINVOICE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE"
         Height          =   195
         Left            =   630
         TabIndex        =   3
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lblScan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "条码扫描框"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   900
      End
   End
End
Attribute VB_Name = "Frm_LVSPLUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub cmdCus_Click()
Opt.Visible = True
Opt4.Visible = False
If txtOPQty.Text = "" Then
        MsgBox "请先扫描外箱总数量", vbInformation
        Exit Sub
End If

txtScan2.SetFocus

txtPkgSeq.Text = Get_OracleStr("select lvs_seq.nextval from dual")

Opt.Value = True
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdNoCus_Click()
Opt.Visible = False
Opt4.Visible = True
 If txtOPQty.Text = "" Then
        MsgBox "请先扫描外箱总数量", vbInformation
        Exit Sub
 End If

txtScan2.SetFocus

txtPkgSeq.Text = Get_OracleStr("select lvs_seq.nextval from dual")

Opt4.Value = True

End Sub

Private Sub cmdReset_Click()
Unload Me
Frm_LVSPLUS.Show

End Sub

Private Sub Form_Activate()
txtScan.SetFocus
'Opt.Value = True
End Sub

Private Sub Form_Load()

Opt4.Visible = False
Opt.Visible = False
End Sub

Private Sub Opt_Click()
txtScan2.SetFocus
End Sub

Private Sub Opt2_Click()
txtScan2.SetFocus
End Sub

Private Sub Opt3_Click()
Timer2.Interval = 2000
txtScan2.SetFocus

'txtPkgSeq.Text = ""
End Sub

Private Sub Opt4_Click()
txtScan2.SetFocus
End Sub

Private Sub Opt5_Click()
txtScan2.SetFocus
End Sub

Private Sub Opt6_Click()
Timer2.Interval = 4000
txtScan2.SetFocus
End Sub

Private Sub Opt7_Click()
txtScan2.SetFocus
End Sub

Private Sub Timer1_Timer()

' 查询固定信息
InitData

' 匹配信息
MatchData

' 清空扫描框
txtScan.Text = ""
End Sub


Private Sub InitData()

Dim sOra As String
Dim sDN As String
Dim sHead As String
Dim Rs As New ADODB.Recordset

If txtScan.Text = "" Then
    Exit Sub
End If

sHead = Left$(txtScan.Text, 1)
sDN = Mid$(txtScan.Text, 2)

If sHead = "I" Then
    sOra = "select distinct delivery, purchasingdocno, customerpartnumber, marketingpn from CUSTOMERSHIPPINGUPTBL where delivery = '" & sDN & "'"
    Set Rs = Get_OracleRs(sOra)
    
    If Rs.RecordCount = 0 Then
        MsgBox "扫描的DN号不存在", vbInformation
        Exit Sub
    End If
    
    txtDN.Text = IIf(IsNull(Rs.fields(0).Value), "", Rs.fields(0).Value)
    txtPO.Text = IIf(IsNull(Rs.fields(1).Value), "", Rs.fields(1).Value)
    txtCPN.Text = IIf(IsNull(Rs.fields(2).Value), "", Rs.fields(2).Value)
    txtMPN.Text = IIf(IsNull(Rs.fields(3).Value), "", Rs.fields(3).Value)
End If

End Sub

Private Sub MatchData()

Dim sScan As String
Dim sHead As String
Dim sHead2 As String
Dim sSel As String
Dim sSel2 As String

sScan = txtScan.Text
If sScan = "" Then
    Exit Sub
End If

sHead = Left$(txtScan.Text, 1)
sHead2 = Left$(txtScan.Text, 2)
sSel = Mid$(txtScan.Text, 2)
sSel2 = Mid$(txtScan.Text, 3)

If sHead = "K" Then
    If sSel = txtPO.Text Then
        txtLog.Text = txtLog.Text & Time() & "----" & "外箱(K) PURCHASE ORDER: " & sSel & "匹配正确" & vbCrLf
    Else
        txtLog.Text = txtLog.Text & Time() & "----" & "外箱(K) PURCHASE ORDER: " & sSel & "错误" & vbCrLf
        MsgBox "外箱(K) PURCHASE ORDER: " & sSel & "错误", vbInformation
        Exit Sub
    End If

End If

If sHead = "P" And InStr(sScan, "-") Then
    If sSel = txtCPN.Text Then
        txtLog.Text = txtLog.Text & Time() & "----" & "外箱(P) CUSTOMER PART NO: " & sSel & "匹配正确" & vbCrLf
    Else
        txtLog.Text = txtLog.Text & Time() & "----" & "外箱(P) CUSTOMER PART NO: " & sSel & "错误" & vbCrLf
        MsgBox "外箱(P) CUSTOMER PART NO: " & sSel & "错误", vbInformation
        Exit Sub
    End If
    
End If

If sHead = "Z" Then
    If sSel = txtMPN.Text Then
        txtLog.Text = txtLog.Text & Time() & "----" & "外箱(Z) MFG PART NO: " & sSel & "匹配正确" & vbCrLf
    Else
        txtLog.Text = txtLog.Text & Time() & "----" & "外箱(Z) MFG PART NO: " & sSel & "错误" & vbCrLf
        MsgBox "外箱(Z) MFG PART NO: " & sSel & "错误", vbInformation
        Exit Sub
    End If
    
End If

If sHead2 = "1P" Then
    If sSel2 = txtMPN.Text Then
        txtLog.Text = txtLog.Text & Time() & "----" & "SemTech分标签(1P) MFG: " & sSel & "匹配正确" & vbCrLf
    Else
        txtLog.Text = txtLog.Text & Time() & "----" & "SemTech分标签(1P) MFG: " & sSel & "错误" & vbCrLf
     '   MsgBox "外箱小标签(1P) MFG: " & sSel & "错误", vbInformation
        Exit Sub
    End If
End If

If InStr(sScan, "DPTK") And InStr(sScan, "-") Then
    If InStr(sScan, txtCPN.Text) Then
        txtLog.Text = txtLog.Text & Time() & "----" & "客户分标签 : " & sScan & "匹配正确" & vbCrLf
    Else
        txtLog.Text = txtLog.Text & Time() & "----" & "客户分标签 : " & sScan & "错误" & vbCrLf
        MsgBox "客户分标签 : " & sScan & "错误", vbInformation
        Exit Sub
    End If
End If

If sHead = "Q" Then
    txtOPQty.Text = sSel
    txtLog.Text = txtLog.Text & Time() & "----" & "外箱总数量 : " & sSel & vbCrLf
End If

End Sub

Private Sub Timer2_Timer()

' 外箱信息录入
If Opt.Value = True Then
    ChkOutPkgCus
    InserOutPkg
    MatchOutPkg
    
ElseIf Opt4.Value = True Then
    '无客户版
    ChkOutPkgNoCus
    InserOutPkg
    MatchOutPkg
    
ElseIf Opt5.Value = True Then
    '外箱比较
    ChkInPkg
    MatchInPkg
ElseIf Opt2.Value = True Then
' 内盒核对
    ChkInPkg
    MatchInPkg
ElseIf Opt7.Value = True Then
    
    InitInPkg
ElseIf Opt6.Value = True Then
' 卷盘核对
    InserTrPkg2
    ChkTrPkg
ElseIf Opt3.Value = True Then
    InserTrPkg
End If

' 清空扫描框
txtScan2.Text = ""

End Sub

Private Sub ChkOutPkgNoCus()

Dim sScan As String
Dim sHead As String
Dim sHead2 As String
Dim sSel As String
Dim sSel2 As String
Dim sOra As String

sScan = txtScan2.Text
If sScan = "" Then
    Exit Sub
End If

sHead = Left$(txtScan2.Text, 1)
sHead2 = Left$(txtScan2.Text, 2)
sSel = Mid$(txtScan2.Text, 2)
sSel2 = Mid$(txtScan2.Text, 3)

If sHead2 = "1T" Then
    If txtDN.Text = "" Then
        MsgBox "请先扫描DN", vbInformation
        Exit Sub
    End If

    sOra = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "' and batchnumber = '" & sSel2 & "' "
    If Get_OracleCnt(sOra) = 0 Then
        txtLog2.Text = txtLog2.Text & Time() & "----" & "DN:" & txtDN.Text & "中不包含该LOT:" & sSel2 & vbCrLf
        MsgBox "DN:" & txtDN.Text & "中不包含该LOT", vbInformation
        Exit Sub
    Else
       ' txtLog2.Text = txtLog2.Text & Time() & "----" & "DN:" & txtDN.Text & "中包含该LOT:" & sSel2 & vbCrLf
        txtLot.Text = sSel2
    End If
    
End If

If sHead = "Q" Then
    txtQty.Text = sSel
End If

End Sub

Private Sub ChkOutPkgCus()

Dim sScan As String
Dim sHead As String
Dim sHead2 As String
Dim sSel As String
Dim sSel2 As String
Dim sOra As String

sScan = txtScan2.Text
If sScan = "" Then
    Exit Sub
End If

sHead = Left$(txtScan2.Text, 1)
sHead2 = Left$(txtScan2.Text, 2)
sSel = Mid$(txtScan2.Text, 2)
sSel2 = Mid$(txtScan2.Text, 3)

If sHead = "P" Then
    If txtDN.Text = "" Then
        MsgBox "请先扫描DN", vbInformation
        Exit Sub
    End If

    sOra = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "' and batchnumber = '" & sSel & "' "
    If Get_OracleCnt(sOra) = 0 Then
        txtLog2.Text = txtLog2.Text & Time() & "----" & "DN:" & txtDN.Text & "中不包含该LOT:" & sSel & vbCrLf
        MsgBox "DN:" & txtDN.Text & "中不包含该LOT", vbInformation
        Exit Sub
    Else
        'txtLog2.Text = txtLog2.Text & Time() & "----" & "DN:" & txtDN.Text & "中包含该LOT:" & sSel & vbCrLf
        txtLot.Text = sSel
    End If
    
End If

If sHead = "Q" Then
    txtQty.Text = sSel
End If

End Sub

Private Sub InserOutPkg()

Dim sOra As String

If txtScan2.Text = "" Then
    Exit Sub
End If

If txtLot.Text = "" Or txtQty.Text = "" Or txtPkgSeq.Text = "" Then
    Exit Sub
End If

sOra = "Insert into LVS_TBL values('" & txtLot.Text & "', '" & txtQty.Text & "', '" & txtPkgSeq.Text & "', 'OP','0', sysdate) "
Exec_Ora (sOra)

'txtLog2.Text = txtLog2.Text & Time() & "----" & "外箱Lot: " & txtLot.Text & "  " & "Qty: " & txtQty.Text & "累加中" & vbCrLf

txtLot.Text = ""
txtQty.Text = ""

End Sub

Private Sub MatchOutPkg()

Dim sScan As String
Dim sHead As String
Dim sHead2 As String
Dim sSel As String
Dim sSel2 As String
Dim sOra As String

If txtScan2.Text = "" Then
    Exit Sub
End If

sOra = "select sum(QTY) from LVS_TBL where SEQ = '" & txtPkgSeq.Text & "' and type = 'OP' "

If Get_OracleNo(sOra) = (txtOPQty.Text) Then
    txtLog2.Text = txtLog2.Text & Time() & "----" & "外箱信息录入完成: " & Get_OracleNo(sOra) & vbCrLf
   ' txtLog2.Text = txtLog2.Text & Time() & "----" & "外箱与内盒LOT核对准备, 请扫描内盒Lot和Qty:" & vbCrLf
    Opt5.Value = True
    
ElseIf Get_OracleNo(sOra) < (txtOPQty.Text) Then
    txtLog2.Text = txtLog2.Text & Time() & "----" & "外箱总数累计: " & Get_OracleNo(sOra) & vbCrLf
Else
    txtLog2.Text = txtLog2.Text & Time() & "----" & "外箱总数错误: " & Get_OracleNo(sOra) & vbCrLf
    MsgBox "外箱总数大于总标签条码数", vbInformation
    Exit Sub
    
End If

End Sub

Private Sub ChkInPkg()

Dim sScan As String
Dim sHead As String
Dim sHead2 As String
Dim sSel As String
Dim sSel2 As String
Dim sOra As String

sScan = txtScan2.Text
If sScan = "" Then
    Exit Sub
End If

sHead = Left$(txtScan2.Text, 1)
sHead2 = Left$(txtScan2.Text, 2)
sSel = Mid$(txtScan2.Text, 2)
sSel2 = Mid$(txtScan2.Text, 3)

If sHead2 = "1T" Then
    If txtDN.Text = "" Then
        MsgBox "请先扫描DN", vbInformation
        Exit Sub
    End If

    sOra = "select * from LVS_TBL where lot = '" & sSel2 & "' "
    If Get_OracleCnt(sOra) = 0 Then
        txtLog2.Text = txtLog2.Text & Time() & "----" & "外箱标签中发现混料, 不包含该LOT:" & sSel2 & vbCrLf
        MsgBox "DN:" & txtDN.Text & "中发现混料,不包含该LOT", vbInformation
        Exit Sub
    Else
        txtLot.Text = sSel2
    End If
    
End If

If sHead = "Q" Then
    txtQty.Text = sSel
End If

End Sub

Private Sub MatchInPkg()

Dim sOra As String

If txtScan2.Text = "" Then
    Exit Sub
End If

If txtLot.Text = "" Or txtQty.Text = "" Or txtPkgSeq.Text = "" Then
    Exit Sub
End If

sOra = "select * from LVS_TBL where lot = '" & txtLot.Text & "' and seq = '" & txtPkgSeq.Text & "' and qty = '" & txtQty.Text & "' and type = 'OP'"
If Get_OracleCnt(sOra) = 0 Then
    MsgBox "Lot:" & txtLot.Text & " Qty:" & txtQty.Text & " 不匹配,发现混料"
    txtLog2.Text = txtLog2.Text & Time() & "----" & "Lot:" & txtLot.Text & " Qty:" & txtQty.Text & " 不匹配,发现混料" & vbCrLf
    
    txtLot.Text = ""
    txtQty.Text = ""
    
    Exit Sub
End If

sOra = "select * from LVS_TBL where lot = '" & txtLot.Text & "' and seq = '" & txtPkgSeq.Text & "' and qty = '" & txtQty.Text & "' and type = 'OP' and status = '0'"
If Get_OracleCnt(sOra) > 0 Then
    txtLog2.Text = txtLog2.Text & Time() & "----" & "Lot:" & txtLot.Text & " Qty:" & txtQty.Text & vbCrLf
    
    ' 更新状态
    sOra = "update LVS_TBL set status = '1' where lot = '" & txtLot.Text & "' and seq = '" & txtPkgSeq.Text & "' and qty = '" & txtQty.Text & "' and status <> '1' and ROWNUM <= 1"
    Exec_Ora (sOra)
End If

sOra = "select * from LVS_TBL where lot = '" & txtLot.Text & "' and seq = '" & txtPkgSeq.Text & "' and qty = '" & txtQty.Text & "' and type = 'OP' and status = '0'"
If Get_OracleCnt(sOra) = 0 Then
        txtLog2.Text = txtLog2.Text & Time() & "----" & "核对完成" & vbCrLf
     '  txtPkgSeq.Text = ""
       ' Opt3.Value = True
End If

txtLot.Text = ""
txtQty.Text = ""

End Sub

Private Sub InitInPkg()

Dim sScan As String
Dim sHead As String
Dim sHead2 As String
Dim sSel As String
Dim sSel2 As String
Dim sOra As String

sScan = txtScan2.Text
If sScan = "" Then
    Exit Sub
End If

sHead = Left$(txtScan2.Text, 1)
sHead2 = Left$(txtScan2.Text, 2)
sSel = Mid$(txtScan2.Text, 2)
sSel2 = Mid$(txtScan2.Text, 3)

If sHead2 = "1T" Then
    txtInLot.Text = sSel2
End If

If sHead = "Q" Then
    txtInQty.Text = sSel
End If

If txtInLot.Text = "" Or txtInQty.Text = "" Then
    Exit Sub
End If

' 准备新的Seq
txtPkgSeq.Text = Get_OracleStr("select lvs_seq.nextval from dual")

' 插入内盒标签Lot数据
sOra = "Insert into LVS_TBL values('" & txtInLot.Text & "', '" & txtInQty.Text & "', '" & txtPkgSeq.Text & "', 'IP','0', sysdate) "
Exec_Ora (sOra)

txtLog2.Text = txtLog2.Text & Time() & "----" & "内盒Lot: " & txtInLot.Text & "  " & "Qty: " & txtInQty.Text & vbCrLf

End Sub

Private Sub InserTrPkg()

Dim sScan As String
Dim sTQty As String ' 卷盘数量
Dim sOra As String

sScan = txtScan2.Text
If sScan = "" Then
    Exit Sub
End If

If InStr(sScan, txtInQty.Text) = 0 Then
    'MsgBox "SAMSUNG标签: " & sScan & "和内盒(Semtech)标签Qty: " & txtInQty.Text & "不一致,发现混料", vbInformation
    txtLog2.Text = txtLog2.Text & Time() & "----" & "客户标签: " & sScan & "和内盒(Semtech)标签Qty: " & txtInQty.Text & "不一致,发现混料" & vbCrLf
    Exit Sub
Else
    txtLog2.Text = txtLog2.Text & Time() & "----" & "核对完成" & vbCrLf
End If

End Sub

Private Sub InserTrPkg2()

Dim sScan As String
Dim sTQty As String ' 卷盘数量
Dim sOra As String

sScan = txtScan2.Text
If sScan = "" Then
    Exit Sub
End If

If InStr(sScan, txtInLot.Text) = 0 Then
    'MsgBox "SAMSUNG标签: 卷盘Job: " & sScan & "和内盒(Semtech)Job: " & txtInLot.Text & "不一致,发现混料", vbInformation
    txtLog2.Text = txtLog2.Text & Time() & "----" & "SAMSUNG标签: 卷盘Job: " & sScan & "和内盒(Semtech)Job: " & txtInLot.Text & "不一致,发现混料" & vbCrLf
    Exit Sub
Else
    sTQty = Mid(sScan, InStr(sScan, txtInLot.Text) + Len(txtInLot.Text) + 3)

    If IsNumeric(sTQty) And CLng(sTQty) > 0 = True Then
        txtLot.Text = txtInLot.Text
        txtQty.Text = sTQty
    
        ' 插入卷盘标签Lot数据
        sOra = "Insert into LVS_TBL values('" & txtLot.Text & "', '" & txtQty.Text & "', '" & txtPkgSeq.Text & "', 'TP','0', sysdate) "
        Exec_Ora (sOra)
        
    End If
End If

End Sub

Private Sub ChkTrPkg()

Dim sOra As String

If txtScan2.Text = "" Then
    Exit Sub
End If

If txtLot.Text = "" Or txtQty.Text = "" Or txtPkgSeq.Text = "" Then
    Exit Sub
End If

sOra = "select sum(QTY) from LVS_TBL where SEQ = '" & txtPkgSeq.Text & "' and type = 'TP' "
If Get_OracleNo(sOra) = (txtInQty.Text) Then
    txtLog2.Text = txtLog2.Text & Time() & "----" & "卷盘总数核对正确: " & Get_OracleNo(sOra) & vbCrLf
ElseIf Get_OracleNo(sOra) < (txtInQty.Text) Then
    txtLog2.Text = txtLog2.Text & Time() & "----" & "卷盘总数累计: " & Get_OracleNo(sOra) & vbCrLf
Else
    txtLog2.Text = txtLog2.Text & Time() & "----" & "卷盘总数错误: " & Get_OracleNo(sOra) & vbCrLf
    MsgBox "卷盘总数大于内盒标签条码数", vbInformation
    Exit Sub
    
End If
End Sub
