VERSION 5.00
Begin VB.Form Frm_LVS 
   Caption         =   "标签核对系统LVS"
   ClientHeight    =   10755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16410
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
   ScaleHeight     =   10755
   ScaleWidth      =   16410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.TextBox TextInnerLot 
      Height          =   1695
      Left            =   7560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   8895
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   15495
      Begin VB.TextBox txtCusIner 
         Height          =   1695
         Left            =   9720
         TabIndex        =   26
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtTotalQty 
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtLOT2 
         Height          =   1695
         Left            =   4320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtLOT 
         Height          =   1695
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtOutPkgQty 
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtMPN 
         Height          =   285
         Left            =   6120
         TabIndex        =   13
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtCPN 
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtPO 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtDN 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton OptTrayPkg 
         Caption         =   "卷盘"
         Height          =   255
         Left            =   7560
         TabIndex        =   5
         Top             =   690
         Width           =   735
      End
      Begin VB.OptionButton OptInnerPkg 
         Caption         =   "内盒"
         Height          =   375
         Left            =   6180
         TabIndex        =   4
         Top             =   630
         Width           =   735
      End
      Begin VB.OptionButton OptOutPkg 
         Caption         =   "外箱"
         Height          =   195
         Left            =   4560
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   10680
         Top             =   1200
      End
      Begin VB.TextBox txtScan 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   675
         Width           =   1695
      End
      Begin VB.Label lblCus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户内盒标签"
         Height          =   195
         Left            =   8520
         TabIndex        =   27
         Top             =   2280
         Width           =   1080
      End
      Begin VB.Label lblLabel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3960
         TabIndex        =   25
         Top             =   4200
         Width           =   45
      End
      Begin VB.Label lbl22223 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户外箱大标签"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label lblSemTech 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SemTech内盒标签"
         Height          =   195
         Left            =   5640
         TabIndex        =   23
         Top             =   2400
         Width           =   1365
      End
      Begin VB.Label lblLabel1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1560
         TabIndex        =   21
         Top             =   4680
         Width           =   45
      End
      Begin VB.Label lbl22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱小标签"
         Height          =   195
         Left            =   3240
         TabIndex        =   19
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label lblLOTNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOT NO"
         Height          =   195
         Left            =   1560
         TabIndex        =   17
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label lblOutPkgQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OutPkgQty"
         Height          =   195
         Left            =   7680
         TabIndex        =   14
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label lblMPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MPN"
         Height          =   195
         Left            =   5640
         TabIndex        =   12
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label lblCPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPN"
         Height          =   195
         Left            =   3240
         TabIndex        =   10
         Top             =   1680
         Width           =   300
      End
      Begin VB.Label lblPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO"
         Height          =   195
         Left            =   960
         TabIndex        =   8
         Top             =   1680
         Width           =   210
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN"
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描框:"
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Top             =   720
         Width           =   600
      End
   End
End
Attribute VB_Name = "Frm_LVS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aOutLot(50) As tLVS
Dim iPos As Integer


Private Sub Form_Activate()
txtScan.SetFocus
OptOutPkg.Value = True
End Sub

Private Sub Form_Load()
sDN = ""
iPos = 0
End Sub

Private Sub Timer1_Timer()

If OptOutPkg.Value = True Then
    ChkOutPkg
ElseIf OptInnerPkg.Value = True Then
    ChkInnerPkg
Else
    ChkTrayPkg
End If

txtScan.Text = ""

End Sub

Private Sub ChkOutPkg()

Dim sScan As String
Dim sScanHeadChar As String
Dim sScanHeadChar2 As String
Dim sOra As String
Dim sSel As String
Dim sSel2 As String

sScan = txtScan.Text
If sScan = "" Then
    Exit Sub
End If

sScanHeadChar = Left$(sScan, 1)
sScanHeadChar2 = Left$(sScan, 2)
sSel = Mid$(sScan, 2)
sSel2 = Mid$(sScan, 3)

If sScanHeadChar = "I" Then
    If txtDN.Text <> "" Then
        txtDN.Text = ""
        txtPO.Text = ""
        txtCPN.Text = ""
        txtMPN.Text = ""
    End If


    sOra = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & sSel & "'"
    If Get_OracleCnt(sOra) = 0 Then
        MsgBox "DN错误, 请确认", vbInformation
        Exit Sub
    End If
    
    txtDN.Text = sSel
Else
    If sScanHeadChar <> "I" And txtDN.Text = "" Then
        MsgBox "请先扫描外箱的DN", vbInformation
        Exit Sub
    End If
End If

If sScanHeadChar = "K" Then
    sOra = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "' and purchasingdocno = '" & sSel & "'"
    
    If Get_OracleCnt(sOra) = 0 Then
        MsgBox "PO错误, 请确认", vbInformation
        Exit Sub
    End If
    
    txtPO.Text = sSel
End If

If sScanHeadChar = "P" And InStr(sScan, "-") <> 0 Then
    sOra = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "' and customerpartnumber = '" & sSel & "'"
    
    If Get_OracleCnt(sOra) = 0 Then
        MsgBox "CPN错误, 请确认", vbInformation
        Exit Sub
    End If
    
    txtCPN.Text = sSel
End If

If sScanHeadChar = "Z" Then
    sOra = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "' and Marketingpn = '" & sSel & "'"
    
    If Get_OracleCnt(sOra) = 0 Then
        MsgBox "MPN错误, 请确认", vbInformation
        Exit Sub
    End If
    
    txtMPN.Text = sSel
End If

If sScanHeadChar = "Q" Then
    txtOutPkgQty.Text = sSel
End If

If sScanHeadChar = "P" And InStr(sScan, "-") = 0 Then
    
    If InStr(txtLOT.Text, sSel) Then
        MsgBox "请不要重复扫描同一个LOT", vbInformation
        Exit Sub
    End If
    
    aOutLot(iPos).sLot = sSel
    aOutLot(iPos).sQty = Get_OracleNo("select quantity from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "' and batchnumber = '" & sSel & "'")
    
    txtLOT.Text = txtLOT.Text & sSel & vbCrLf
    iPos = iPos + 1
End If


' 小标签
If sScanHeadChar2 = "1T" Then
    
    If InStr(txtLOT2.Text, sSel2) Then
        MsgBox "请不要重复扫描同一个LOT", vbInformation
        Exit Sub
    End If

    If InStr(txtLOT.Text, sSel2) = 0 Then
        MsgBox "外箱小标签的LOT有误", vbInformation
        Exit Sub
    End If
    
    txtLOT2.Text = txtLOT2.Text & sSel2 & vbCrLf
End If

If sScanHeadChar2 = "1P" Then
    sOra = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "' and Marketingpn = '" & sSel2 & "'"
    
    If Get_OracleCnt(sOra) = 0 Then
        MsgBox "外箱小标签MPN错误, 请确认", vbInformation
        Exit Sub
    End If
        
    txtMPN.ForeColor = vbBlue
End If


' 计算累加和
Dim sumQty As Long
sumQty = 0

For i = 0 To UBound(aOutLot) - 1
    If aOutLot(i).sLot <> "" Then
        
        sumQty = sumQty + aOutLot(i).sQty
    End If
Next

txtTotalQty.Text = sumQty

If txtOutPkgQty.Text <> "" Then
    If txtTotalQty.Text = txtOutPkgQty.Text Then
        lblLabel1.Caption = "外箱数量核对完成"
    End If
    
End If

' 小标签累加完成
If txtLOT2.Text <> "" Then
    If Len(txtLOT2.Text) - Len(Replace(txtLOT2.Text, vbCrLf, "")) = Len(txtLOT.Text) - Len(Replace(txtLOT.Text, vbCrLf, "")) Then
        lblLabel2.Caption = "外箱小标签核对完成"
    End If
End If


End Sub

Private Sub ChkInnerPkg()

Dim sScan As String
Dim sScanHeadChar As String
Dim sScanHeadChar2 As String
Dim sOra As String
Dim sSel As String
Dim sSel2 As String

sScan = txtScan.Text
If sScan = "" Then
    Exit Sub
End If

sScanHeadChar = Left$(sScan, 1)
sScanHeadChar2 = Left$(sScan, 2)
sSel = Mid$(sScan, 2)
sSel2 = Mid$(sScan, 3)

If sScanHeadChar2 = "1T" Then
    
    If InStr(TextInnerLot.Text, sSel2) Then
        MsgBox "请不要重复扫描同一个LOT", vbInformation
        Exit Sub
    End If

    If InStr(txtLOT.Text, sSel2) = 0 Then
        MsgBox "内盒Semtech标签的LOT有误", vbInformation
        Exit Sub
    End If
    
    TextInnerLot.Text = TextInnerLot.Text & sSel2 & vbCrLf
End If

If sScanHeadChar2 = "1P" Then
    sOra = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "' and Marketingpn = '" & sSel2 & "'"
    
    If Get_OracleCnt(sOra) = 0 Then
        MsgBox "内盒Semtech标签MPN错误, 请确认", vbInformation
        Exit Sub
    End If
        
    txtMPN.ForeColor = vbGreen
End If


' 小标签累加完成
If TextInnerLot.Text <> "" Then
    If Len(TextInnerLot.Text) - Len(Replace(TextInnerLot.Text, vbCrLf, "")) = Len(txtLOT.Text) - Len(Replace(txtLOT.Text, vbCrLf, "")) Then
        lblLabel2.Caption = "Semtech内盒标签核对完成"
    End If
End If







End Sub

Private Sub ChkTrayPkg()



End Sub
