VERSION 5.00
Begin VB.Form FrmLVS 
   Caption         =   "标签核对系统 (LVS)"
   ClientHeight    =   11130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17265
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
   ScaleHeight     =   11130
   ScaleWidth      =   17265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   8895
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   15135
      Begin VB.TextBox txtQty 
         Height          =   1485
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txtJob 
         Height          =   1485
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txtDev 
         Height          =   285
         Left            =   5625
         TabIndex        =   10
         Top             =   2715
         Width           =   1695
      End
      Begin VB.TextBox txtCPN 
         Height          =   285
         Left            =   3489
         TabIndex        =   9
         Top             =   2715
         Width           =   1215
      End
      Begin VB.TextBox txtPO 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   2715
         Width           =   1260
      End
      Begin VB.Timer Timer1 
         Interval        =   600
         Left            =   13320
         Top             =   1080
      End
      Begin VB.TextBox txtScan 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtDN 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "卷盘标签"
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "内盒标签"
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   660
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "外箱标签"
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JOB_1"
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY"
         Height          =   195
         Left            =   3360
         TabIndex        =   15
         Top             =   3360
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEVICE"
         Height          =   195
         Left            =   5012
         TabIndex        =   14
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPN"
         Height          =   195
         Left            =   2881
         TabIndex        =   13
         Top             =   2760
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO"
         Height          =   195
         Left            =   840
         TabIndex        =   12
         Top             =   2760
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描框:"
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN#"
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   2160
         Width           =   330
      End
   End
End
Attribute VB_Name = "FrmLVS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bSel As Boolean
Dim tData() As tLVSData

Private Sub Form_Activate()

txtScan.SetFocus
End Sub

Private Sub Form_Load()

' 默认先外箱
Option1.Value = True
bSel = False

End Sub

Private Sub Timer1_Timer()

' 外箱
If Option1.Value = True Then
    ChkOutPkg
ElseIf Option2.Value = True Then
   ' ChkInPkg
Else
   ' ChkTray
End If


' 清空扫描框,等待再次扫描
txtScan.Text = ""

End Sub

' 外箱核对
Private Sub ChkOutPkg()
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim sOra As String

' 先锁定DN
If Left(txtScan.Text, 1) = "I" Then
    txtDN.Text = Replace(txtScan.Text, "I", "")

    ' 判断后台是否有该DN没有则报错
    sOra = "select * from CUSTOMERSHIPPINGUPTBL where delivery = '" & txtDN.Text & "'"
    Set rs = Get_OracleRs(sOra)
    i = rs.RecordCount

    If i = 0 Then
        MsgBox "未查询到该DN号, 请确认是否有误", vbInformation
        txtScan.Text = ""
        Exit Sub
    Else
        bSel = True

        ReDim tData(i) As tLVSData

        If i = 1 Then
            tData(0).sDN = IIf(IsNull(rs.fields("delivery").Value), "", rs.fields("delivery").Value)
            tData(0).sPO = IIf(IsNull(rs.fields("purchasingdocno").Value), "", rs.fields("purchasingdocno").Value)
            tData(0).sCPN = IIf(IsNull(rs.fields("customerpartnumber").Value), "", rs.fields("customerpartnumber").Value)
            tData(0).sDev = IIf(IsNull(rs.fields("marketingpn").Value), "", rs.fields("marketingpn").Value)
            tData(0).sJobNo = IIf(IsNull(rs.fields("batchnumber").Value), "", rs.fields("batchnumber").Value)
            tData(0).sJobQty = IIf(IsNull(rs.fields("quantity").Value), "", rs.fields("quantity").Value)
            
            txtPO.Text = tData(0).sPO
            txtCPN.Text = tData(0).sCPN
            txtDev.Text = tData(0).sDev
            txtJob.Text = tData(0).sJobNo
            txtQty.Text = tData(0).sJobQty
            
        Else
             ' 查询赋值
            For j = 0 To i - 1
                tData(j).sDN = IIf(IsNull(rs.fields("delivery").Value), "", rs.fields("delivery").Value)
                tData(j).sPO = IIf(IsNull(rs.fields("purchasingdocno").Value), "", rs.fields("purchasingdocno").Value)
                tData(j).sCPN = IIf(IsNull(rs.fields("customerpartnumber").Value), "", rs.fields("customerpartnumber").Value)
                tData(j).sDev = IIf(IsNull(rs.fields("marketingpn").Value), "", rs.fields("marketingpn").Value)
                tData(j).sJobNo = IIf(IsNull(rs.fields("batchnumber").Value), "", rs.fields("batchnumber").Value)
                tData(j).sJobQty = IIf(IsNull(rs.fields("quantity").Value), "", rs.fields("quantity").Value)
                
                rs.MoveNext
            Next
            
            
            
            For j = 0 To i - 1
                txtPO.Text = tData(0).sPO
                txtCPN.Text = tData(0).sCPN
                txtDev.Text = tData(0).sDev
                txtJob.Text = txtJob.Text & tData(j).sJobNo & "  " & tData(j).sJobQty & vbCrLf
    
            Next
           
        End If

        

    End If

End If

If bSel = False Then

    txtScan.Text = ""
    Exit Sub
End If

' 判断条码
If Left$(txtScan.Text, 1) = "K" Then
    If txtScan.Text = "K" & txtPO.Text Then
        txtPO.ForeColor = vbBlue
    Else
        txtPO.ForeColor = vbRed
    End If

End If

If Left$(txtScan.Text, 1) = "P" And InStr(txtScan.Text, "-") Then
    If txtScan.Text = "P" & txtCPN.Text Then
        txtCPN.ForeColor = vbBlue
    Else
        txtCPN.ForeColor = vbRed
    End If

End If

If Left$(txtScan.Text, 1) = "Z" Then
    If txtScan.Text = "Z" & txtDev.Text Then
        txtDev.ForeColor = vbBlue
    Else
        txtDev.ForeColor = vbRed
    End If

End If

If Left$(txtScan.Text, 1) = "P" And (InStr(txtScan.Text, "-") = 0) Then
    If txtScan.Text = "P" & txtJob.Text Then
        txtJob.ForeColor = vbBlue
    Else
        txtJob.ForeColor = vbRed
    End If

End If

End Sub
