VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_ReelInPos_37 
   Caption         =   "37仓位录入"
   ClientHeight    =   12630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13755
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
   ScaleHeight     =   12630
   ScaleWidth      =   13755
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ReelInPos_37.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ReelInPos_37.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ReelInPos_37.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ReelInPos_37.frx":24F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   1535
      ButtonWidth     =   1455
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "开始扫码"
            Key             =   "START"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "绑定仓位"
            Key             =   "BAND"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "NextJob"
            Key             =   "NEXT"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "CLOSE"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   15375
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   13815
      Begin VB.CheckBox chk 
         Caption         =   "全选/反选"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   5280
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtScan 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   1185
         Visible         =   0   'False
         Width           =   4815
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   12375
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   2160
         Width           =   4815
         _Version        =   524288
         _ExtentX        =   8493
         _ExtentY        =   21828
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
         MaxCols         =   3
         MaxRows         =   0
         SpreadDesigner  =   "Frm_ReelInPos_37.frx":2848
         Appearance      =   1
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin MSForms.TextBox txtPos 
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1455
         VariousPropertyBits=   746604563
         ForeColor       =   255
         BorderStyle     =   1
         Size            =   "2566;661"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblPos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "仓位:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3120
         TabIndex        =   8
         Top             =   1680
         Width           =   630
      End
      Begin MSForms.TextBox txtJobID 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1455
         VariousPropertyBits=   746604567
         ForeColor       =   255
         BorderStyle     =   1
         Size            =   "2566;661"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblJOBID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JOBID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   840
         TabIndex        =   6
         Top             =   1680
         Width           =   630
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   8520
         TabIndex        =   5
         Top             =   1200
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
      Begin VB.Label lblReelInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘信息:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   480
         TabIndex        =   4
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Label lblScan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描JOB:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   600
         TabIndex        =   1
         Top             =   1200
         Width           =   900
      End
   End
End
Attribute VB_Name = "Frm_ReelInPos_37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Play(sFileName As String)

    Dim sPath   As String

    Dim sSuffix As String

    sPath = "\\10.160.1.84\public\media_source\"
    sSuffix = ".wav"
    media.url = sPath & sFileName & sSuffix
    
End Sub

Private Sub chk_Click()

    Dim i As Integer

    If chk.Value = 1 Then

        For i = 1 To fpS(0).MaxRows

            With fpS(0)
                .Row = i
                .Col = 3
                .Text = 1

            End With

        Next i
        
    ElseIf chk.Value = 0 Then

        For i = 1 To fpS(0).MaxRows

            With fpS(0)
                .Row = i
                .Col = 3
                .Text = 0

            End With

        Next i
        
    End If
End Sub

Private Sub Form_Load()
    InitCtrl

End Sub

Private Sub InitCtrl()

    With fpS(0)
        .ReDraw = False
        .MaxCols = 3
        .MaxRows = 0
        .FontBold = True
    
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .Col = 3
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .Lock = False
        
        .SetText 1, 0, "卷盘ID"
        .SetText 2, 0, "仓位ID"
        .SetText 3, 0, "选择"
        
        .ColWidth(1) = 18
        .ColWidth(2) = 10
        .ColWidth(3) = 4
    
        .RowHeight(0) = 20
        .RowHeight(-1) = 15

        .ReDraw = True

    End With
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key

        Case "START"
            txtScan.Visible = True
            txtScan.SetFocus
            
            If txtJobID.Text = "" Then
                Play ("请扫描JOB号")
            ElseIf txtPos.Text = "" Then
                Play ("请扫描仓位号")
            End If
            
        Case "BAND"
            If SaveReelInfo = True Then
    
                Call ListReelInfo(txtJobID.Text)
            End If
        Case "NEXT"
            lblScan.Caption = "扫描JOB"
            txtJobID.Text = ""

        Case "CLOSE"
            Unload Me

    End Select

End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)

    Dim strScan As String

    If KeyAscii <> vbKeyReturn Or Len(Trim(txtScan.Text)) = 0 Then Exit Sub

    strScan = UCase(Trim$(txtScan.Text))
    
    If strScan = "BANGDING" Then
        If SaveReelInfo = True Then
    
            Call ListReelInfo(txtJobID.Text)
            lblScan.Caption = "扫描JOB"
            txtJobID.Text = ""
        End If
        
        txtScan.Text = ""
        Exit Sub
    End If
    
    If txtJobID.Text = "" Then
        Call DoJobID(strScan) ' JobID
    Else
        
        Call DoPosID(strScan) ' 仓位

    End If
    
    txtScan.Text = ""

End Sub

Private Sub DoJobID(strJobID As String)
    Dim i As Integer
    strJobID = Mid$(strJobID, 3)
    If ChkJobID(strJobID) = False Then Exit Sub
    txtJobID.Text = strJobID

    Call ListReelInfo(txtJobID.Text)
    lblScan.Caption = "扫描仓位"
    
    If txtPos.Text = "" Then
        Play ("卷盘已呈现, 请扫描正确的仓位号")
    Else
        Play ("卷盘已呈现, 请勾选需要绑定仓位的卷盘")
    End If
    
    With fpS(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 3
            .Text = "1"
        Next
    End With
    
End Sub

Private Function ChkJobID(strJobID As String) As Boolean

    Dim strSql As String

    ChkJobID = False

    strSql = "select * from customeroitbl_test where TEST_MTRL_DESC = '" & strJobID & "'"

    If Get_OracleCnt(strSql) = 0 Then
        MsgBox "JOBID 错误或不存在, 请确认", vbCritical, "警告"
        Exit Function

    End If

    ChkJobID = True

End Function

Private Sub ListReelInfo(strJobID As String)

    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    rs.Source = "select distinct d.箱号, d.仓位, '' as 选择 from erpbase..tblCustomerOI y  " & _
"inner join ERPBASE..tblmappingData x on convert(varchar(100), y.id) = x.FILENAME and y.SOURCE_BATCH_ID = x.LOTID and y.CUSTOMERSHORTNAME = '37' " & _
"inner join erpdata..tblErpInStockRelation b on SUBSTRING(replace(b.WAFER_ID, b.SFC_ID + ',', ''),1,CHARINDEX('::',replace(b.WAFER_ID, b.SFC_ID + ',', ''),1) - 1) = x.SUBSTRATEID " & _
"inner join erpdata..tblErpInStockDetailInfo c on SUBSTRING(c.KEY_VALUE, 2, 8) =SUBSTRING(REPLACE(B.SFC_ID, 'SFCBO:1020,', ''), 1, 8) and b.BOX_ID = c.BOX_ID " & _
"inner join erpdata..tblStockNumTree d on d.箱号 =  replace(substring(c.KEY_VALUE,1,charindex('|', c.KEY_VALUE)),'|','') " & _
"where c.KEY_NAME = 'CONTAINER_NAME' and y.TEST_MTRL_DESC = '" & strJobID & "' order by d.箱号 "
    
    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    With fpS(0)
        .MaxRows = 0

        If rs.RecordCount > 0 Then
            Set .DataSource = rs
        Else
            MsgBox "查询不到卷盘信息, 请确认", vbInformation, "提示"
            Exit Sub

        End If

    End With

    rs.Close
    Set rs = Nothing

End Sub

Private Function SaveReelInfo() As Boolean
Dim i As Integer
Dim strReelID As String, strPosID As String
Dim bSel As Boolean


SaveReelInfo = False
bSel = False

strPosID = Trim(txtPos.Text)

With fpS(0)
    For i = 1 To .MaxRows
        .Row = i
        .Col = 3
        
        If .Text = "1" Then
            bSel = True
        End If
    Next
End With

If bSel = False Then
    Play ("请勾选需要绑定或更新仓位号的卷盘")
   ' MsgBox "请勾选需要绑定或更新仓位号的卷盘", vbInformation, "提示"
    Exit Function
End If

With fpS(0)
    For i = 1 To .MaxRows
        .Row = i
        .Col = 3
        
        If .Text = "1" Then
            .Col = 1
            strReelID = Trim$(.Text)
            
            .Col = 2
            If Len(Trim(.Text)) <> 0 And UCase$(Trim$(.Text)) <> strPosID Then
                If MsgBox("该卷盘: " & strReelID & " 已绑定仓位: " & .Text & vbCrLf & "是否重置该仓位,并更新为当前扫描仓位", vbYesNo, "提示") = vbYes Then
                    Call UpdateReelInfo(strReelID, strPosID)
              
                End If
            Else
                Call UpdateReelInfo(strReelID, strPosID)
            End If
        
        End If
    
    Next
    
End With

Play ("绑定完成")
SaveReelInfo = True
End Function

Private Sub UpdateReelInfo(strReelID As String, strPosID As String)

Dim strSql As String

strSql = "update erpdata..tblStockNumTree set 仓位 = '" & strPosID & "' where 箱号 = '" & strReelID & "' "
AddSql2 (strSql)

End Sub


Private Sub DoPosID(strPosID As String)
txtPos.Text = strPosID

Play ("仓位号已设置")
End Sub
