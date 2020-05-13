VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_CPVS 
   Caption         =   "37提货核对系统"
   ClientHeight    =   11625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16215
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
   ScaleHeight     =   11625
   ScaleWidth      =   16215
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   10815
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   16245
      _ExtentX        =   28654
      _ExtentY        =   19076
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "提货单与实物外箱核对"
      TabPicture(0)   =   "Form_CPVS.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "media"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDN"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblQBox"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fpSpread1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtScanCode"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDN"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtQBox"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabPicture(1)   =   "Form_CPVS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl2"
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(2)=   "txtTHD"
      Tab(1).Control(3)=   "txtZXD"
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtQBox 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Top             =   893
         Width           =   3015
      End
      Begin VB.TextBox txtDN 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   1380
         Width           =   3015
      End
      Begin VB.TextBox txtZXD 
         BackColor       =   &H00FFC0FF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71280
         TabIndex        =   9
         Top             =   1613
         Width           =   4935
      End
      Begin VB.TextBox txtTHD 
         BackColor       =   &H00FFC0FF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71280
         TabIndex        =   7
         Top             =   900
         Width           =   7815
      End
      Begin VB.TextBox txtScanCode 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   6375
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   8655
         Left            =   1200
         TabIndex        =   2
         Top             =   1920
         Width           =   6375
         _Version        =   524288
         _ExtentX        =   11245
         _ExtentY        =   15266
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
         MaxRows         =   0
         SpreadDesigner  =   "Form_CPVS.frx":0038
         AppearanceStyle =   0
      End
      Begin VB.Label lblQBox 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前外箱箱号"
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
         Left            =   1200
         TabIndex        =   12
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前外箱DN"
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
         Left            =   1200
         TabIndex        =   11
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "装箱单二维码(单个DN那份):"
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
         Left            =   -74280
         TabIndex        =   8
         Top             =   1680
         Width           =   2865
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提货单二维码(多个DN那份):"
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
         Left            =   -74280
         TabIndex        =   6
         Top             =   960
         Width           =   2865
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   8760
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   3135
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
         _cx             =   5530
         _cy             =   873
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫码框:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   420
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12480
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_CPVS.frx":0430
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_CPVS.frx":1082
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   1402
      ButtonWidth     =   2090
      ButtonHeight    =   1349
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "导出核对记录"
            Key             =   "EXPORT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "EXIT"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "Form_CPVS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lSeq   As Long
Dim bFlag3 As Boolean

Rem: 播放音频提醒
Private Sub Play(sFileName As String)
Dim sPath   As String
Dim sSuffix As String
sPath = "\\10.160.1.84\public\media_source\"
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix

End Sub

Private Sub Form_Activate()
txtScanCode.SetFocus

End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
InitFps
Play ("请扫描提货单条码")

End Sub

Private Sub InitFps()

With fpSpread1
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
    .SetText 1, 0, "DN"
    .SetText 2, 0, "目标箱号"
    .SetText 3, 0, "已扫箱号"
    .ColWidth(1) = 15
    .ColWidth(2) = 15
    .ColWidth(3) = 15
    .RowHeight(0) = 20
    .RowHeight(-1) = 15
    .ReDraw = True

End With

bFlag3 = False

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.Tab = 1 Then
MsgBox "此页面功能作废"

End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "EXPORT"
        ExporToExcel ("select * from CPVSTBL order by matchdate desc")

    Case "EXIT"
        Unload Me

End Select

End Sub

Private Sub txtScanCode_KeyPress(KeyAscii As Integer)

If KeyAscii <> vbKeyReturn Then
    Exit Sub

End If

If txtScanCode.Text = "" Then Exit Sub
If fpSpread1.MaxRows = 0 Then
    DoScan (UCase(Trim$(txtScanCode.Text)))
Else
    DoMatch (UCase(Trim$(txtScanCode.Text)))

End If

txtScanCode.Text = ""

End Sub

Private Sub DoScan(strCode As String)
Dim i As Integer
Dim strList
Dim strSql As String
Dim strDN  As String
strList = Split(strCode, "/")

For i = 0 To UBound(strList)
    strDN = strList(i)

    If strDN = "" Then
        Play ("DN号码格式不正确")
        MsgBox "DN号码格式不正确"
        Exit For

    End If

    ' 判断DN是否合法
    strSql = "select * from packing_detailed where dn_num = '" & strDN & "'"

    If Get_OracleCnt(strSql) = 0 Then
        Play ("DN号码不存在")
        MsgBox "DN号码不存在"
        Exit Sub

    End If

    ' 判断DN是否已经检过
    strSql = "select * from CPVSTBL where instr(DNLIST, '" & strDN & "') > 0 and MATCHRES = 'Y'"

    If Get_OracleCnt(strSql) > 0 Then
        MsgBox "该DN号码已经核对了, 请确认是否有误", vbInformation, "警告"
        Exit Sub

    End If

    ListData (strDN)
Next
Play ("提货单已扫描完成,请核对外箱")
'MsgBox "提货单已扫描完成,请核对外箱"
lSeq = Get_OracleNo("select CPVSSEQ.NEXTVAL from dual")
AddSql ("insert into CPVSTBL(DNLIST, MATCHBY,MATCHDATE,MATCHRES,ID) values('" & strCode & "', '" & gUserName & "', sysdate, 'N', '" & lSeq & "')")

End Sub

Private Sub DoMatch(strCode As String)
Dim i      As Integer
Dim bFlag  As Boolean
Dim bFlag2 As Boolean
Dim j      As Integer
bFlag = False
bFlag2 = False

If bFlag3 = True Then
    If strCode = txtDN.Text Then
        Play ("装箱单放置正确,请继续扫描其他外箱")

        With fpSpread1

            For j = 1 To .MaxRows
                .Row = j
                .Col = 2

                If .Text = txtQBox.Text Then
                    .Col = 3
                    .Text = txtQBox.Text

                End If

            Next

        End With

        With fpSpread1

            For i = 1 To .MaxRows
                .Row = i
                .Col = 3

                If .Text = "" Then
                    bFlag2 = True

                End If

            Next

        End With

        If bFlag2 = False Then
            Play ("提货单核对完成,可以出货")
            txtQBox.Text = ""
            txtDN.Text = ""
            MsgBox "提货单核对完成,可以出货", vbInformation, "提示"
            
            AddSql ("update CPVSTBL set MATCHRES = 'Y' where id = '" & lSeq & "'")
            fpSpread1.MaxRows = 0

        End If

        bFlag3 = False
    Else
        Play ("装箱单不正确,请扫描正确的装箱单")
        bFlag3 = True

    End If

    Exit Sub

End If

If Left(strCode, 1) <> "Q" Then
    MsgBox "没有扫描到外箱号", vbInformation, "提示"
    Play ("没有扫描到外箱号")
    Exit Sub

End If

txtQBox.Text = strCode
txtDN.Text = Get_OracleStr("select dn_num from PACKING_DETAILED where CARTON = '" & strCode & "'")

With fpSpread1

    For i = 1 To .MaxRows
        .Row = i
        .Col = 2

        If .Text = strCode Then
            .Col = 3

            If .Text <> "" Then
                Play ("该外箱号已经扫描过,请勿重复扫描")
                MsgBox "该外箱号已经扫描过,请勿重复扫描", vbInformation, "提示"
                bFlag = True
                Exit For

            End If

            '检查装箱单
            If Get_OracleStr("SELECT kid FROM PACKING_DETAILED WHERE CARTON = '" & strCode & "'") = "K1" Then
                Play ("该外箱是第一箱,请扫描装箱单二维码")
                bFlag3 = True
                bFlag = True
                Exit For

            End If

            .Text = strCode
            Play ("外箱扫描正确")
            bFlag = True
            Exit For

        End If

    Next

End With

If bFlag = False Then
    Play ("外箱号不正确")
    MsgBox "外箱号不正确"

End If

With fpSpread1

    For i = 1 To .MaxRows
        .Row = i
        .Col = 3

        If .Text = "" Then
            bFlag2 = True

        End If

    Next

End With

If bFlag2 = False Then
    Play ("提货单核对完成,可以出货")
    txtQBox.Text = ""
    txtDN.Text = ""
    MsgBox "提货单核对完成,可以出货", vbInformation, "提示"
    AddSql ("update CPVSTBL set MATCHRES = 'Y' where id = '" & lSeq & "'")
    fpSpread1.MaxRows = 0

End If

End Sub

Private Sub ListData(strDN As String)
Dim i      As Integer
Dim strSql As String
Dim sOra   As String
Dim rs     As New ADODB.Recordset
sOra = "select distinct dn_num,carton,'' from packing_detailed where dn_num = '" & strDN & "'"
Set rs = Get_OracleRs(sOra)

With fpSpread1

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            .Text = rs(0)
            .Col = 2
            .Text = rs(1)
            .Col = 3
            .Text = "" & rs(2)
            rs.MoveNext
        Next

    End If

End With

End Sub

Private Sub txtTHD_KeyPress(KeyAscii As Integer)

If KeyAscii <> vbKeyReturn Then Exit Sub
If Len(Trim$(txtTHD.Text)) = 0 Then Exit Sub
Play ("提货单二维码已扫描,请扫描装箱单二维码")
txtTHD.Enabled = False
txtZXD.Enabled = True
txtZXD.SetFocus

End Sub

Private Sub txtZXD_KeyPress(KeyAscii As Integer)

If KeyAscii <> vbKeyReturn Then Exit Sub
If Len(Trim$(txtZXD.Text)) = 0 Then Exit Sub
If InStr(Trim(txtTHD.Text), Trim(txtZXD.Text)) > 0 Then
    Play ("装箱单正确")
    txtTHD.Text = Replace(txtTHD.Text, Trim$(txtZXD.Text), "")
    txtZXD.Text = ""
Else
    MsgBox "装箱单:" & txtZXD.Text & "和提货单:" & txtTHD.Text & "不匹配,请确认是否放错文件", vbCritical, "警告"
    txtZXD.Text = ""
    Exit Sub

End If

If Trim$(Replace(txtTHD.Text, "/", "")) = "" Then
    Play ("提货单和装箱单信息一致,核对完毕")
    txtTHD.Enabled = True
    txtTHD.Text = ""
    txtZXD.Enabled = False
    txtZXD.Text = ""

End If

End Sub
