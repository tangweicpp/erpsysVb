VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.ocx"
Begin VB.Form Frm_Label_Checking_System 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "通用标签核对系统(GLCS)_二维码"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19110
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
   ScaleHeight     =   9570
   ScaleWidth      =   19110
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   11775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      Begin VB.TextBox txtNXQty 
         Height          =   285
         Left            =   11280
         TabIndex        =   25
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtWXQty 
         Height          =   285
         Left            =   11280
         TabIndex        =   24
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtIbCnt 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtLvCnt 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   4560
         TabIndex        =   22
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtPackingQtyAdd 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   4275
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtPackingQty 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtPackingNO 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3465
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtBoxID 
         BackColor       =   &H00FFC0FF&
         Height          =   405
         Left            =   7560
         TabIndex        =   14
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CheckBox chk 
         Caption         =   "卷盘"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00C0C0C0&
         Caption         =   "删除核对箱号记录"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton cmdUpload 
         BackColor       =   &H00C0C0C0&
         Caption         =   "导出核对记录"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cmbCombo1 
         BackColor       =   &H00FFC0FF&
         Height          =   315
         ItemData        =   "Frm_Label_Checking_System.frx":0000
         Left            =   1200
         List            =   "Frm_Label_Checking_System.frx":0002
         TabIndex        =   7
         Top             =   1185
         Width           =   2775
      End
      Begin VB.CheckBox chk 
         Caption         =   "铝箔袋"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H80000004&
         Caption         =   "内箱"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "外箱"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtScan 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   9375
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00C0C0C0&
         Caption         =   "开始扫码"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   6855
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Width           =   14655
         _Version        =   524288
         _ExtentX        =   25850
         _ExtentY        =   12091
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
         SpreadDesigner  =   "Frm_Label_Checking_System.frx":0004
         Appearance      =   1
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lblPackingQtyAdd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "累计数量:"
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
         Left            =   690
         TabIndex        =   20
         Top             =   4320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPackingQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前外箱数量:"
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
         TabIndex        =   18
         Top             =   3840
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblPackingNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前外箱箱号:"
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
         TabIndex        =   16
         Top             =   3480
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lbl222 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "箱号"
         Height          =   195
         Left            =   6960
         TabIndex        =   15
         Top             =   1785
         Width           =   360
      End
      Begin VB.Line Line1 
         X1              =   8520
         X2              =   8520
         Y1              =   1440
         Y2              =   1680
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   615
         Left            =   480
         TabIndex        =   12
         Top             =   9960
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "核对类型"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   960
      End
   End
End
Attribute VB_Name = "Frm_Label_Checking_System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : Frm_Label_Checking_System
'    Project    : 正式工程1
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Enum E_Lbl

    e_CARTON = 1
    E_BOX
    E_Reel

End Enum

Enum E_CheckStatus

    E_NO_CHECKED
    E_CARTON_CHECKED
    E_ALL_CHECKED

End Enum

Dim strLblInfo()  As LBL_WAFER_INFO
Dim gCntRow       As Integer
Dim gNoCheckRow() As Integer
Dim gUniqueRow()  As Integer
Dim gSplitFlag    As String
Dim gStatus       As Integer
Dim gMaxRow       As Integer
Dim gID           As Long
Dim gLVCntSum     As Long
Dim gIBCntSum     As Long
Dim strPart_C()   As String, strPart_B() As String, strPart_R() As String
Dim strBoxID      As String
Dim lWXQty As Long
Dim lNXQty As Long


Private Sub cmbCombo1_Click()

Select Case cmbCombo1.text

    Case "HK037"
        gSplitFlag = ";"
        gMaxRow = 12
        ReDim gNoCheckRow(3)
        gNoCheckRow(0) = 1
        gNoCheckRow(1) = 6
        gNoCheckRow(2) = 9
        gNoCheckRow(3) = 11
        gCntRow = 5
        ReDim gUniqueRow(1)
        gUniqueRow(0) = 10

    Case "DA69"
        gSplitFlag = ";"
        gMaxRow = 7
        ReDim gNoCheckRow(1)
        gNoCheckRow(0) = 0
        gCntRow = 4
        ReDim gUniqueRow(1)
        gUniqueRow(0) = 5

    Case "AB18"
        gSplitFlag = "+"
        gMaxRow = 12
        ReDim gNoCheckRow(2)
        gNoCheckRow(0) = 10
        gNoCheckRow(1) = 11
        gCntRow = 2
        ReDim gUniqueRow(1)
        gUniqueRow(0) = 12

    Case "HK037_铝箔袋_卷盘"
        gSplitFlag = ";"
        gMaxRow = 12
        ReDim gUniqueRow(1)
        gUniqueRow(0) = 9

        With Fps(0)
            .Col = -1
            .Row = -1
            .Lock = True
            .Col = 1
            .Row = 0
            .FontSize = 10
            .Col = 2
            .Row = 0
            .FontSize = 10
            .SetText 1, 0, "铝箔袋"
            .SetText 2, 0, "卷盘"
            .ColWidth(1) = 31
            .ColWidth(2) = 31

        End With

    Case "SH50"
        gSplitFlag = "@"
        gMaxRow = 7

End Select

Select Case cmbCombo1.text

    Case "GC"
        txtPackingNO.Visible = True
        lblPackingNO.Visible = True
        txtPackingQty.Visible = True
        lblPackingQty.Visible = True
        txtPackingQtyAdd.Visible = True
        lblPackingQtyAdd.Visible = True
        Fps(0).Visible = False

    Case Else
        txtPackingNO.Visible = False
        lblPackingNO.Visible = False
        txtPackingQty.Visible = False
        lblPackingQty.Visible = False
        txtPackingQtyAdd.Visible = False
        lblPackingQtyAdd.Visible = False
        Fps(0).Visible = True

End Select

If cmbCombo1.text = "US026" Then

    With Fps(0)
        .Col = -1
        .Row = -1
        .Lock = True
        .MaxCols = 6
        .TypeMaxEditLen = 5000
        .Col = 1
        .Row = 0
        .FontSize = 10
        .Col = 2
        .Row = 0
        .FontSize = 10
        .Col = 3
        .Row = 0
        .FontSize = 10
        .SetText 1, 0, "Device"
        .SetText 2, 0, "Wafer Lot"
        .SetText 3, 0, "Wafer ID"
        .SetText 4, 0, "Die Qty"
        .SetText 5, 0, "Date Code"
        .SetText 6, 0, "HT LotID"
        .ColWidth(1) = 20
        .ColWidth(2) = 10
        .ColWidth(3) = 10

    End With

Else

    With Fps(0)
        .Col = -1
        .Row = -1
        .Lock = True
        .MaxCols = 3
        .TypeMaxEditLen = 5000
        .Col = 1
        .Row = 0
        .FontSize = 10
        .Col = 2
        .Row = 0
        .FontSize = 10
        .Col = 3
        .Row = 0
        .FontSize = 10
        .SetText 1, 0, "外  箱(C)"
        .SetText 2, 0, "内  箱(B)"
        .SetText 3, 0, "铝箔袋(R)"
        .ColWidth(1) = 31
        .ColWidth(2) = 31
        .ColWidth(3) = 31
        .Row = 5

    End With

End If

End Sub

Private Sub CmdClear_Click()
Dim strBoxID As String
If cmbCombo1.text = "" Then
    MsgBox "请选择模板", vbInformation, "提示"
    Exit Sub

End If

Select Case cmbCombo1.text

    Case "HK037_铝箔袋_卷盘"

        If InStr("07885", gUserName) > 0 Then
            
            strBoxID = UCase$(Trim$(txtBoxID.text))

            If Len(strBoxID) = 0 Then
                MsgBox "请输入要删除的历史箱号", vbInformation, "提示"
                Exit Sub

            End If

            AddSql ("insert into unique_tbl_bak select * from unique_tbl where key_value = '" & strBoxID & "' ")
            AddSql ("delete from unique_tbl where key_value = '" & strBoxID & "' ")
            MsgBox "历史记录已经清空", vbInformation, "提示"
        Else
            MsgBox "你没有删除的权限,请联系IT进行删除", vbInformation, "警告"
            Exit Sub

        End If

    Case "DA24"
        
        strBoxID = UCase$(Trim$(txtBoxID.text))

        If Len(strBoxID) = 0 Then
            MsgBox "请输入要删除的历史箱号", vbInformation, "提示"
            Exit Sub

        End If

        AddSql ("delete from unique_tbl_new where KEYFROM = 'DA24' and KEYNAME='PACKNO' and keyvalue = '" & strBoxID & "' ")
        MsgBox "历史箱号已删除", vbInformation, "提示"
        Exit Sub

    Case "HD"
        strBoxID = UCase$(Trim$(txtBoxID.text))

        If Len(strBoxID) = 0 Then
            MsgBox "请输入要删除的历史箱号", vbInformation, "提示"
            Exit Sub

        End If

        AddSql ("delete from UNIQUE_TBL_NEW where KEYFROM = 'HD' and keyvalue = '" & strBoxID & "' ")
        MsgBox "历史箱号:" & strBoxID & " 已删除", vbInformation, "提示"
        Exit Sub

    Case "GC"
        MsgBox "箱号不可删除", vbInformation, "提示"

    Case "SH50"
        MsgBox "该客户没有箱号可以删除", vbInformation, "提示"

    Case Else

        If InStr("07885", gUserName) > 0 Then
            
            strBoxID = UCase$(Trim$(txtBoxID.text))

            If Len(strBoxID) = 0 Then
                MsgBox "请输入要删除的历史箱号", vbInformation, "提示"
                Exit Sub

            End If

            AddSql ("insert into unique_tbl_bak select * from unique_tbl where key_value = '" & strBoxID & "' ")
            AddSql ("delete from unique_tbl where key_value = '" & strBoxID & "' ")
            MsgBox "历史记录已经清空", vbInformation, "提示"
        Else
            MsgBox "你没有删除的权限,请联系IT进行删除", vbInformation, "警告"
            Exit Sub

        End If
End Select

End Sub

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdStart_Click()
If cmbCombo1.text = "" Then
    MsgBox "请选择核对类型", vbInformation, "提示"
    Exit Sub

End If

InitCheckStatus

End Sub

Private Sub InitCheckStatus()
gStatus = E_CheckStatus.E_NO_CHECKED
Fps(0).MaxRows = 0
chk(0).Value = 0
chk(1).Value = 0
chk(2).Value = 0
chk(3).Value = 0
gLVCntSum = 0
gIBCntSum = 0
txtIbCnt.text = gIBCntSum
txtLvCnt.text = gLVCntSum
txtScan.Visible = True
txtScan.SetFocus
lWXQty = 0
lNXQty = 0

Select Case cmbCombo1.text

    Case "HK037_铝箔袋_卷盘"

    Case "DA24"
        'Play ("请依次扫描外箱,内箱,铝箔袋的二维码标签")
        Fps(0).MaxRows = 7

    Case "GC"
        If txtPackingNO.text = "" Then
            Play ("请扫描GC的外箱标签二维码")

        End If

    Case "SH50"
        Play ("请依次扫描外箱,内箱,铝箔袋的二维码标签")
        Fps(0).MaxRows = 7

    Case "HD"
        Play ("请扫描外箱标签二维码")
        Fps(0).MaxRows = 12

    Case "US026"
        Play ("请扫描外箱标签二维码")
      
    Case Else

End Select

Dim i As Integer

Erase strPart_C
Erase strPart_B
Erase strPart_R

End Sub

Private Sub NextCheckStatus()
gStatus = E_CheckStatus.E_CARTON_CHECKED
chk(1).Value = 0
chk(2).Value = 0
txtScan.Visible = True
txtScan.SetFocus

End Sub

Private Sub cmdUpload_Click()
If cmbCombo1.text = "" Then
    MsgBox "请选择类型", vbInformation, "提示"
    Exit Sub

End If

Select Case cmbCombo1.text

    Case "HK037_铝箔袋_卷盘"
        ExporToExcel ("select * from unique_tbl order by update_time desc")

    Case "DA24"
        ExporToExcel ("select KEYFROM as 标签客户, KEYNAME as 标签类型, KEYVALUE as 标签值, KEYTIME as 日期, KEYBY as 人员  from UNIQUE_TBL_NEW order by KEYTIME desc")

    Case "GC"
        ExporToExcel ("select KEYFROM as 客户, KEYNAME as 外箱, KEYVALUE as 箱号, KEYTIME as 核对日期, KEYBY as 核对人员  from UNIQUE_TBL_NEW where KEYFROM = 'GC' order by KEYTIME desc")

    Case "SH50"
        MsgBox "该客户没有箱号可以导出", vbInformation, "提示"

    Case "HD"
        ExporToExcel ("select KEYFROM as 客户, KEYNAME as 外箱, KEYVALUE as 箱号, KEYTIME as 核对日期, KEYBY as 核对人员  from UNIQUE_TBL_NEW where KEYFROM = 'HD' order by KEYTIME desc")

    Case Else
        ExporToExcel ("select * from unique_tbl order by update_time desc")

End Select

End Sub

Private Sub Form_Load()
InitCtrls

End Sub

Private Sub InitCtrls()
txtPackingQtyAdd.text = 0

With Fps(0)
    .Col = -1
    .Row = -1
    .Lock = True
    .TypeMaxEditLen = 5000
    .Col = 1
    .Row = 0
    .FontSize = 10
    .Col = 2
    .Row = 0
    .FontSize = 10
    .Col = 3
    .Row = 0
    .FontSize = 10
    .SetText 1, 0, "外  箱(C)"
    .SetText 2, 0, "内  箱(B)"
    .SetText 3, 0, "铝箔袋(R)"
    .ColWidth(1) = 31
    .ColWidth(2) = 31
    .ColWidth(3) = 31
    .Row = 5

End With

cmbCombo1.AddItem ("HK037")
cmbCombo1.AddItem ("DA69")
cmbCombo1.AddItem ("AB18")
cmbCombo1.AddItem ("HK037_铝箔袋_卷盘")
cmbCombo1.AddItem ("DA24")
cmbCombo1.AddItem ("GC")
cmbCombo1.AddItem ("SH50")
cmbCombo1.AddItem ("HD")
cmbCombo1.AddItem ("US026")

End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Or txtScan.text = "" Then Exit Sub
Call CheckHandle(UCase$(Trim$(txtScan.text)))
txtScan.text = ""

End Sub

Private Sub CheckHandle(strCode As String)

Select Case cmbCombo1.text

    Case "HK037_铝箔袋_卷盘"
        ListData_HK037 (strCode)

    Case "HK037----"
        ListData_HK037_2 (strCode)

    Case "DA24"
        ListData_DA24 (strCode)

    Case "GC"
        verifyLbl_GC (strCode)

    Case "SH50"
        ListData_SH50 (strCode)

    Case "HD"
        ListData_HD (strCode)

    Case "US026"
        ListData_US026 (strCode)

    Case Else
        ListData (strCode)

End Select

End Sub

Private Sub ListData(strCode As String)
Dim strPart() As String, i As Integer, lTmp As Long

strPart = Split(strCode, gSplitFlag)
If gMaxRow <> UBound(strPart) + 1 Then
    MsgBox "请扫描正确的二维码", vbInformation, "提示"
    Exit Sub

End If

If chk(0).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "O" Then
            MsgBox "外箱标签二维码:首位O字符错误:", vbCritical, "警告"
            Exit Sub

        End If

        If strPart(8) <> "000004" Then
            MsgBox "请扫描000004外箱标签", vbInformation, "警告"
            Exit Sub

        End If

    End If

    With Fps(0)

        For i = 0 To UBound(strPart)
            .MaxRows = .MaxRows + 1
            If i = gCntRow - 1 Then
                .SetText E_Lbl.e_CARTON, i + 1, Replace(strPart(i), "Q", "")
            Else
                .SetText E_Lbl.e_CARTON, i + 1, strPart(i)

            End If

        Next

    End With

    If InStr(strPart(gUniqueRow(0) - 1), "-C") = 0 Then
        Fps(0).Row = gUniqueRow(0)
        Fps(0).Col = 1
        Fps(0).BackColor = vbRed
        MsgBox "外箱请扫描-C标签", vbInformation, "提示"
        Exit Sub

    End If

    chk(0).Value = 1
    Play ("外箱标签已扫描")
ElseIf chk(1).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "I" Then
            MsgBox "内箱标签二维码:首位I字符错误:", vbCritical, "警告"
            Exit Sub

        End If

        If strPart(8) <> "000003" Then
            MsgBox "请扫描000003内盒标签", vbInformation, "警告"
            Exit Sub

        End If

    End If

    With Fps(0)

        For i = 0 To UBound(strPart)
            If i = gCntRow - 1 Then
                .Row = gCntRow
                .Col = 2
                If .text = "" Then
                    lTmp = 0
                Else
                    lTmp = CLng(.text)

                End If

                .SetText E_Lbl.E_BOX, i + 1, CLng(Replace(strPart(i), "Q", "")) + lTmp
            Else
                .SetText E_Lbl.E_BOX, i + 1, strPart(i)

            End If

        Next i

    End With

    '    If InStr(strPart(gUniqueRow(0) - 1), "-B") = 0 Then
    '        Fps(0).Row = gUniqueRow(0)
    '        Fps(0).Col = 2
    '        Fps(0).BackColor = vbRed
    '        MsgBox "内箱请扫描-B标签", vbInformation, "提示"
    '        Exit Sub
    '
    '    End If
    chk(1).Value = 1
    Play ("内箱标签已扫描")
ElseIf chk(2).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "I" Then
            MsgBox "铝箔袋标签二维码:首位I字符错误:", vbCritical, "警告"
            Exit Sub

        End If

        If strPart(8) <> "000002" Then
            MsgBox "请扫描000002铝箔袋标签", vbInformation, "警告"
            Exit Sub

        End If

    End If

    With Fps(0)

        For i = 0 To UBound(strPart)
            If i = gCntRow - 1 Then
                .Row = gCntRow
                .Col = 3
                If .text = "" Then
                    lTmp = 0
                Else
                    lTmp = CLng(.text)

                End If

                .SetText E_Lbl.E_Reel, i + 1, CLng(Replace(strPart(i), "Q", "")) + lTmp
            Else
                .SetText E_Lbl.E_Reel, i + 1, strPart(i)

            End If

        Next i

    End With

    If InStr(strPart(gUniqueRow(0) - 1), "-R") = 0 Then
        Fps(0).Row = gUniqueRow(0)
        Fps(0).Col = 3
        Fps(0).BackColor = vbRed
        MsgBox "铝箔袋请扫描-R标签", vbInformation, "提示"
        Exit Sub

    End If

    chk(2).Value = 1
    Play ("铝箔袋标签已扫描")
    gID = Get_OracleStr("select UNIQUE_SEQ.NEXTVAL from dual")
    '开始核对
    CheckData
Else

End If

End Sub

Private Sub ListData_HK037_2(strCode As String)
Dim strPart() As String, i As Integer, lTmp As Long

strPart = Split(strCode, gSplitFlag)
If gMaxRow <> UBound(strPart) + 1 Then
    MsgBox "请扫描正确的二维码", vbInformation, "提示"
    Exit Sub

End If

If chk(0).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "O" Then
            MsgBox "外箱标签二维码:首位O字符错误:", vbCritical, "警告"
            Exit Sub

        End If

    End If

    With Fps(0)

        For i = 0 To UBound(strPart)
            .MaxRows = .MaxRows + 1
            If i = gCntRow - 1 Then
                .SetText E_Lbl.e_CARTON, i + 1, Replace(strPart(i), "Q", "")
            Else
                .SetText E_Lbl.e_CARTON, i + 1, strPart(i)

            End If

        Next

    End With

    If InStr(strPart(gUniqueRow(0) - 1), "-C") = 0 Then
        Fps(0).Row = gUniqueRow(0)
        Fps(0).Col = 1
        Fps(0).BackColor = vbRed
        MsgBox "外箱请扫描-C标签", vbInformation, "提示"
        Exit Sub

    End If

    chk(0).Value = 1
    Play ("外箱标签已扫描")
ElseIf chk(1).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "I" Then
            MsgBox "内箱标签二维码:首位I字符错误:", vbCritical, "警告"
            Exit Sub

        End If

    End If

    With Fps(0)

        For i = 0 To UBound(strPart)
            If i = gCntRow - 1 Then
                .Row = gCntRow
                .Col = 2
                If .text = "" Then
                    lTmp = 0
                Else
                    lTmp = CLng(.text)

                End If

                .SetText E_Lbl.E_BOX, i + 1, (Replace(strPart(i), "Q", ""))
            Else
                .SetText E_Lbl.E_BOX, i + 1, strPart(i)

            End If

        Next i

    End With

    If InStr(strPart(gUniqueRow(0) - 1), "-B") = 0 Then
        Fps(0).Row = gUniqueRow(0)
        Fps(0).Col = 2
        Fps(0).BackColor = vbRed
        MsgBox "内箱请扫描-B标签", vbInformation, "提示"
        Exit Sub

    End If

    chk(1).Value = 1
    Play ("内箱标签已扫描")
ElseIf chk(2).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "I" Then
            MsgBox "铝箔袋标签二维码:首位I字符错误:", vbCritical, "警告"
            Exit Sub

        End If

    End If

    With Fps(0)

        For i = 0 To UBound(strPart)
            If i = gCntRow - 1 Then
                .Row = gCntRow
                .Col = 3
                If .text = "" Then
                    lTmp = 0
                Else
                    lTmp = CLng(.text)

                End If

                '.SetText E_Lbl.E_Reel, I + 1, CLng(Replace(strPart(I), "Q", "")) + lTmp
                .SetText E_Lbl.E_Reel, i + 1, (Replace(strPart(i), "Q", ""))
            Else
                .SetText E_Lbl.E_Reel, i + 1, strPart(i)

            End If

        Next i

    End With

    If InStr(strPart(gUniqueRow(0) - 1), "-R") = 0 Then
        Fps(0).Row = gUniqueRow(0)
        Fps(0).Col = 3
        Fps(0).BackColor = vbRed
        MsgBox "铝箔袋请扫描-R标签", vbInformation, "提示"
        Exit Sub

    End If

    chk(2).Value = 1
    Play ("铝箔袋标签已扫描")
    gID = Get_OracleStr("select UNIQUE_SEQ.NEXTVAL from dual")
    '开始核对
    CheckData2
Else

End If

End Sub

Private Sub ListData_HK037(strCode As String)
Dim strPart() As String, i As Integer, lTmp As Long

strPart = Split(strCode, gSplitFlag)
If gMaxRow <> UBound(strPart) + 1 Then
    MsgBox "请扫描正确的二维码", vbInformation, "提示"
    Exit Sub

End If

If chk(2).Value = 0 Then
    If strPart(0) <> "I" Then
        MsgBox "铝箔袋标签二维码:首位I字符错误:", vbCritical, "警告"
        Exit Sub

    End If

    With Fps(0)

        For i = 0 To UBound(strPart)
            .MaxRows = .MaxRows + 1
            .SetText E_Lbl.e_CARTON, i + 1, strPart(i)
        Next i

    End With

    chk(2).Value = 1
    Play ("铝箔袋标签已扫描")
ElseIf chk(3).Value = 0 Then
    If strPart(0) <> "I" Then
        MsgBox "卷盘标签二维码:首位I字符错误:", vbCritical, "警告"
        Exit Sub

    End If

    With Fps(0)

        For i = 0 To UBound(strPart)
            .MaxRows = .MaxRows + 1
            .SetText E_Lbl.E_BOX, i + 1, strPart(i)
        Next i

    End With

    chk(3).Value = 1
    Play ("卷盘标签已扫描")
    CheckData_HK037
Else

End If

End Sub

Private Sub ListData_DA24(strCode As String)
If chk(0).Value = 0 Then ' 1.外箱(C)
    strPart_C = Split(strCode, ";")
    If UBound(strPart_C) <> 5 Then
        MsgBox "外箱二维码不正确", vbInformation, "提示"
        Exit Sub

    End If

    strPart_C(2) = Replace$(strPart_C(2), "PCS", "")

    With Fps(0)
        .SetText 1, 1, strPart_C(0)
        .SetText 1, 2, strPart_C(1)
        .SetText 1, 3, strPart_C(2)
        .SetText 1, 4, strPart_C(3)
        .SetText 1, 5, ""
        .SetText 1, 6, strPart_C(4)
        .SetText 1, 7, strPart_C(5)
        If InStr(strPart_C(5), "-C") = 0 Then
            .Col = 1
            .Row = 7
            .BackColor = vbRed
            MsgBox "请扫描包含-C的外箱二维码标签", vbInformation, "提示"
            Exit Sub

        End If

        If Get_OracleCnt("select * from UNIQUE_TBL_NEW where keyfrom = 'DA24' and keyname = 'PACKNO' and KEYVALUE = '" & strPart_C(5) & "'") > 0 Then
            .Col = 1
            .Row = 7
            .BackColor = vbRed
            MsgBox "系统已经存在同一外箱箱号", vbInformation, "提示"
            Exit Sub
        Else
            AddSql ("insert into UNIQUE_TBL_NEW(KEYNAME, KEYVALUE, KEYFROM,KEYTIME,KEYBY) values('PACKNO','" & strPart_C(5) & "','DA24',sysdate, '" & gUserName & "')")

        End If

    End With

    chk(0).Value = 1
    Play ("外箱标签已扫描,请扫描内箱标签")
ElseIf chk(1).Value = 0 Then    ' 2.内箱(B)
    strPart_B = Split(strCode, ";")
    If UBound(strPart_B) <> 6 Then
        MsgBox "内箱二维码不正确", vbInformation, "提示"
        Exit Sub

    End If

    strPart_B(2) = Replace$(strPart_B(2), "PCS", "")

    With Fps(0)
        .Col = 2
        .Row = 3
        If .text <> "" Then
            strPart_B(2) = CLng(.text) + CLng(strPart_B(2))
        Else
            strPart_B(2) = CLng(strPart_B(2))

        End If

        .SetText 2, 1, strPart_B(0)
        .SetText 2, 2, strPart_B(1)
        .SetText 2, 3, strPart_B(2)
        .SetText 2, 4, strPart_B(3)
        .SetText 2, 5, strPart_B(4)
        .SetText 2, 6, strPart_B(5)
        .SetText 2, 7, strPart_B(6)
        If strPart_B(0) <> strPart_C(0) Then
            '.Col = 2
            .Row = 1
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_B(1) <> strPart_C(1) Then
            '.Col = 2
            .Row = 2
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If InStr(strPart_C(4), strPart_B(5)) = 0 Then
            '.Col = 2
            .Row = 6
            .BackColor = vbRed
            MsgBox "标签包含关系不一致", vbInformation, "提示"

        End If

        If InStr(strPart_B(6), "-B") = 0 Then
            .Col = 2
            .Row = 7
            .BackColor = vbRed
            MsgBox "请扫描包含-B的内箱二维码标签", vbInformation, "提示"
            Exit Sub

        End If

        If CLng(strPart_B(2)) > CLng(strPart_C(2)) Then
            .Row = 3
            .BackColor = vbRed
            MsgBox "内箱数量不能大于外箱数量,出错", vbInformation, "提示"
            Exit Sub

        End If

        If Get_OracleCnt("select * from UNIQUE_TBL_NEW where keyfrom = 'DA24' and keyname = 'PACKNO' and KEYVALUE = '" & strPart_B(6) & "'") > 0 Then
            .Col = 2
            .Row = 7
            .BackColor = vbRed
            MsgBox "系统已经存在同一内箱箱号", vbInformation, "提示"
            Exit Sub
        Else
            AddSql ("insert into UNIQUE_TBL_NEW(KEYNAME, KEYVALUE, KEYFROM,KEYTIME,KEYBY) values('PACKNO','" & strPart_B(6) & "','DA24',sysdate, '" & gUserName & "')")

        End If

        chk(1).Value = 1
        Play ("内箱标签已扫描,请扫描铝箔袋标签")

    End With

ElseIf chk(2).Value = 0 Then    ' 3.铝箔袋(R)
    strPart_R = Split(strCode, ";")
    If UBound(strPart_R) <> 6 Then
        MsgBox "内箱二维码不正确", vbInformation, "提示"
        Exit Sub

    End If

    strPart_R(2) = Replace$(strPart_R(2), "PCS", "")

    With Fps(0)
        .Col = 3
        .Row = 3
        If .text <> "" Then
            strPart_R(2) = CLng(.text) + CLng(strPart_R(2))
        Else
            strPart_R(2) = CLng(strPart_R(2))

        End If

        .SetText 3, 1, strPart_R(0)
        .SetText 3, 2, strPart_R(1)
        .SetText 3, 3, strPart_R(2)
        .SetText 3, 4, strPart_R(3)
        .SetText 3, 5, strPart_R(4)
        .SetText 3, 6, strPart_R(5)
        .SetText 3, 7, strPart_R(6)
        If strPart_R(0) <> strPart_B(0) Then
            '.Col = 3
            .Row = 1
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_R(1) <> strPart_B(1) Then
            '.Col = 3
            .Row = 2
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_R(2) <> strPart_B(2) Then
            .Row = 3
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If InStr(strPart_C(4), strPart_R(5)) = 0 Then
            '.Col = 3
            .Row = 6
            .BackColor = vbRed
            MsgBox "标签包含关系不一致", vbInformation, "提示"

        End If

        If InStr(strPart_R(6), "-R") = 0 Then
            .Col = 3
            .Row = 7
            .BackColor = vbRed
            MsgBox "请扫描包含-R的铝箔袋二维码标签", vbInformation, "提示"
            Exit Sub

        End If

        If Get_OracleCnt("select * from UNIQUE_TBL_NEW where keyfrom = 'DA24' and keyname = 'PACKNO' and KEYVALUE = '" & strPart_R(6) & "'") > 0 Then
            .Col = 3
            .Row = 7
            .BackColor = vbRed
            MsgBox "系统已经存在同一铝箔袋箱号", vbInformation, "提示"
            Exit Sub
        Else
            AddSql ("insert into UNIQUE_TBL_NEW(KEYNAME, KEYVALUE, KEYFROM,KEYTIME,KEYBY) values('PACKNO','" & strPart_R(6) & "','DA24',sysdate, '" & gUserName & "')")

        End If

        If CLng(strPart_R(2)) = CLng(strPart_C(2)) Then
            InitCheckStatus
            Play ("该外箱已全部比对完成,请继续比对其他外箱")
        Else
            chk(1).Value = 0
            chk(2).Value = 0
            Play ("铝箔袋已比对完,数量不足, 请继续比对下一个内盒")

        End If

    End With

End If

End Sub

Private Sub ListData_SH50(strCode As String)
If chk(0).Value = 0 Then ' 1.外箱(C)
    strPart_C = Split(strCode, "@")
    If UBound(strPart_C) <> 6 Then
        MsgBox "外箱二维码不正确", vbInformation, "提示"
        Exit Sub

    End If

    strPart_C(3) = Replace$(strPart_C(3), "PCS", "")

    With Fps(0)
        .SetText 1, 1, strPart_C(0)
        .SetText 1, 2, strPart_C(1)
        .SetText 1, 3, strPart_C(2)
        .SetText 1, 4, strPart_C(3)
        .SetText 1, 5, strPart_C(4)
        .SetText 1, 6, strPart_C(5)
        .SetText 1, 7, strPart_C(6)

    End With

    chk(0).Value = 1
    Play ("外箱标签已扫描,请扫描内箱标签")
ElseIf chk(1).Value = 0 Then    ' 2.内箱(B)
    strPart_B = Split(strCode, "@")
    If UBound(strPart_B) <> 6 Then
        MsgBox "内箱二维码不正确", vbInformation, "提示"
        Exit Sub

    End If

    strPart_B(3) = Replace$(strPart_B(3), "PCS", "")

    With Fps(0)
        .Col = 2
        .Row = 4
        If .text <> "" Then
            strPart_B(3) = CLng(.text) + CLng(strPart_B(3))
        Else
            strPart_B(3) = CLng(strPart_B(3))

        End If

        .SetText 2, 1, strPart_B(0)
        .SetText 2, 2, strPart_B(1)
        .SetText 2, 3, strPart_B(2)
        .SetText 2, 4, strPart_B(3)
        .SetText 2, 5, strPart_B(4)
        .SetText 2, 6, strPart_B(5)
        .SetText 2, 7, strPart_B(6)
        If strPart_B(0) <> strPart_C(0) Then
            '.Col = 2
            .Row = 1
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_B(1) <> strPart_C(1) Then
            '.Col = 2
            .Row = 2
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_B(2) <> strPart_C(2) Then
            '.Col = 2
            .Row = 3
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_B(4) <> strPart_C(4) Then
            '.Col = 2
            .Row = 5
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_B(5) <> strPart_C(5) Then
            '.Col = 2
            .Row = 6
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_B(6) <> strPart_C(6) Then
            '.Col = 2
            .Row = 7
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If CLng(strPart_B(3)) > CLng(strPart_C(3)) Then
            .Row = 4
            .BackColor = vbRed
            MsgBox "内箱数量不能大于外箱数量,出错", vbInformation, "提示"
            Exit Sub

        End If

        chk(1).Value = 1
        Play ("内箱标签已扫描,请扫描铝箔袋标签")

    End With

ElseIf chk(2).Value = 0 Then    ' 3.铝箔袋(R)
    strPart_R = Split(strCode, "@")
    If UBound(strPart_R) <> 6 Then
        MsgBox "内箱二维码不正确", vbInformation, "提示"
        Exit Sub

    End If

    strPart_R(3) = Replace$(strPart_R(3), "PCS", "")

    With Fps(0)
        .Col = 3
        .Row = 4
        If .text <> "" Then
            strPart_R(3) = CLng(.text) + CLng(strPart_R(3))
        Else
            strPart_R(3) = CLng(strPart_R(3))

        End If

        .SetText 3, 1, strPart_R(0)
        .SetText 3, 2, strPart_R(1)
        .SetText 3, 3, strPart_R(2)
        .SetText 3, 4, strPart_R(3)
        .SetText 3, 5, strPart_R(4)
        .SetText 3, 6, strPart_R(5)
        .SetText 3, 7, strPart_R(6)
        If strPart_R(0) <> strPart_B(0) Then
            .Row = 1
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_R(1) <> strPart_B(1) Then
            .Row = 2
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_R(2) <> strPart_B(2) Then
            .Row = 3
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If CLng(strPart_R(3)) <> CLng(strPart_B(3)) Then
            .Row = 4
            .BackColor = vbRed
            MsgBox "铝箔袋和内盒数量不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_R(4) <> strPart_B(4) Then
            .Row = 5
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_R(5) <> strPart_B(5) Then
            .Row = 6
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If strPart_R(6) <> strPart_B(6) Then
            .Row = 7
            .BackColor = vbRed
            MsgBox "标签不一致", vbInformation, "提示"
            Exit Sub

        End If

        If CLng(strPart_R(3)) = CLng(strPart_C(3)) Then
            InitCheckStatus
            Play ("该外箱已全部比对完成,请继续比对其他外箱")
        Else
            chk(1).Value = 0
            chk(2).Value = 0
            Play ("铝箔袋已比对完,数量不足, 请继续比对下一个内盒")

        End If

    End With

End If

End Sub

Private Sub CheckData()
Dim i         As Integer
Dim j         As Integer
Dim strcarton As String
Dim strBox    As String

On Error GoTo ErrHandle

Cnn.BeginTrans

With Fps(0)

    For i = 1 To .MaxRows
        For j = 0 To UBound(gNoCheckRow)
            If i = gNoCheckRow(j) Then
                GoTo NextRow

            End If

        Next
        If i = gCntRow Then
            .Row = i
            .Col = 1
            strcarton = Trim$(.text)
            .Col = 2
            strBox = Trim$(.text)
            If CheckIsEnough(i) = False Then
                GoTo ErrHandle
            Else
                GoTo NextRow

            End If

        End If

        For j = 0 To UBound(gUniqueRow) - 1
            If i = gUniqueRow(j) Then
                If CheckIsUnique(i) = False Then
                    GoTo ErrHandle
                Else
                    GoTo NextRow

                End If

            End If

        Next
        If CheckIsSame(i) = False Then
            GoTo ErrHandle

        End If

NextRow:
    Next

End With

' 判断状态
If CLng(strBox) < CLng(strcarton) Then
    Cnn.CommitTrans
    Play ("该内箱已核对完成, 请继续下一个内箱")
    '        MsgBox "该内箱已核对完成, 请继续下一个内箱", vbInformation, "提示"
    NextCheckStatus
Else
    Cnn.CommitTrans
    Play ("该外箱已核对完成, 请继续下一个外箱")
    '    MsgBox "该外箱已核对完成, 请继续下一个外箱", vbInformation, "提示"
    InitCheckStatus

End If

Exit Sub
ErrHandle:
Cnn.RollbackTrans

End Sub

Private Sub CheckData2()
Dim i         As Integer
Dim j         As Integer
Dim strcarton As String
Dim strBox    As String

On Error GoTo ErrHandle

Cnn.BeginTrans

With Fps(0)

    For i = 1 To .MaxRows
        For j = 0 To UBound(gNoCheckRow) - 1
            If i = gNoCheckRow(j) Then
                GoTo NextRow

            End If

        Next

        For j = 0 To UBound(gUniqueRow) - 1
            If i = gUniqueRow(j) Then
                If CheckIsUnique(i) = False Then
                    GoTo ErrHandle
                Else
                    GoTo NextRow

                End If

            End If

        Next
        If CheckIsSame2(i) = False Then
            GoTo ErrHandle

        End If

NextRow:
    Next

End With

' 判断状态
Cnn.CommitTrans
Play ("该外箱已核对完成, 请继续下一个外箱")
'    MsgBox "该外箱已核对完成, 请继续下一个外箱", vbInformation, "提示"
InitCheckStatus
Exit Sub
ErrHandle:
Cnn.RollbackTrans

End Sub

Private Sub CheckData_HK037()
Dim i         As Integer
Dim j         As Integer
Dim strcarton As String
Dim strBox    As String

With Fps(0)

    For i = 1 To .MaxRows
        If i = 9 Then
            .Col = 1
            .Row = i
            strcarton = Trim$(.text)
            .Col = 2
            .Row = i
            strBox = Trim$(.text)
            If strcarton = strBox Then
                .Row = i
                .Col = 1
                .BackColor = vbRed
                .Row = i
                .Col = 2
                .BackColor = vbRed
                Play ("铝箔袋与卷盘标签唯一码重复")
                MsgBox "铝箔袋与卷盘标签唯一码重复,请确认是否标签异常", vbInformation, "警告"
                Exit Sub

            End If

        Else
            .Col = 1
            .Row = i
            strcarton = Trim$(.text)
            .Col = 2
            .Row = i
            strBox = Trim$(.text)
            If strcarton <> strBox Then
                .Row = i
                .Col = 1
                .BackColor = vbRed
                .Row = i
                .Col = 2
                .BackColor = vbRed
                MsgBox "铝箔袋与卷盘其他信息不一致,请确认是否标签异常", vbInformation, "警告"
                Exit Sub

            End If

        End If

    Next

End With

chk(0).Value = 0
chk(1).Value = 0
chk(2).Value = 0
chk(3).Value = 0
Fps(0).MaxRows = 0
txtScan.Visible = True
txtScan.SetFocus
Play ("rightCode")

End Sub

Private Function CheckIsSame(irow As Integer) As Boolean
Dim strcarton As String
Dim strBox    As String
Dim strReel   As String

CheckIsSame = False

With Fps(0)
    .Col = 1
    .Row = irow
    strcarton = Trim$(.text)
    .Col = 2
    .Row = irow
    strBox = Trim$(.text)
    .Col = 3
    .Row = irow
    strReel = Trim$(.text)
    If strcarton <> strBox Then
        .Row = irow
        .Col = 1
        .BackColor = vbRed
        .Row = irow
        .Col = 2
        .BackColor = vbRed
        Play ("标签不一致")
        MsgBox "外箱和内箱标签不一致,请确认是否标签异常", vbInformation, "警告"
        Exit Function

    End If

    If strReel <> strBox Then
        .Row = irow
        .Col = 2
        .BackColor = vbRed
        .Row = irow
        .Col = 3
        .BackColor = vbRed
        Play ("标签不一致")
        MsgBox "内箱和铝箔袋标签不一致,请确认是否标签异常", vbInformation, "警告"
        Exit Function

    End If

End With

CheckIsSame = True

End Function

Private Function CheckIsSame2(irow As Integer) As Boolean
Dim strcarton As String
Dim strBox    As String
Dim strReel   As String

CheckIsSame2 = False

With Fps(0)
    .Col = 1
    .Row = irow
    strcarton = Trim$(.text)
    .Col = 2
    .Row = irow
    strBox = Trim$(.text)
    .Col = 3
    .Row = irow
    strReel = Trim$(.text)
    If strcarton <> strBox Then
        If InStr(strBox, "/") > 0 Then
            If Split(strBox, "/")(0) <> strcarton And Split(strBox, "/")(1) <> strcarton Then
                .Row = irow
                .Col = 1
                .BackColor = vbRed
                .Row = irow
                .Col = 2
                .BackColor = vbRed
                Play ("标签不一致")
                MsgBox "外箱和内箱标签不一致,请确认是否标签异常", vbInformation, "警告"
                Exit Function

            End If

        End If

    End If

    If strReel <> strBox Then
        .Row = irow
        .Col = 2
        .BackColor = vbRed
        .Row = irow
        .Col = 3
        .BackColor = vbRed
        Play ("标签不一致")
        MsgBox "内箱和铝箔袋标签不一致,请确认是否标签异常", vbInformation, "警告"
        Exit Function

    End If

End With

CheckIsSame2 = True

End Function

Private Function CheckIsEnough(irow As Integer) As Boolean
Dim strcarton As String
Dim strBox    As String
Dim strReel   As String

CheckIsEnough = False

With Fps(0)
    .Col = 1
    .Row = irow
    strcarton = Trim$(.text)
    .Col = 2
    .Row = irow
    strBox = Trim$(.text)
    .Col = 3
    .Row = irow
    strReel = Trim$(.text)
    If CLng(strBox) > CLng(strcarton) Then
        .Row = irow
        .Col = 1
        .BackColor = vbRed
        .Row = irow
        .Col = 2
        .BackColor = vbRed
        Play ("标签数量错误")
        MsgBox "内箱数量大于外箱数量,请确认是否标签异常", vbInformation, "警告"
        Exit Function

    End If

    If CLng(strReel) <> CLng(strBox) Then
        .Row = irow
        .Col = 2
        .BackColor = vbRed
        .Row = irow
        .Col = 3
        .BackColor = vbRed
        Play ("标签数量错误")
        MsgBox "内箱和铝箔袋标签数量不一致,请确认是否标签异常", vbInformation, "警告"
        Exit Function

    End If

End With

CheckIsEnough = True

End Function

Private Sub Play(sFileName As String)
Dim sPath   As String
Dim sSuffix As String

sPath = "\\10.160.1.84\public\media_source\"
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix

End Sub

Private Function CheckIsUnique(irow As Integer) As Boolean
Dim strSql  As String
Dim strCode As String
Dim i       As Integer

CheckIsUnique = False

With Fps(0)

    For i = 1 To 3
        .Row = irow
        .Col = i
        If .Col = 1 And gStatus = E_CheckStatus.E_CARTON_CHECKED Then
            GoTo NEXTCOL

        End If

        Select Case i

            Case 1
                strCode = Trim$(.text) & "000004"

            Case 2
                strCode = Trim$(.text) & "000003"

            Case 3
                strCode = Trim$(.text) & "000002"

        End Select

        strSql = "select * from UNIQUE_TBL where key_value = '" & strCode & "'"
        If Get_OracleCnt(strSql) > 0 Then
            .Row = irow
            .Col = i
            .BackColor = vbRed
            Play ("标签唯一码重复")
            MsgBox "存在相同的唯一码, 请确认是否有误", vbInformation, "警告"
            Exit Function

        End If

        If i = 3 And InStr(.text, "-R") Then
            Dim strTHis As String, strLast As String

            strTHis = Replace(Replace(Trim$(.text), "-R", ""), "-B", "")
            .Col = 2
            .Row = irow
            strLast = Replace(Replace(Trim$(.text), "-R", ""), "-B", "")
            If strTHis <> strLast Then
                .Row = irow
                .Col = 3
                .BackColor = vbRed
                .Row = irow
                .Col = 2
                .BackColor = vbRed
                MsgBox "铝箔袋和内箱ID不能对应", vbInformation, "警告"
                Exit Function

            End If

        End If

        AddSql ("insert into unique_tbl(KEY_ID, KEY_VALUE,UPDATE_TIME,UPDATE_BY) values('" & gID & "', '" & strCode & "', sysdate, '" & gUserName & "') ")
NEXTCOL:
    Next

End With

CheckIsUnique = True

End Function

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       verifyLbl_GC
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       DESKTOP-91AFCV3
' Date-Time  :       2019/4/4-15:50:14
'
' Parameters :       strCode (String)
'--------------------------------------------------------------------------------
Private Sub verifyLbl_GC(strCode As String)
Dim strArray()  As String
Dim strArray2() As String
Dim bExisted    As Boolean
Dim i           As Integer
Dim lQty        As Long

bExisted = False
If txtPackingNO.text = "" Then
    strArray = Split(strCode, ",")
    If UBound(strArray) <> 13 Then
        MsgBox "请扫描正确的GC外箱二维码,或外箱标签模板已变更", vbExclamation + vbOKOnly, "错误"
        Exit Sub

    End If

    If Get_OracleCnt("select * from unique_tbl_new where KEYFROM = 'GC' and KEYNAME = '外箱' and KEYVALUE= '" & strArray(12) & "' ") > 0 Then
        MsgBox "该外箱号:" & strArray(12) & vbCrLf & "之前已经核对过, 请确认本次是否是重复异常标签", vbExclamation + vbOKOnly, "错误"
        ExporToExcel ("select KEYFROM as 客户, KEYNAME as 外箱, KEYVALUE as 箱号, KEYTIME as 核对日期, KEYBY as 核对人员  from UNIQUE_TBL_NEW where KEYFROM = 'GC' and KEYNAME = '外箱' and KEYVALUE= '" & strArray(12) & "'  order by KEYTIME desc")
        Exit Sub

    End If

    txtPackingNO.text = strArray(12)
    txtPackingQty.text = strArray(6)
    Play ("外箱已扫描,请依次扫描铝箔袋标签")
    ReDim strLblInfo(strArray(4))

    For i = 0 To UBound(strLblInfo) - 1
        strArray2 = Split(strArray(3), " ")
        strLblInfo(i).strWaferID = strArray(2) & Right("0" & strArray2(i), 2)
        strLblInfo(i).strCodePP = strArray(5)
        strLblInfo(i).strSecCode = strArray(11)
        strLblInfo(i).strCusDev = Split(strArray(0), "/")(1)
        strLblInfo(i).bChecked = False
    Next
Else
    strArray = Split(strCode, ",")
    If UBound(strArray) <> 8 Then
        MsgBox "请扫描正确的铝箔袋二维码,或外箱标签模板已变更", vbExclamation + vbOKOnly, "错误"
        Exit Sub

    End If

    For i = 0 To UBound(strLblInfo) - 1
        If (strLblInfo(i).strWaferID = Replace(strArray(0), "-", "")) Then
            bExisted = True
            If strLblInfo(i).strSecCode <> strArray(3) Then
                Play ("铝箔袋标签不正确")
                MsgBox "铝箔袋二级代码: " & strArray(3) & "  不正确", vbCritical, "警告"
                Exit Sub

            End If

            If strLblInfo(i).strCodePP <> strArray(2) Then
                Play ("铝箔袋标签不正确")
                MsgBox "锡球代码: " & strArray(2) & "  不正确", vbCritical, "警告"
                Exit Sub

            End If

            If strLblInfo(i).strCusDev <> Replace(strArray(1), "-3", "") Then
                MsgBox "机种错误:" & strArray(1), vbCritical, "警告"
                Exit Sub

            End If

            If strLblInfo(i).bChecked = True Then
                Play ("请确认是否重复扫描或标签出错")
                MsgBox "该卷盘:" & strLblInfo(i).strWaferID & "  已经核对过" & vbCrLf & "请确认是否重复扫描或标签出错", vbCritical, "警告"
                Exit Sub

            End If

            lQty = CLng(txtPackingQtyAdd.text) + strArray(5)
            If lQty = CLng(txtPackingQty.text) Then
                Play ("该外箱已全部核对正确")
                AddSql ("insert into UNIQUE_TBL_NEW(KEYNAME, KEYVALUE, KEYFROM,KEYTIME,KEYBY) values('外箱','" & txtPackingNO.text & "','GC',sysdate, '" & gUserName & "')")
                MsgBox "全部核对完成", vbInformation, "提示"
                clearLbl_GC
                Exit Sub
            ElseIf lQty > CLng(txtPackingQty.text) Then
                Play ("数量超出,铝箔袋和外箱数量不相等")
                MsgBox "数量超出,铝箔袋和外箱数量不相等", vbCritical, "警告"
                Exit Sub

            End If

            Play ("铝箔袋正确")
            txtPackingQtyAdd.text = lQty
            strLblInfo(i).bChecked = True

        End If

    Next
    If bExisted = False Then
        Play ("铝箔袋标签不正确")
        MsgBox "铝箔袋WaferID: " & Replace(strArray(0), "-", "") & "  不正确", vbCritical, "警告"
        Exit Sub

    End If

End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       ListData_HD
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       DESKTOP-MSUG5JD
' Date-Time  :       2019/10/29-8:51:30
'
' Parameters :       strCode (String)
'--------------------------------------------------------------------------------
Private Sub ListData_HD(strCode As String)
Dim strPSN       As String
Dim strSameItems As String
Dim i            As Integer
Dim j            As Integer
Dim strArray()   As String

strSameItems = "1,3,4,9,10,11,12"
If chk(0).Value = 0 Then ' 1.外箱(C)
    strBoxID = ""
    strPart_C = Split(strCode, "/")
    If UBound(strPart_C) <> 11 Then
        MsgBox "外箱二维码不正确", vbInformation, "提示"
        Exit Sub

    End If

    With Fps(0)
        .SetText 1, 1, strPart_C(0)
        .SetText 1, 2, strPart_C(1)
        .SetText 1, 3, strPart_C(2)
        .SetText 1, 4, strPart_C(3)
        .SetText 1, 5, strPart_C(4)
        .SetText 1, 6, strPart_C(5)
        .SetText 1, 7, strPart_C(6)
        .SetText 1, 8, strPart_C(7)
        .SetText 1, 9, strPart_C(8)
        .SetText 1, 10, strPart_C(9)
        .SetText 1, 11, strPart_C(10)
        .SetText 1, 12, strPart_C(11)

    End With

    '0.第7项必须为0
    If strPart_C(6) <> "0" Then
        Fps(0).Col = 1
        Fps(0).Row = 7
        Fps(0).BackColor = vbRed
        MsgBox "标签第7项必须为0", vbCritical, "错误"
        Exit Sub

    End If

    '1.检查标签唯一码
    strPSN = strPart_C(1)
    If InStr(strBoxID, strPSN) > 0 Then
        MsgBox "外箱唯一码:" & strPSN & "已扫描,不可重复", vbCritical, "警告"
        Fps(0).Col = 1
        Fps(0).Row = 2
        Fps(0).BackColor = vbRed
        Exit Sub

    End If

    If Get_OracleCnt("select * from UNIQUE_TBL_NEW where KEYFROM = 'HD' and KEYVALUE = '" & strPSN & "' ") Then
        MsgBox "外箱唯一码:" & strPSN & "已存在,不可重复", vbCritical, "警告"
        Fps(0).Col = 1
        Fps(0).Row = 2
        Fps(0).BackColor = vbRed
        Exit Sub

    End If

    '2.检查标签唯一码第5位标志位
    If Mid$(strPSN, 6, 1) <> "C" Then
        MsgBox "外箱标签唯一码:" & strPSN & " 第六位必须为C", vbCritical, "警告"
        Fps(0).Col = 1
        Fps(0).Row = 2
        Fps(0).BackColor = vbRed
        Exit Sub

    End If

    '3.数量检查
    If strPart_C(5) > 60000 Then
        Fps(0).Col = 1
        Fps(0).Row = 6
        Fps(0).BackColor = vbRed
        MsgBox "外箱最大数量不可大于60000", vbCritical, "警告"
        Exit Sub

    End If

    '4.批号子串查询
    For i = 1 To UBound(Split(strPart_C(4), "|"))
        If Left(Split(strPart_C(4), "|")(i), 2) <> "0)" Then
            Fps(0).Col = 1
            Fps(0).Row = 5
            Fps(0).BackColor = vbRed
            MsgBox "外箱批号:" & Split(strPart_C(4), "|")(i - 1) & "的不良品数量" & Left(Split(strPart_C(4), "|")(i), 2) & ",不为0", vbCritical, "警告"
            Exit Sub

        End If

    Next
    Dim strSumWX As Long

    For i = 0 To UBound(Split(strPart_C(4), "|")) - 1
        strSumWX = strSumWX + CLng(Split(Split(strPart_C(4), "|")(i), "(")(1))
    Next
    If strPart_C(5) <> strSumWX Then
        Fps(0).Col = 1
        Fps(0).Row = 5
        Fps(0).BackColor = vbRed
        MsgBox "外箱批号数量总和:" & strSumWX & " 不等于实际外箱数量:" & strPart_C(5), vbCritical, "警告"
        Exit Sub

    End If

    If Len(strPart_C(4)) - Len(Replace(strPart_C(4), "|", "")) > 8 Then
        Fps(0).Col = 1
        Fps(0).Row = 5
        Fps(0).BackColor = vbRed
        MsgBox "外箱包含批号不可大于8", vbCritical, "警告"
        Exit Sub

    End If

    With Fps(0)
        .Col = 1

        For j = 1 To .MaxRows
            .Row = j
            If .BackColor = vbRed Then
                .BackColor = vbWhite

            End If

        Next

    End With

    Play ("外箱标签已扫描,请扫描内箱标签")
    strBoxID = strBoxID & strPSN & ","
    chk(0).Value = 1
ElseIf chk(1).Value = 0 Then    ' 2.内箱(B)
    strPart_B = Split(strCode, "/")
    If UBound(strPart_B) <> 11 Then
        MsgBox "内箱二维码不正确", vbInformation, "提示"
        Exit Sub

    End If

    With Fps(0)
        .SetText 2, 1, strPart_B(0)
        .SetText 2, 2, strPart_B(1)
        .SetText 2, 3, strPart_B(2)
        .SetText 2, 4, strPart_B(3)
        .SetText 2, 5, strPart_B(4)
        .SetText 2, 6, strPart_B(5)
        .SetText 2, 7, strPart_B(6)
        .SetText 2, 8, strPart_B(7)
        .SetText 2, 9, strPart_B(8)
        .SetText 2, 10, strPart_B(9)
        .SetText 2, 11, strPart_B(10)
        .SetText 2, 12, strPart_B(11)
        '0.第7项必须为0
        If strPart_B(6) <> "0" Then
            Fps(0).Col = 2
            Fps(0).Row = 7
            Fps(0).BackColor = vbRed
            MsgBox "标签第7项必须为0", vbCritical, "错误"
            ClearIb
            Exit Sub

        End If

        '1.检查标签唯一码
        strPSN = strPart_B(1)
        If InStr(strBoxID, strPSN) > 0 Then
            MsgBox "内箱唯一码:" & strPSN & "已扫描,不可重复", vbCritical, "警告"
            Fps(0).Col = 2
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            Exit Sub

        End If

        If Get_OracleCnt("select * from UNIQUE_TBL_NEW where KEYFROM = 'HD' and KEYVALUE = '" & strPSN & "' ") Then
            Fps(0).Col = 2
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            MsgBox "内箱唯一码:" & strPSN & "已存在,不可重复", vbCritical, "警告"
            ClearIb
            Exit Sub

        End If

        '2.检查标签唯一码第5位标志位
        If Mid$(strPSN, 6, 1) <> "B" Then
            Fps(0).Col = 2
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            MsgBox "内箱标签唯一码:" & strPSN & " 第六位必须为B", vbCritical, "警告"
            ClearIb
            Exit Sub

        End If

        '3.相同项检查
        strArray = Split(strSameItems, ",")

        For i = 0 To UBound(strArray)
            If strPart_B(strArray(i) - 1) <> strPart_C(strArray(i) - 1) Then
                Fps(0).Col = 1
                Fps(0).Row = strArray(i)
                Fps(0).BackColor = vbRed
                Fps(0).Col = 2
                Fps(0).Row = strArray(i)
                Fps(0).BackColor = vbRed
                MsgBox "标签不一致", vbInformation, "提示"
                ClearIb
                Exit Sub

            End If

        Next
        '4.特殊项检查
        If (Left(strPart_B(1), 5) <> Left(strPart_C(1), 5)) Or (Mid$(strPart_B(1), 7, 2) <> Mid$(strPart_C(1), 7, 2)) Then
            Fps(0).Col = 1
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            Fps(0).Col = 2
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            MsgBox "标签唯一码:" & strPSN & " 特殊位不一致", vbCritical, "警告"
            ClearIb
            Exit Sub

        End If

        '5.数量检查
        If strPart_B(5) > 15000 Then
            Fps(0).Col = 2
            Fps(0).Row = 6
            Fps(0).BackColor = vbRed
            MsgBox "内箱最大数量不可大于15000", vbCritical, "警告"
            ClearIb
            Exit Sub

        End If

        '6.日期检查
        If Abs(DateDiff("d", strPart_B(7), strPart_C(7))) > 30 Then
            Fps(0).Col = 1
            Fps(0).Row = 8
            Fps(0).BackColor = vbRed
            Fps(0).Col = 2
            Fps(0).Row = 8
            Fps(0).BackColor = vbRed
            MsgBox "内箱标签和外箱标签日期间隔不可超过三十天", vbCritical, "日期错误"
            ClearIb
            Exit Sub

        End If

        '7.批号子串查询
        For i = 1 To UBound(Split(strPart_B(4), "|"))
            If Left(Split(strPart_B(4), "|")(i), 2) <> "0)" Then
                Fps(0).Col = 2
                Fps(0).Row = 5
                Fps(0).BackColor = vbRed
                MsgBox "内箱批号:" & Split(strPart_B(4), "|")(i - 1) & "的不良品数量" & Left(Split(strPart_B(4), "|")(i), 2) & ",不为0", vbCritical, "警告"
                Exit Sub

            End If

        Next
        Dim strSumNH As Long

        For i = 0 To UBound(Split(strPart_B(4), "|")) - 1
            strSumNH = strSumNH + CLng(Split(Split(strPart_B(4), "|")(i), "(")(1))
        Next
        If strPart_B(5) <> strSumNH Then
            Fps(0).Col = 2
            Fps(0).Row = 5
            Fps(0).BackColor = vbRed
            MsgBox "内箱批号数量总和:" & strSumNH & " 不等于实际内箱数量:" & strPart_B(5), vbCritical, "警告"
            Exit Sub

        End If

        If Len(strPart_B(4)) - Len(Replace(strPart_B(4), "|", "")) > 5 Then
            Fps(0).Col = 2
            Fps(0).Row = 5
            Fps(0).BackColor = vbRed
            MsgBox "内箱包含批号不可大于5", vbCritical, "警告"
            Exit Sub

        End If

        Dim strArrNH() As String

        strArrNH = Split(strPart_B(4), "|0)")
        Fps(0).Col = 1
        Fps(0).Row = 5

        For i = 0 To UBound(strArrNH) - 1
            If InStr(.text, strArrNH(i)) = 0 Then
                MsgBox "外箱批号:" & .text & "不包含该内箱批号:" & strArrNH(i)
                Fps(0).Col = 2
                Fps(0).Row = 5
                Fps(0).BackColor = vbRed
                Exit Sub

            End If

        Next
        chk(1).Value = 1

        With Fps(0)
            .Col = 2

            For j = 1 To .MaxRows
                .Row = j
                If .BackColor = vbRed Then
                    .BackColor = vbWhite

                End If

            Next
            .Col = 1

            For j = 1 To .MaxRows
                .Row = j
                If .BackColor = vbRed Then
                    .BackColor = vbWhite

                End If

            Next

        End With

        gIBCntSum = gIBCntSum + strPart_B(5)
        txtIbCnt.text = gIBCntSum
        Play ("内箱标签已扫描,请扫描铝箔袋标签")
        strBoxID = strBoxID & strPSN & ","

    End With

ElseIf chk(2).Value = 0 Then    ' 3.铝箔袋(R)
    strPart_R = Split(strCode, "/")
    If UBound(strPart_R) <> 11 Then
        MsgBox "内箱二维码不正确", vbInformation, "提示"
        Exit Sub

    End If

    With Fps(0)
        .SetText 3, 1, strPart_R(0)
        .SetText 3, 2, strPart_R(1)
        .SetText 3, 3, strPart_R(2)
        .SetText 3, 4, strPart_R(3)
        .SetText 3, 5, strPart_R(4)
        .SetText 3, 6, strPart_R(5)
        .SetText 3, 7, strPart_R(6)
        .SetText 3, 8, strPart_R(7)
        .SetText 3, 9, strPart_R(8)
        .SetText 3, 10, strPart_R(9)
        .SetText 3, 11, strPart_R(10)
        .SetText 3, 12, strPart_R(11)
        '0.第7项必须为0
        If strPart_R(6) <> "0" Then
            Fps(0).Col = 3
            Fps(0).Row = 7
            Fps(0).BackColor = vbRed
            MsgBox "标签第7项必须为0", vbCritical, "错误"
            ClearLv
            Exit Sub

        End If

        '1.检查标签唯一码
        strPSN = strPart_R(1)
        If InStr(strBoxID, strPSN) > 0 Then
            MsgBox "铝箔袋唯一码:" & strPSN & "已扫描,不可重复", vbCritical, "警告"
            Fps(0).Col = 3
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            Exit Sub

        End If

        If Get_OracleCnt("select * from UNIQUE_TBL_NEW where KEYFROM = 'HD' and KEYVALUE = '" & strPSN & "' ") Then
            MsgBox "铝箔袋唯一码:" & strPSN & "已存在,不可重复", vbCritical, "警告"
            ClearLv
            Exit Sub

        End If

        '2.检查标签唯一码第5位标志位
        If Mid$(strPSN, 6, 1) <> "A" Then
            MsgBox "铝箔袋标签唯一码:" & strPSN & " 第六位必须为A", vbCritical, "警告"
            ClearLv
            Exit Sub

        End If

        '3.数量检查
        If strPart_R(5) > 3000 Then
            Fps(0).Col = 3
            Fps(0).Row = 6
            Fps(0).BackColor = vbRed
            MsgBox "铝箔袋最大数量不可大于3000", vbCritical, "警告"
            ClearLv
            Exit Sub

        End If

        '4.相同项检查
        strArray = Split(strSameItems, ",")

        For i = 0 To UBound(strArray)
            If strPart_R(strArray(i) - 1) <> strPart_B(strArray(i) - 1) Then
                .Col = 2
                .Row = strArray(i)
                .BackColor = vbRed
                .Col = 3
                .Row = strArray(i)
                .BackColor = vbRed
                MsgBox "标签不一致", vbInformation, "提示"
                ClearLv
                Exit Sub

            End If

        Next
        '5.日期检查
        If Abs(DateDiff("d", strPart_R(7), strPart_B(7))) > 30 Then
            Fps(0).Col = 3
            Fps(0).Row = 8
            Fps(0).BackColor = vbRed
            Fps(0).Col = 2
            Fps(0).Row = 8
            Fps(0).BackColor = vbRed
            MsgBox "铝箔袋标签和内盒标签日期间隔不可超过三十天", vbCritical, "日期错误"
            ClearLv
            Exit Sub

        End If

        If Abs(DateDiff("d", strPart_R(7), strPart_C(7))) > 30 Then
            Fps(0).Col = 3
            Fps(0).Row = 8
            Fps(0).BackColor = vbRed
            Fps(0).Col = 1
            Fps(0).Row = 8
            Fps(0).BackColor = vbRed
            MsgBox "铝箔袋标签和外箱标签日期间隔不可超过三十天", vbCritical, "日期错误"
            ClearLv
            Exit Sub

        End If

        '6.批号子串查询
        For i = 1 To UBound(Split(strPart_R(4), "|"))
            If Left(Split(strPart_R(4), "|")(i), 2) <> "0)" Then
                Fps(0).Col = 3
                Fps(0).Row = 5
                Fps(0).BackColor = vbRed
                MsgBox "铝箔袋批号:" & Split(strPart_R(4), "|")(i - 1) & "的不良品数量" & Left(Split(strPart_R(4), "|")(i), 2) & ",不为0", vbCritical, "警告"
                Exit Sub

            End If

        Next
        Dim strSumLV As Long

        For i = 0 To UBound(Split(strPart_R(4), "|")) - 1
            strSumLV = strSumLV + CLng(Split(Split(strPart_R(4), "|")(i), "(")(1))
        Next
        If strPart_R(5) <> strSumLV Then
            Fps(0).Col = 3
            Fps(0).Row = 5
            Fps(0).BackColor = vbRed
            MsgBox "铝箔袋批号数量总和:" & strSumLV & " 不等于实际铝箔袋数量:" & strPart_R(5), vbCritical, "警告"
            Exit Sub

        End If

        If Len(strPart_R(4)) - Len(Replace(strPart_R(4), "|", "")) > 2 Then
            Fps(0).Col = 3
            Fps(0).Row = 5
            Fps(0).BackColor = vbRed
            MsgBox "铝箔袋包含批号不可大于2", vbCritical, "警告"
            Exit Sub

        End If

        Dim strArrLV() As String

        strArrLV = Split(strPart_R(4), "|0)")
        Fps(0).Col = 2
        Fps(0).Row = 5

        For i = 0 To UBound(strArrLV) - 1
            If InStr(.text, strArrLV(i)) = 0 Then
                MsgBox "内箱批号:" & .text & "不包含该铝箔袋批号:" & strArrLV(i)
                Fps(0).Col = 3
                Fps(0).Row = 5
                Fps(0).BackColor = vbRed
                Exit Sub

            End If

        Next
        '7.数量累计
        gLVCntSum = gLVCntSum + CLng(strPart_R(5))
        txtLvCnt.text = gLVCntSum
        If gLVCntSum = CLng(strPart_B(5)) Then
            strBoxID = strBoxID & strPSN & ","
            chk(1).Value = 0
            chk(2).Value = 0
            gLVCntSum = 0
            txtLvCnt.text = gLVCntSum
            Play ("该内盒已经核对完成,请继续下一个内盒")
            ClearLv
            ClearIb
            If gIBCntSum = CLng(strPart_C(5)) Then
                Call InitCheckStatus
                Call SaveBoxID
                Play ("该外箱已全部比对完成,请继续比对其他外箱")
            ElseIf gIBCntSum > CLng(strPart_C(5)) Then
                MsgBox "内盒数量累计总和" & gIBCntSum & "不可大于外箱数量:" & strPart_C(5), vbCritical, "数量错误"

            End If

        ElseIf gLVCntSum > CLng(strPart_B(5)) Then
            MsgBox "铝箔袋数量累计总和:" & gLVCntSum & "不可大于内箱数量:" & strPart_B(5), vbCritical, "数量错误"
            ClearLv
            Exit Sub
        Else
            Play ("该铝箔袋已扫描请继续扫描下一个铝箔袋")
            strBoxID = strBoxID & strPSN & ","

            With Fps(0)
                .Col = 3

                For j = 1 To .MaxRows
                    .Row = j
                    .text = ""
                Next

            End With

            chk(2).Value = 0

            With Fps(0)
                .Col = 3

                For j = 1 To .MaxRows
                    .Row = j
                    If .BackColor = vbRed Then
                        .BackColor = vbWhite

                    End If

                Next
                .Col = 2

                For j = 1 To .MaxRows
                    .Row = j
                    If .BackColor = vbRed Then
                        .BackColor = vbWhite

                    End If

                Next
                .Col = 1

                For j = 1 To .MaxRows
                    .Row = j
                    If .BackColor = vbRed Then
                        .BackColor = vbWhite

                    End If

                Next

            End With

        End If

    End With

End If

End Sub

Private Sub SaveBoxID()
Dim i               As Integer
Dim strSql          As String
Dim strBoxIDArray() As String

strBoxIDArray = Split(strBoxID, ",")

For i = 0 To UBound(strBoxIDArray) - 1
    strSql = "insert into UNIQUE_TBL_NEW(KEYNAME,KEYVALUE,KEYFROM,KEYTIME,KEYBY) values('箱号唯一码','" & strBoxIDArray(i) & "','HD',sysdate,'" & gUserName & "') "
    AddSql (strSql)
Next

End Sub

Private Sub ClearLv()
Dim i As Integer

With Fps(0)
    .Col = 3

    For i = 1 To .MaxRows
        .Row = i
        .text = ""
        .BackColor = vbWhite
    Next

End With

End Sub

Private Sub ClearIb()
Dim i As Integer

With Fps(0)
    .Col = 2

    For i = 1 To .MaxRows
        .Row = i
        .text = ""
        .BackColor = vbWhite
    Next

End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       正式工程1
' Procedure  :       clearLbl_GC
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       DESKTOP-91AFCV3
' Date-Time  :       2019/4/8-10:46:54
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub clearLbl_GC()
txtPackingNO.text = ""
txtPackingQty.text = ""
txtPackingQtyAdd.text = 0
Erase strLblInfo

End Sub

Private Sub ListData_US026(strCode As String)


If chk(0).Value = 0 Then ' 1.外箱(C)
    strPart_C = Split(strCode, ",")
    If UBound(strPart_C) <> 7 Then
        MsgBox "外箱二维码不正确", vbInformation, "提示"
        Exit Sub

    End If

    With Fps(0)
        .MaxRows = .MaxRows + 1
        .SetText 1, .MaxRows, strPart_C(1)
        .SetText 2, .MaxRows, strPart_C(2)
        .SetText 3, .MaxRows, strPart_C(3)
        .SetText 4, .MaxRows, strPart_C(5)
        .SetText 5, .MaxRows, strPart_C(6)
        .SetText 6, .MaxRows, Left(strPart_C(7), 8)

    End With

    lWXQty = CLng(strPart_C(5))
    Play ("外箱标签已扫描,请扫描内箱标签")
    chk(0).Value = 1
ElseIf chk(1).Value = 0 Then    ' 2.内箱(B)
    strPart_B = Split(strCode, ",")
    If UBound(strPart_B) <> 7 Then
        MsgBox "内箱二维码不正确", vbInformation, "提示"
        Exit Sub

    End If

    With Fps(0)
        .MaxRows = .MaxRows + 1
        .SetText 1, .MaxRows, strPart_B(0)
        .SetText 2, .MaxRows, strPart_B(1)
        .SetText 3, .MaxRows, strPart_B(5)
        .SetText 4, .MaxRows, strPart_B(2)
        .SetText 5, .MaxRows, strPart_B(3)
        .SetText 6, .MaxRows, Left(strPart_B(4), 8)
        '1.Device
        If strPart_B(0) <> strPart_C(1) Then
            Fps(0).Col = 1
            Fps(0).Row = Fps(0).MaxRows
            Fps(0).BackColor = vbRed
            MsgBox "客户机种不一致", vbCritical, "错误"
            Fps(0).DeleteRows .MaxRows, 1
            Fps(0).MaxRows = .MaxRows - 1
            Exit Sub

        End If

        '2.LotID
        If strPart_C(2) <> strPart_B(1) Then
            .Col = 2
            .Row = .MaxRows
            .BackColor = vbRed
            MsgBox "LotID不一致", vbCritical, "错误"
            Fps(0).DeleteRows .MaxRows, 1
            Fps(0).MaxRows = .MaxRows - 1
            Exit Sub

        End If

    
        '3.Date code
        If strPart_B(3) <> strPart_C(6) Then
            Fps(0).Col = 5
            Fps(0).Row = Fps(0).MaxRows
            Fps(0).BackColor = vbRed
            MsgBox "DateCode不一致", vbCritical, "错误"
            Fps(0).DeleteRows .MaxRows, 1
            Fps(0).MaxRows = .MaxRows - 1
            Exit Sub

        End If

        '4.HT Lot
        If Left(strPart_B(4), 8) <> Left(strPart_C(7), 8) Then
            Fps(0).Col = 6
            Fps(0).Row = Fps(0).MaxRows
            Fps(0).BackColor = vbRed
            MsgBox "厂内批号不一致", vbCritical, "错误"
            Fps(0).DeleteRows .MaxRows, 1
            Fps(0).MaxRows = .MaxRows - 1
            Exit Sub

        End If
        
        '2.Wafer ID
        If InStr(strPart_C(3), strPart_B(5)) = 0 Then
            Fps(0).Col = 3
            Fps(0).Row = Fps(0).MaxRows
            Fps(0).BackColor = vbRed
            MsgBox "WaferID不存在", vbCritical, "错误"
            Fps(0).DeleteRows Fps(0).MaxRows, 1
            Fps(0).MaxRows = .MaxRows - 1
            Exit Sub

        End If

        
        If Replace(Replace$(strPart_C(3), strPart_B(5), ""), " ", "") = "" Then
            If lNXQty + CLng(strPart_B(2)) = lWXQty Then
                
                Call InitCheckStatus
                Play ("该外箱已全部核对完成,请核对其他外箱")
            Else
        
                Fps(0).Col = 4
                Fps(0).Row = Fps(0).MaxRows
                Fps(0).BackColor = vbRed
                MsgBox "内外箱数量不对应", vbCritical, "错误"
                Fps(0).DeleteRows .MaxRows, 1
                Fps(0).MaxRows = .MaxRows - 1
                Exit Sub

            End If

        Else
            Play ("该内盒已扫描,请扫描下个内盒")
            Fps(0).Row = 1
            Fps(0).Col = 3
            Fps(0).text = Replace$(strPart_C(3), strPart_B(5), "")
            strPart_C(3) = Replace$(strPart_C(3), strPart_B(5), "")
            
            lNXQty = lNXQty + CLng(strPart_B(2))

        End If

    End With

End If

txtWXQty.text = lWXQty
txtNXQty.text = lNXQty

End Sub
