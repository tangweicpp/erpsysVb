VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.ocx"
Begin VB.Form Frm_Label_Checking_System 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "ͨ�ñ�ǩ�˶�ϵͳ(GLCS)_��ά��"
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
         Caption         =   "����"
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
         Caption         =   "ɾ���˶���ż�¼"
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
         Caption         =   "�����˶Լ�¼"
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
         Caption         =   "�˳�"
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
         Caption         =   "������"
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
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "����"
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
         Caption         =   "��ʼɨ��"
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
         Caption         =   "�ۼ�����:"
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
         Caption         =   "��ǰ��������:"
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
         Caption         =   "��ǰ�������:"
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
         Caption         =   "���"
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
         Caption         =   "�˶�����"
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
'    Project    : ��ʽ����1
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

    Case "HK037_������_����"
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
            .SetText 1, 0, "������"
            .SetText 2, 0, "����"
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
        .SetText 1, 0, "��  ��(C)"
        .SetText 2, 0, "��  ��(B)"
        .SetText 3, 0, "������(R)"
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
    MsgBox "��ѡ��ģ��", vbInformation, "��ʾ"
    Exit Sub

End If

Select Case cmbCombo1.text

    Case "HK037_������_����"

        If InStr("07885", gUserName) > 0 Then
            
            strBoxID = UCase$(Trim$(txtBoxID.text))

            If Len(strBoxID) = 0 Then
                MsgBox "������Ҫɾ������ʷ���", vbInformation, "��ʾ"
                Exit Sub

            End If

            AddSql ("insert into unique_tbl_bak select * from unique_tbl where key_value = '" & strBoxID & "' ")
            AddSql ("delete from unique_tbl where key_value = '" & strBoxID & "' ")
            MsgBox "��ʷ��¼�Ѿ����", vbInformation, "��ʾ"
        Else
            MsgBox "��û��ɾ����Ȩ��,����ϵIT����ɾ��", vbInformation, "����"
            Exit Sub

        End If

    Case "DA24"
        
        strBoxID = UCase$(Trim$(txtBoxID.text))

        If Len(strBoxID) = 0 Then
            MsgBox "������Ҫɾ������ʷ���", vbInformation, "��ʾ"
            Exit Sub

        End If

        AddSql ("delete from unique_tbl_new where KEYFROM = 'DA24' and KEYNAME='PACKNO' and keyvalue = '" & strBoxID & "' ")
        MsgBox "��ʷ�����ɾ��", vbInformation, "��ʾ"
        Exit Sub

    Case "HD"
        strBoxID = UCase$(Trim$(txtBoxID.text))

        If Len(strBoxID) = 0 Then
            MsgBox "������Ҫɾ������ʷ���", vbInformation, "��ʾ"
            Exit Sub

        End If

        AddSql ("delete from UNIQUE_TBL_NEW where KEYFROM = 'HD' and keyvalue = '" & strBoxID & "' ")
        MsgBox "��ʷ���:" & strBoxID & " ��ɾ��", vbInformation, "��ʾ"
        Exit Sub

    Case "GC"
        MsgBox "��Ų���ɾ��", vbInformation, "��ʾ"

    Case "SH50"
        MsgBox "�ÿͻ�û����ſ���ɾ��", vbInformation, "��ʾ"

    Case Else

        If InStr("07885", gUserName) > 0 Then
            
            strBoxID = UCase$(Trim$(txtBoxID.text))

            If Len(strBoxID) = 0 Then
                MsgBox "������Ҫɾ������ʷ���", vbInformation, "��ʾ"
                Exit Sub

            End If

            AddSql ("insert into unique_tbl_bak select * from unique_tbl where key_value = '" & strBoxID & "' ")
            AddSql ("delete from unique_tbl where key_value = '" & strBoxID & "' ")
            MsgBox "��ʷ��¼�Ѿ����", vbInformation, "��ʾ"
        Else
            MsgBox "��û��ɾ����Ȩ��,����ϵIT����ɾ��", vbInformation, "����"
            Exit Sub

        End If
End Select

End Sub

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdStart_Click()
If cmbCombo1.text = "" Then
    MsgBox "��ѡ��˶�����", vbInformation, "��ʾ"
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

    Case "HK037_������_����"

    Case "DA24"
        'Play ("������ɨ������,����,�������Ķ�ά���ǩ")
        Fps(0).MaxRows = 7

    Case "GC"
        If txtPackingNO.text = "" Then
            Play ("��ɨ��GC�������ǩ��ά��")

        End If

    Case "SH50"
        Play ("������ɨ������,����,�������Ķ�ά���ǩ")
        Fps(0).MaxRows = 7

    Case "HD"
        Play ("��ɨ�������ǩ��ά��")
        Fps(0).MaxRows = 12

    Case "US026"
        Play ("��ɨ�������ǩ��ά��")
      
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
    MsgBox "��ѡ������", vbInformation, "��ʾ"
    Exit Sub

End If

Select Case cmbCombo1.text

    Case "HK037_������_����"
        ExporToExcel ("select * from unique_tbl order by update_time desc")

    Case "DA24"
        ExporToExcel ("select KEYFROM as ��ǩ�ͻ�, KEYNAME as ��ǩ����, KEYVALUE as ��ǩֵ, KEYTIME as ����, KEYBY as ��Ա  from UNIQUE_TBL_NEW order by KEYTIME desc")

    Case "GC"
        ExporToExcel ("select KEYFROM as �ͻ�, KEYNAME as ����, KEYVALUE as ���, KEYTIME as �˶�����, KEYBY as �˶���Ա  from UNIQUE_TBL_NEW where KEYFROM = 'GC' order by KEYTIME desc")

    Case "SH50"
        MsgBox "�ÿͻ�û����ſ��Ե���", vbInformation, "��ʾ"

    Case "HD"
        ExporToExcel ("select KEYFROM as �ͻ�, KEYNAME as ����, KEYVALUE as ���, KEYTIME as �˶�����, KEYBY as �˶���Ա  from UNIQUE_TBL_NEW where KEYFROM = 'HD' order by KEYTIME desc")

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
    .SetText 1, 0, "��  ��(C)"
    .SetText 2, 0, "��  ��(B)"
    .SetText 3, 0, "������(R)"
    .ColWidth(1) = 31
    .ColWidth(2) = 31
    .ColWidth(3) = 31
    .Row = 5

End With

cmbCombo1.AddItem ("HK037")
cmbCombo1.AddItem ("DA69")
cmbCombo1.AddItem ("AB18")
cmbCombo1.AddItem ("HK037_������_����")
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

    Case "HK037_������_����"
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
    MsgBox "��ɨ����ȷ�Ķ�ά��", vbInformation, "��ʾ"
    Exit Sub

End If

If chk(0).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "O" Then
            MsgBox "�����ǩ��ά��:��λO�ַ�����:", vbCritical, "����"
            Exit Sub

        End If

        If strPart(8) <> "000004" Then
            MsgBox "��ɨ��000004�����ǩ", vbInformation, "����"
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
        MsgBox "������ɨ��-C��ǩ", vbInformation, "��ʾ"
        Exit Sub

    End If

    chk(0).Value = 1
    Play ("�����ǩ��ɨ��")
ElseIf chk(1).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "I" Then
            MsgBox "�����ǩ��ά��:��λI�ַ�����:", vbCritical, "����"
            Exit Sub

        End If

        If strPart(8) <> "000003" Then
            MsgBox "��ɨ��000003�ںб�ǩ", vbInformation, "����"
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
    '        MsgBox "������ɨ��-B��ǩ", vbInformation, "��ʾ"
    '        Exit Sub
    '
    '    End If
    chk(1).Value = 1
    Play ("�����ǩ��ɨ��")
ElseIf chk(2).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "I" Then
            MsgBox "��������ǩ��ά��:��λI�ַ�����:", vbCritical, "����"
            Exit Sub

        End If

        If strPart(8) <> "000002" Then
            MsgBox "��ɨ��000002��������ǩ", vbInformation, "����"
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
        MsgBox "��������ɨ��-R��ǩ", vbInformation, "��ʾ"
        Exit Sub

    End If

    chk(2).Value = 1
    Play ("��������ǩ��ɨ��")
    gID = Get_OracleStr("select UNIQUE_SEQ.NEXTVAL from dual")
    '��ʼ�˶�
    CheckData
Else

End If

End Sub

Private Sub ListData_HK037_2(strCode As String)
Dim strPart() As String, i As Integer, lTmp As Long

strPart = Split(strCode, gSplitFlag)
If gMaxRow <> UBound(strPart) + 1 Then
    MsgBox "��ɨ����ȷ�Ķ�ά��", vbInformation, "��ʾ"
    Exit Sub

End If

If chk(0).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "O" Then
            MsgBox "�����ǩ��ά��:��λO�ַ�����:", vbCritical, "����"
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
        MsgBox "������ɨ��-C��ǩ", vbInformation, "��ʾ"
        Exit Sub

    End If

    chk(0).Value = 1
    Play ("�����ǩ��ɨ��")
ElseIf chk(1).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "I" Then
            MsgBox "�����ǩ��ά��:��λI�ַ�����:", vbCritical, "����"
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
        MsgBox "������ɨ��-B��ǩ", vbInformation, "��ʾ"
        Exit Sub

    End If

    chk(1).Value = 1
    Play ("�����ǩ��ɨ��")
ElseIf chk(2).Value = 0 Then
    If cmbCombo1.text = "HK037" Then
        If strPart(0) <> "I" Then
            MsgBox "��������ǩ��ά��:��λI�ַ�����:", vbCritical, "����"
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
        MsgBox "��������ɨ��-R��ǩ", vbInformation, "��ʾ"
        Exit Sub

    End If

    chk(2).Value = 1
    Play ("��������ǩ��ɨ��")
    gID = Get_OracleStr("select UNIQUE_SEQ.NEXTVAL from dual")
    '��ʼ�˶�
    CheckData2
Else

End If

End Sub

Private Sub ListData_HK037(strCode As String)
Dim strPart() As String, i As Integer, lTmp As Long

strPart = Split(strCode, gSplitFlag)
If gMaxRow <> UBound(strPart) + 1 Then
    MsgBox "��ɨ����ȷ�Ķ�ά��", vbInformation, "��ʾ"
    Exit Sub

End If

If chk(2).Value = 0 Then
    If strPart(0) <> "I" Then
        MsgBox "��������ǩ��ά��:��λI�ַ�����:", vbCritical, "����"
        Exit Sub

    End If

    With Fps(0)

        For i = 0 To UBound(strPart)
            .MaxRows = .MaxRows + 1
            .SetText E_Lbl.e_CARTON, i + 1, strPart(i)
        Next i

    End With

    chk(2).Value = 1
    Play ("��������ǩ��ɨ��")
ElseIf chk(3).Value = 0 Then
    If strPart(0) <> "I" Then
        MsgBox "���̱�ǩ��ά��:��λI�ַ�����:", vbCritical, "����"
        Exit Sub

    End If

    With Fps(0)

        For i = 0 To UBound(strPart)
            .MaxRows = .MaxRows + 1
            .SetText E_Lbl.E_BOX, i + 1, strPart(i)
        Next i

    End With

    chk(3).Value = 1
    Play ("���̱�ǩ��ɨ��")
    CheckData_HK037
Else

End If

End Sub

Private Sub ListData_DA24(strCode As String)
If chk(0).Value = 0 Then ' 1.����(C)
    strPart_C = Split(strCode, ";")
    If UBound(strPart_C) <> 5 Then
        MsgBox "�����ά�벻��ȷ", vbInformation, "��ʾ"
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
            MsgBox "��ɨ�����-C�������ά���ǩ", vbInformation, "��ʾ"
            Exit Sub

        End If

        If Get_OracleCnt("select * from UNIQUE_TBL_NEW where keyfrom = 'DA24' and keyname = 'PACKNO' and KEYVALUE = '" & strPart_C(5) & "'") > 0 Then
            .Col = 1
            .Row = 7
            .BackColor = vbRed
            MsgBox "ϵͳ�Ѿ�����ͬһ�������", vbInformation, "��ʾ"
            Exit Sub
        Else
            AddSql ("insert into UNIQUE_TBL_NEW(KEYNAME, KEYVALUE, KEYFROM,KEYTIME,KEYBY) values('PACKNO','" & strPart_C(5) & "','DA24',sysdate, '" & gUserName & "')")

        End If

    End With

    chk(0).Value = 1
    Play ("�����ǩ��ɨ��,��ɨ�������ǩ")
ElseIf chk(1).Value = 0 Then    ' 2.����(B)
    strPart_B = Split(strCode, ";")
    If UBound(strPart_B) <> 6 Then
        MsgBox "�����ά�벻��ȷ", vbInformation, "��ʾ"
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
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_B(1) <> strPart_C(1) Then
            '.Col = 2
            .Row = 2
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If InStr(strPart_C(4), strPart_B(5)) = 0 Then
            '.Col = 2
            .Row = 6
            .BackColor = vbRed
            MsgBox "��ǩ������ϵ��һ��", vbInformation, "��ʾ"

        End If

        If InStr(strPart_B(6), "-B") = 0 Then
            .Col = 2
            .Row = 7
            .BackColor = vbRed
            MsgBox "��ɨ�����-B�������ά���ǩ", vbInformation, "��ʾ"
            Exit Sub

        End If

        If CLng(strPart_B(2)) > CLng(strPart_C(2)) Then
            .Row = 3
            .BackColor = vbRed
            MsgBox "�����������ܴ�����������,����", vbInformation, "��ʾ"
            Exit Sub

        End If

        If Get_OracleCnt("select * from UNIQUE_TBL_NEW where keyfrom = 'DA24' and keyname = 'PACKNO' and KEYVALUE = '" & strPart_B(6) & "'") > 0 Then
            .Col = 2
            .Row = 7
            .BackColor = vbRed
            MsgBox "ϵͳ�Ѿ�����ͬһ�������", vbInformation, "��ʾ"
            Exit Sub
        Else
            AddSql ("insert into UNIQUE_TBL_NEW(KEYNAME, KEYVALUE, KEYFROM,KEYTIME,KEYBY) values('PACKNO','" & strPart_B(6) & "','DA24',sysdate, '" & gUserName & "')")

        End If

        chk(1).Value = 1
        Play ("�����ǩ��ɨ��,��ɨ����������ǩ")

    End With

ElseIf chk(2).Value = 0 Then    ' 3.������(R)
    strPart_R = Split(strCode, ";")
    If UBound(strPart_R) <> 6 Then
        MsgBox "�����ά�벻��ȷ", vbInformation, "��ʾ"
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
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_R(1) <> strPart_B(1) Then
            '.Col = 3
            .Row = 2
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_R(2) <> strPart_B(2) Then
            .Row = 3
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If InStr(strPart_C(4), strPart_R(5)) = 0 Then
            '.Col = 3
            .Row = 6
            .BackColor = vbRed
            MsgBox "��ǩ������ϵ��һ��", vbInformation, "��ʾ"

        End If

        If InStr(strPart_R(6), "-R") = 0 Then
            .Col = 3
            .Row = 7
            .BackColor = vbRed
            MsgBox "��ɨ�����-R����������ά���ǩ", vbInformation, "��ʾ"
            Exit Sub

        End If

        If Get_OracleCnt("select * from UNIQUE_TBL_NEW where keyfrom = 'DA24' and keyname = 'PACKNO' and KEYVALUE = '" & strPart_R(6) & "'") > 0 Then
            .Col = 3
            .Row = 7
            .BackColor = vbRed
            MsgBox "ϵͳ�Ѿ�����ͬһ���������", vbInformation, "��ʾ"
            Exit Sub
        Else
            AddSql ("insert into UNIQUE_TBL_NEW(KEYNAME, KEYVALUE, KEYFROM,KEYTIME,KEYBY) values('PACKNO','" & strPart_R(6) & "','DA24',sysdate, '" & gUserName & "')")

        End If

        If CLng(strPart_R(2)) = CLng(strPart_C(2)) Then
            InitCheckStatus
            Play ("��������ȫ���ȶ����,������ȶ���������")
        Else
            chk(1).Value = 0
            chk(2).Value = 0
            Play ("�������ѱȶ���,��������, ������ȶ���һ���ں�")

        End If

    End With

End If

End Sub

Private Sub ListData_SH50(strCode As String)
If chk(0).Value = 0 Then ' 1.����(C)
    strPart_C = Split(strCode, "@")
    If UBound(strPart_C) <> 6 Then
        MsgBox "�����ά�벻��ȷ", vbInformation, "��ʾ"
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
    Play ("�����ǩ��ɨ��,��ɨ�������ǩ")
ElseIf chk(1).Value = 0 Then    ' 2.����(B)
    strPart_B = Split(strCode, "@")
    If UBound(strPart_B) <> 6 Then
        MsgBox "�����ά�벻��ȷ", vbInformation, "��ʾ"
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
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_B(1) <> strPart_C(1) Then
            '.Col = 2
            .Row = 2
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_B(2) <> strPart_C(2) Then
            '.Col = 2
            .Row = 3
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_B(4) <> strPart_C(4) Then
            '.Col = 2
            .Row = 5
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_B(5) <> strPart_C(5) Then
            '.Col = 2
            .Row = 6
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_B(6) <> strPart_C(6) Then
            '.Col = 2
            .Row = 7
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If CLng(strPart_B(3)) > CLng(strPart_C(3)) Then
            .Row = 4
            .BackColor = vbRed
            MsgBox "�����������ܴ�����������,����", vbInformation, "��ʾ"
            Exit Sub

        End If

        chk(1).Value = 1
        Play ("�����ǩ��ɨ��,��ɨ����������ǩ")

    End With

ElseIf chk(2).Value = 0 Then    ' 3.������(R)
    strPart_R = Split(strCode, "@")
    If UBound(strPart_R) <> 6 Then
        MsgBox "�����ά�벻��ȷ", vbInformation, "��ʾ"
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
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_R(1) <> strPart_B(1) Then
            .Row = 2
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_R(2) <> strPart_B(2) Then
            .Row = 3
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If CLng(strPart_R(3)) <> CLng(strPart_B(3)) Then
            .Row = 4
            .BackColor = vbRed
            MsgBox "���������ں�������һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_R(4) <> strPart_B(4) Then
            .Row = 5
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_R(5) <> strPart_B(5) Then
            .Row = 6
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If strPart_R(6) <> strPart_B(6) Then
            .Row = 7
            .BackColor = vbRed
            MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
            Exit Sub

        End If

        If CLng(strPart_R(3)) = CLng(strPart_C(3)) Then
            InitCheckStatus
            Play ("��������ȫ���ȶ����,������ȶ���������")
        Else
            chk(1).Value = 0
            chk(2).Value = 0
            Play ("�������ѱȶ���,��������, ������ȶ���һ���ں�")

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

' �ж�״̬
If CLng(strBox) < CLng(strcarton) Then
    Cnn.CommitTrans
    Play ("�������Ѻ˶����, �������һ������")
    '        MsgBox "�������Ѻ˶����, �������һ������", vbInformation, "��ʾ"
    NextCheckStatus
Else
    Cnn.CommitTrans
    Play ("�������Ѻ˶����, �������һ������")
    '    MsgBox "�������Ѻ˶����, �������һ������", vbInformation, "��ʾ"
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

' �ж�״̬
Cnn.CommitTrans
Play ("�������Ѻ˶����, �������һ������")
'    MsgBox "�������Ѻ˶����, �������һ������", vbInformation, "��ʾ"
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
                Play ("����������̱�ǩΨһ���ظ�")
                MsgBox "����������̱�ǩΨһ���ظ�,��ȷ���Ƿ��ǩ�쳣", vbInformation, "����"
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
                MsgBox "�����������������Ϣ��һ��,��ȷ���Ƿ��ǩ�쳣", vbInformation, "����"
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
        Play ("��ǩ��һ��")
        MsgBox "����������ǩ��һ��,��ȷ���Ƿ��ǩ�쳣", vbInformation, "����"
        Exit Function

    End If

    If strReel <> strBox Then
        .Row = irow
        .Col = 2
        .BackColor = vbRed
        .Row = irow
        .Col = 3
        .BackColor = vbRed
        Play ("��ǩ��һ��")
        MsgBox "�������������ǩ��һ��,��ȷ���Ƿ��ǩ�쳣", vbInformation, "����"
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
                Play ("��ǩ��һ��")
                MsgBox "����������ǩ��һ��,��ȷ���Ƿ��ǩ�쳣", vbInformation, "����"
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
        Play ("��ǩ��һ��")
        MsgBox "�������������ǩ��һ��,��ȷ���Ƿ��ǩ�쳣", vbInformation, "����"
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
        Play ("��ǩ��������")
        MsgBox "��������������������,��ȷ���Ƿ��ǩ�쳣", vbInformation, "����"
        Exit Function

    End If

    If CLng(strReel) <> CLng(strBox) Then
        .Row = irow
        .Col = 2
        .BackColor = vbRed
        .Row = irow
        .Col = 3
        .BackColor = vbRed
        Play ("��ǩ��������")
        MsgBox "�������������ǩ������һ��,��ȷ���Ƿ��ǩ�쳣", vbInformation, "����"
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
            Play ("��ǩΨһ���ظ�")
            MsgBox "������ͬ��Ψһ��, ��ȷ���Ƿ�����", vbInformation, "����"
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
                MsgBox "������������ID���ܶ�Ӧ", vbInformation, "����"
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
' Project    :       ��ʽ����1
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
        MsgBox "��ɨ����ȷ��GC�����ά��,�������ǩģ���ѱ��", vbExclamation + vbOKOnly, "����"
        Exit Sub

    End If

    If Get_OracleCnt("select * from unique_tbl_new where KEYFROM = 'GC' and KEYNAME = '����' and KEYVALUE= '" & strArray(12) & "' ") > 0 Then
        MsgBox "�������:" & strArray(12) & vbCrLf & "֮ǰ�Ѿ��˶Թ�, ��ȷ�ϱ����Ƿ����ظ��쳣��ǩ", vbExclamation + vbOKOnly, "����"
        ExporToExcel ("select KEYFROM as �ͻ�, KEYNAME as ����, KEYVALUE as ���, KEYTIME as �˶�����, KEYBY as �˶���Ա  from UNIQUE_TBL_NEW where KEYFROM = 'GC' and KEYNAME = '����' and KEYVALUE= '" & strArray(12) & "'  order by KEYTIME desc")
        Exit Sub

    End If

    txtPackingNO.text = strArray(12)
    txtPackingQty.text = strArray(6)
    Play ("������ɨ��,������ɨ����������ǩ")
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
        MsgBox "��ɨ����ȷ����������ά��,�������ǩģ���ѱ��", vbExclamation + vbOKOnly, "����"
        Exit Sub

    End If

    For i = 0 To UBound(strLblInfo) - 1
        If (strLblInfo(i).strWaferID = Replace(strArray(0), "-", "")) Then
            bExisted = True
            If strLblInfo(i).strSecCode <> strArray(3) Then
                Play ("��������ǩ����ȷ")
                MsgBox "��������������: " & strArray(3) & "  ����ȷ", vbCritical, "����"
                Exit Sub

            End If

            If strLblInfo(i).strCodePP <> strArray(2) Then
                Play ("��������ǩ����ȷ")
                MsgBox "�������: " & strArray(2) & "  ����ȷ", vbCritical, "����"
                Exit Sub

            End If

            If strLblInfo(i).strCusDev <> Replace(strArray(1), "-3", "") Then
                MsgBox "���ִ���:" & strArray(1), vbCritical, "����"
                Exit Sub

            End If

            If strLblInfo(i).bChecked = True Then
                Play ("��ȷ���Ƿ��ظ�ɨ����ǩ����")
                MsgBox "�þ���:" & strLblInfo(i).strWaferID & "  �Ѿ��˶Թ�" & vbCrLf & "��ȷ���Ƿ��ظ�ɨ����ǩ����", vbCritical, "����"
                Exit Sub

            End If

            lQty = CLng(txtPackingQtyAdd.text) + strArray(5)
            If lQty = CLng(txtPackingQty.text) Then
                Play ("��������ȫ���˶���ȷ")
                AddSql ("insert into UNIQUE_TBL_NEW(KEYNAME, KEYVALUE, KEYFROM,KEYTIME,KEYBY) values('����','" & txtPackingNO.text & "','GC',sysdate, '" & gUserName & "')")
                MsgBox "ȫ���˶����", vbInformation, "��ʾ"
                clearLbl_GC
                Exit Sub
            ElseIf lQty > CLng(txtPackingQty.text) Then
                Play ("��������,���������������������")
                MsgBox "��������,���������������������", vbCritical, "����"
                Exit Sub

            End If

            Play ("��������ȷ")
            txtPackingQtyAdd.text = lQty
            strLblInfo(i).bChecked = True

        End If

    Next
    If bExisted = False Then
        Play ("��������ǩ����ȷ")
        MsgBox "������WaferID: " & Replace(strArray(0), "-", "") & "  ����ȷ", vbCritical, "����"
        Exit Sub

    End If

End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
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
If chk(0).Value = 0 Then ' 1.����(C)
    strBoxID = ""
    strPart_C = Split(strCode, "/")
    If UBound(strPart_C) <> 11 Then
        MsgBox "�����ά�벻��ȷ", vbInformation, "��ʾ"
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

    '0.��7�����Ϊ0
    If strPart_C(6) <> "0" Then
        Fps(0).Col = 1
        Fps(0).Row = 7
        Fps(0).BackColor = vbRed
        MsgBox "��ǩ��7�����Ϊ0", vbCritical, "����"
        Exit Sub

    End If

    '1.����ǩΨһ��
    strPSN = strPart_C(1)
    If InStr(strBoxID, strPSN) > 0 Then
        MsgBox "����Ψһ��:" & strPSN & "��ɨ��,�����ظ�", vbCritical, "����"
        Fps(0).Col = 1
        Fps(0).Row = 2
        Fps(0).BackColor = vbRed
        Exit Sub

    End If

    If Get_OracleCnt("select * from UNIQUE_TBL_NEW where KEYFROM = 'HD' and KEYVALUE = '" & strPSN & "' ") Then
        MsgBox "����Ψһ��:" & strPSN & "�Ѵ���,�����ظ�", vbCritical, "����"
        Fps(0).Col = 1
        Fps(0).Row = 2
        Fps(0).BackColor = vbRed
        Exit Sub

    End If

    '2.����ǩΨһ���5λ��־λ
    If Mid$(strPSN, 6, 1) <> "C" Then
        MsgBox "�����ǩΨһ��:" & strPSN & " ����λ����ΪC", vbCritical, "����"
        Fps(0).Col = 1
        Fps(0).Row = 2
        Fps(0).BackColor = vbRed
        Exit Sub

    End If

    '3.�������
    If strPart_C(5) > 60000 Then
        Fps(0).Col = 1
        Fps(0).Row = 6
        Fps(0).BackColor = vbRed
        MsgBox "��������������ɴ���60000", vbCritical, "����"
        Exit Sub

    End If

    '4.�����Ӵ���ѯ
    For i = 1 To UBound(Split(strPart_C(4), "|"))
        If Left(Split(strPart_C(4), "|")(i), 2) <> "0)" Then
            Fps(0).Col = 1
            Fps(0).Row = 5
            Fps(0).BackColor = vbRed
            MsgBox "��������:" & Split(strPart_C(4), "|")(i - 1) & "�Ĳ���Ʒ����" & Left(Split(strPart_C(4), "|")(i), 2) & ",��Ϊ0", vbCritical, "����"
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
        MsgBox "�������������ܺ�:" & strSumWX & " ������ʵ����������:" & strPart_C(5), vbCritical, "����"
        Exit Sub

    End If

    If Len(strPart_C(4)) - Len(Replace(strPart_C(4), "|", "")) > 8 Then
        Fps(0).Col = 1
        Fps(0).Row = 5
        Fps(0).BackColor = vbRed
        MsgBox "����������Ų��ɴ���8", vbCritical, "����"
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

    Play ("�����ǩ��ɨ��,��ɨ�������ǩ")
    strBoxID = strBoxID & strPSN & ","
    chk(0).Value = 1
ElseIf chk(1).Value = 0 Then    ' 2.����(B)
    strPart_B = Split(strCode, "/")
    If UBound(strPart_B) <> 11 Then
        MsgBox "�����ά�벻��ȷ", vbInformation, "��ʾ"
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
        '0.��7�����Ϊ0
        If strPart_B(6) <> "0" Then
            Fps(0).Col = 2
            Fps(0).Row = 7
            Fps(0).BackColor = vbRed
            MsgBox "��ǩ��7�����Ϊ0", vbCritical, "����"
            ClearIb
            Exit Sub

        End If

        '1.����ǩΨһ��
        strPSN = strPart_B(1)
        If InStr(strBoxID, strPSN) > 0 Then
            MsgBox "����Ψһ��:" & strPSN & "��ɨ��,�����ظ�", vbCritical, "����"
            Fps(0).Col = 2
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            Exit Sub

        End If

        If Get_OracleCnt("select * from UNIQUE_TBL_NEW where KEYFROM = 'HD' and KEYVALUE = '" & strPSN & "' ") Then
            Fps(0).Col = 2
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            MsgBox "����Ψһ��:" & strPSN & "�Ѵ���,�����ظ�", vbCritical, "����"
            ClearIb
            Exit Sub

        End If

        '2.����ǩΨһ���5λ��־λ
        If Mid$(strPSN, 6, 1) <> "B" Then
            Fps(0).Col = 2
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            MsgBox "�����ǩΨһ��:" & strPSN & " ����λ����ΪB", vbCritical, "����"
            ClearIb
            Exit Sub

        End If

        '3.��ͬ����
        strArray = Split(strSameItems, ",")

        For i = 0 To UBound(strArray)
            If strPart_B(strArray(i) - 1) <> strPart_C(strArray(i) - 1) Then
                Fps(0).Col = 1
                Fps(0).Row = strArray(i)
                Fps(0).BackColor = vbRed
                Fps(0).Col = 2
                Fps(0).Row = strArray(i)
                Fps(0).BackColor = vbRed
                MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
                ClearIb
                Exit Sub

            End If

        Next
        '4.��������
        If (Left(strPart_B(1), 5) <> Left(strPart_C(1), 5)) Or (Mid$(strPart_B(1), 7, 2) <> Mid$(strPart_C(1), 7, 2)) Then
            Fps(0).Col = 1
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            Fps(0).Col = 2
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            MsgBox "��ǩΨһ��:" & strPSN & " ����λ��һ��", vbCritical, "����"
            ClearIb
            Exit Sub

        End If

        '5.�������
        If strPart_B(5) > 15000 Then
            Fps(0).Col = 2
            Fps(0).Row = 6
            Fps(0).BackColor = vbRed
            MsgBox "��������������ɴ���15000", vbCritical, "����"
            ClearIb
            Exit Sub

        End If

        '6.���ڼ��
        If Abs(DateDiff("d", strPart_B(7), strPart_C(7))) > 30 Then
            Fps(0).Col = 1
            Fps(0).Row = 8
            Fps(0).BackColor = vbRed
            Fps(0).Col = 2
            Fps(0).Row = 8
            Fps(0).BackColor = vbRed
            MsgBox "�����ǩ�������ǩ���ڼ�����ɳ�����ʮ��", vbCritical, "���ڴ���"
            ClearIb
            Exit Sub

        End If

        '7.�����Ӵ���ѯ
        For i = 1 To UBound(Split(strPart_B(4), "|"))
            If Left(Split(strPart_B(4), "|")(i), 2) <> "0)" Then
                Fps(0).Col = 2
                Fps(0).Row = 5
                Fps(0).BackColor = vbRed
                MsgBox "��������:" & Split(strPart_B(4), "|")(i - 1) & "�Ĳ���Ʒ����" & Left(Split(strPart_B(4), "|")(i), 2) & ",��Ϊ0", vbCritical, "����"
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
            MsgBox "�������������ܺ�:" & strSumNH & " ������ʵ����������:" & strPart_B(5), vbCritical, "����"
            Exit Sub

        End If

        If Len(strPart_B(4)) - Len(Replace(strPart_B(4), "|", "")) > 5 Then
            Fps(0).Col = 2
            Fps(0).Row = 5
            Fps(0).BackColor = vbRed
            MsgBox "����������Ų��ɴ���5", vbCritical, "����"
            Exit Sub

        End If

        Dim strArrNH() As String

        strArrNH = Split(strPart_B(4), "|0)")
        Fps(0).Col = 1
        Fps(0).Row = 5

        For i = 0 To UBound(strArrNH) - 1
            If InStr(.text, strArrNH(i)) = 0 Then
                MsgBox "��������:" & .text & "����������������:" & strArrNH(i)
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
        Play ("�����ǩ��ɨ��,��ɨ����������ǩ")
        strBoxID = strBoxID & strPSN & ","

    End With

ElseIf chk(2).Value = 0 Then    ' 3.������(R)
    strPart_R = Split(strCode, "/")
    If UBound(strPart_R) <> 11 Then
        MsgBox "�����ά�벻��ȷ", vbInformation, "��ʾ"
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
        '0.��7�����Ϊ0
        If strPart_R(6) <> "0" Then
            Fps(0).Col = 3
            Fps(0).Row = 7
            Fps(0).BackColor = vbRed
            MsgBox "��ǩ��7�����Ϊ0", vbCritical, "����"
            ClearLv
            Exit Sub

        End If

        '1.����ǩΨһ��
        strPSN = strPart_R(1)
        If InStr(strBoxID, strPSN) > 0 Then
            MsgBox "������Ψһ��:" & strPSN & "��ɨ��,�����ظ�", vbCritical, "����"
            Fps(0).Col = 3
            Fps(0).Row = 2
            Fps(0).BackColor = vbRed
            Exit Sub

        End If

        If Get_OracleCnt("select * from UNIQUE_TBL_NEW where KEYFROM = 'HD' and KEYVALUE = '" & strPSN & "' ") Then
            MsgBox "������Ψһ��:" & strPSN & "�Ѵ���,�����ظ�", vbCritical, "����"
            ClearLv
            Exit Sub

        End If

        '2.����ǩΨһ���5λ��־λ
        If Mid$(strPSN, 6, 1) <> "A" Then
            MsgBox "��������ǩΨһ��:" & strPSN & " ����λ����ΪA", vbCritical, "����"
            ClearLv
            Exit Sub

        End If

        '3.�������
        If strPart_R(5) > 3000 Then
            Fps(0).Col = 3
            Fps(0).Row = 6
            Fps(0).BackColor = vbRed
            MsgBox "����������������ɴ���3000", vbCritical, "����"
            ClearLv
            Exit Sub

        End If

        '4.��ͬ����
        strArray = Split(strSameItems, ",")

        For i = 0 To UBound(strArray)
            If strPart_R(strArray(i) - 1) <> strPart_B(strArray(i) - 1) Then
                .Col = 2
                .Row = strArray(i)
                .BackColor = vbRed
                .Col = 3
                .Row = strArray(i)
                .BackColor = vbRed
                MsgBox "��ǩ��һ��", vbInformation, "��ʾ"
                ClearLv
                Exit Sub

            End If

        Next
        '5.���ڼ��
        If Abs(DateDiff("d", strPart_R(7), strPart_B(7))) > 30 Then
            Fps(0).Col = 3
            Fps(0).Row = 8
            Fps(0).BackColor = vbRed
            Fps(0).Col = 2
            Fps(0).Row = 8
            Fps(0).BackColor = vbRed
            MsgBox "��������ǩ���ںб�ǩ���ڼ�����ɳ�����ʮ��", vbCritical, "���ڴ���"
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
            MsgBox "��������ǩ�������ǩ���ڼ�����ɳ�����ʮ��", vbCritical, "���ڴ���"
            ClearLv
            Exit Sub

        End If

        '6.�����Ӵ���ѯ
        For i = 1 To UBound(Split(strPart_R(4), "|"))
            If Left(Split(strPart_R(4), "|")(i), 2) <> "0)" Then
                Fps(0).Col = 3
                Fps(0).Row = 5
                Fps(0).BackColor = vbRed
                MsgBox "����������:" & Split(strPart_R(4), "|")(i - 1) & "�Ĳ���Ʒ����" & Left(Split(strPart_R(4), "|")(i), 2) & ",��Ϊ0", vbCritical, "����"
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
            MsgBox "���������������ܺ�:" & strSumLV & " ������ʵ������������:" & strPart_R(5), vbCritical, "����"
            Exit Sub

        End If

        If Len(strPart_R(4)) - Len(Replace(strPart_R(4), "|", "")) > 2 Then
            Fps(0).Col = 3
            Fps(0).Row = 5
            Fps(0).BackColor = vbRed
            MsgBox "�������������Ų��ɴ���2", vbCritical, "����"
            Exit Sub

        End If

        Dim strArrLV() As String

        strArrLV = Split(strPart_R(4), "|0)")
        Fps(0).Col = 2
        Fps(0).Row = 5

        For i = 0 To UBound(strArrLV) - 1
            If InStr(.text, strArrLV(i)) = 0 Then
                MsgBox "��������:" & .text & "������������������:" & strArrLV(i)
                Fps(0).Col = 3
                Fps(0).Row = 5
                Fps(0).BackColor = vbRed
                Exit Sub

            End If

        Next
        '7.�����ۼ�
        gLVCntSum = gLVCntSum + CLng(strPart_R(5))
        txtLvCnt.text = gLVCntSum
        If gLVCntSum = CLng(strPart_B(5)) Then
            strBoxID = strBoxID & strPSN & ","
            chk(1).Value = 0
            chk(2).Value = 0
            gLVCntSum = 0
            txtLvCnt.text = gLVCntSum
            Play ("���ں��Ѿ��˶����,�������һ���ں�")
            ClearLv
            ClearIb
            If gIBCntSum = CLng(strPart_C(5)) Then
                Call InitCheckStatus
                Call SaveBoxID
                Play ("��������ȫ���ȶ����,������ȶ���������")
            ElseIf gIBCntSum > CLng(strPart_C(5)) Then
                MsgBox "�ں������ۼ��ܺ�" & gIBCntSum & "���ɴ�����������:" & strPart_C(5), vbCritical, "��������"

            End If

        ElseIf gLVCntSum > CLng(strPart_B(5)) Then
            MsgBox "�����������ۼ��ܺ�:" & gLVCntSum & "���ɴ�����������:" & strPart_B(5), vbCritical, "��������"
            ClearLv
            Exit Sub
        Else
            Play ("����������ɨ�������ɨ����һ��������")
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
    strSql = "insert into UNIQUE_TBL_NEW(KEYNAME,KEYVALUE,KEYFROM,KEYTIME,KEYBY) values('���Ψһ��','" & strBoxIDArray(i) & "','HD',sysdate,'" & gUserName & "') "
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
' Project    :       ��ʽ����1
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


If chk(0).Value = 0 Then ' 1.����(C)
    strPart_C = Split(strCode, ",")
    If UBound(strPart_C) <> 7 Then
        MsgBox "�����ά�벻��ȷ", vbInformation, "��ʾ"
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
    Play ("�����ǩ��ɨ��,��ɨ�������ǩ")
    chk(0).Value = 1
ElseIf chk(1).Value = 0 Then    ' 2.����(B)
    strPart_B = Split(strCode, ",")
    If UBound(strPart_B) <> 7 Then
        MsgBox "�����ά�벻��ȷ", vbInformation, "��ʾ"
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
            MsgBox "�ͻ����ֲ�һ��", vbCritical, "����"
            Fps(0).DeleteRows .MaxRows, 1
            Fps(0).MaxRows = .MaxRows - 1
            Exit Sub

        End If

        '2.LotID
        If strPart_C(2) <> strPart_B(1) Then
            .Col = 2
            .Row = .MaxRows
            .BackColor = vbRed
            MsgBox "LotID��һ��", vbCritical, "����"
            Fps(0).DeleteRows .MaxRows, 1
            Fps(0).MaxRows = .MaxRows - 1
            Exit Sub

        End If

    
        '3.Date code
        If strPart_B(3) <> strPart_C(6) Then
            Fps(0).Col = 5
            Fps(0).Row = Fps(0).MaxRows
            Fps(0).BackColor = vbRed
            MsgBox "DateCode��һ��", vbCritical, "����"
            Fps(0).DeleteRows .MaxRows, 1
            Fps(0).MaxRows = .MaxRows - 1
            Exit Sub

        End If

        '4.HT Lot
        If Left(strPart_B(4), 8) <> Left(strPart_C(7), 8) Then
            Fps(0).Col = 6
            Fps(0).Row = Fps(0).MaxRows
            Fps(0).BackColor = vbRed
            MsgBox "�������Ų�һ��", vbCritical, "����"
            Fps(0).DeleteRows .MaxRows, 1
            Fps(0).MaxRows = .MaxRows - 1
            Exit Sub

        End If
        
        '2.Wafer ID
        If InStr(strPart_C(3), strPart_B(5)) = 0 Then
            Fps(0).Col = 3
            Fps(0).Row = Fps(0).MaxRows
            Fps(0).BackColor = vbRed
            MsgBox "WaferID������", vbCritical, "����"
            Fps(0).DeleteRows Fps(0).MaxRows, 1
            Fps(0).MaxRows = .MaxRows - 1
            Exit Sub

        End If

        
        If Replace(Replace$(strPart_C(3), strPart_B(5), ""), " ", "") = "" Then
            If lNXQty + CLng(strPart_B(2)) = lWXQty Then
                
                Call InitCheckStatus
                Play ("��������ȫ���˶����,��˶���������")
            Else
        
                Fps(0).Col = 4
                Fps(0).Row = Fps(0).MaxRows
                Fps(0).BackColor = vbRed
                MsgBox "��������������Ӧ", vbCritical, "����"
                Fps(0).DeleteRows .MaxRows, 1
                Fps(0).MaxRows = .MaxRows - 1
                Exit Sub

            End If

        Else
            Play ("���ں���ɨ��,��ɨ���¸��ں�")
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
