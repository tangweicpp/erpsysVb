VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form Frm_WaferSide 
   Caption         =   "���ά��"
   ClientHeight    =   12525
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
   ScaleHeight     =   12525
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   12615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   17205
      _ExtentX        =   30348
      _ExtentY        =   22251
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "��λά��"
      TabPicture(0)   =   "Frm_WaferSide.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtPath"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cbType"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtPN"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtLot"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Toolbar1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ImageList1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "com"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Fps(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblType"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblPN"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblLot"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "�����ߴ�ά��"
      TabPicture(1)   =   "Frm_WaferSide.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtCusCode"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Fps(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Toolbar2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtNo"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   12
         Top             =   1260
         Width           =   4575
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   -73320
         TabIndex        =   8
         Top             =   1800
         Width           =   7695
      End
      Begin VB.ComboBox cbType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Frm_WaferSide.frx":0038
         Left            =   -73335
         List            =   "Frm_WaferSide.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtPN 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -68940
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtLot 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -64320
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   450
         Left            =   -74400
         TabIndex        =   1
         Top             =   480
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   794
         ButtonWidth     =   1508
         ButtonHeight    =   741
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ѯ"
               Key             =   "QUERY"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "UPLOAD"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "SAVE"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "EXIT"
               ImageIndex      =   4
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -63480
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   22
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WaferSide.frx":005A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WaferSide.frx":00B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WaferSide.frx":0D0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_WaferSide.frx":195C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog com 
         Left            =   -64080
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   10455
         Index           =   0
         Left            =   -74400
         TabIndex        =   9
         Top             =   2160
         Width           =   8775
         _Version        =   524288
         _ExtentX        =   15478
         _ExtentY        =   18441
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
         MaxCols         =   4
         MaxRows         =   0
         SpreadDesigner  =   "Frm_WaferSide.frx":1CAE
         TextTip         =   2
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   450
         Left            =   360
         TabIndex        =   15
         Top             =   600
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   794
         ButtonWidth     =   1508
         ButtonHeight    =   741
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ѯ"
               Key             =   "QUERY"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "����"
               Key             =   "UPLOAD"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "SAVE"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "EXIT"
               ImageIndex      =   4
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   9615
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   1920
         Width           =   11295
         _Version        =   524288
         _ExtentX        =   19923
         _ExtentY        =   16960
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
         MaxCols         =   4
         MaxRows         =   0
         SpreadDesigner  =   "Frm_WaferSide.frx":2156
         Appearance      =   1
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin MSForms.TextBox txtCusCode 
         Height          =   285
         Left            =   7800
         TabIndex        =   14
         Top             =   1305
         Width           =   1335
         VariousPropertyBits=   746604563
         BorderStyle     =   1
         Size            =   "2355;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����"
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
         Height          =   240
         Left            =   6840
         TabIndex        =   13
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ⵥ���"
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
         Height          =   240
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴ��ļ�:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   240
         Left            =   -74400
         TabIndex        =   10
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ά������"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   -74400
         TabIndex        =   7
         Top             =   1365
         Width           =   1020
      End
      Begin VB.Label lblPN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ�����Ϻ�"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Left            =   -70215
         TabIndex        =   6
         Top             =   1365
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblLot 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ��������"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Left            =   -65640
         TabIndex        =   5
         Top             =   1365
         Visible         =   0   'False
         Width           =   1275
      End
   End
End
Attribute VB_Name = "Frm_WaferSide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbType_Click()
    lblPN.Visible = True
    TxtPN.Visible = True

    Select Case cbType.ListIndex

        Case 0  ' ԭ����
            lblPN.Caption = "ԭ�����Ϻ�"
            lblLOT.Caption = "ԭ��������"
            lblLOT.Visible = True
            txtLot.Visible = True
            
        Case 1  ' ��Ʒ
            lblPN.Caption = "��Ʒ�����"
            lblLOT.Caption = "��������"

    End Select

End Sub

Private Sub Form_Load()

    With Fps(0)
        .Col = -1
        .Row = -1
        .Lock = True
        
        .Col = 1
        .Row = 0
        .FontSize = 5
        
        .Col = 2
        .Row = 0
        .FontSize = 5
        
        .Col = 3
        .Row = 0
        .FontSize = 5
        
        .Col = 4
        .Row = 0
        .FontSize = 5
        
        .ColWidth(1) = 15
        .ColWidth(2) = 15
        .ColWidth(3) = 15
        .ColWidth(4) = 10

    End With
    
    With Fps(1)
        .MaxRows = 0
    
        .DAutoCellTypes = False

        .Col = -1
        .Row = -1
        .Lock = True

        .Col = 3
        .Lock = False
        .BackColor = vbGreen
        
        .Col = 4
        .Lock = False
        .BackColor = vbGreen
        .CellType = CellTypeComboBox
        .TypeComboBoxList = .TypeComboBoxList & "26*23*24"
        .TypeComboBoxList = .TypeComboBoxList & "32*28*16"
        .TypeComboBoxList = .TypeComboBoxList & "33*29*31"
        .TypeComboBoxList = .TypeComboBoxList & "34*34*32"
        .TypeComboBoxList = .TypeComboBoxList & "37*18*10"
        .TypeComboBoxList = .TypeComboBoxList & "40*38*26"
        .TypeComboBoxList = .TypeComboBoxList & "42*36*39"
        .TypeComboBoxList = .TypeComboBoxList & "43*39*40"
        .TypeComboBoxList = .TypeComboBoxList & "44*26*24"
        .TypeComboBoxList = .TypeComboBoxList & "44*40*39"
        .TypeComboBoxList = .TypeComboBoxList & "44*44*24"
        .TypeComboBoxList = .TypeComboBoxList & "57*33*31"
        .TypeComboBoxList = .TypeComboBoxList & "58*34*42"
        .TypeComboBoxList = .TypeComboBoxList & "60*41*23"
        .TypeComboBoxList = .TypeComboBoxList & "60*59*60"
        
        .ColWidth(1) = 15
        .ColWidth(2) = 15
        .ColWidth(3) = 15
        .ColWidth(4) = 10
    End With

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key

        Case "QUERY"
            QueryData
    
        Case "UPLOAD"
            UploadData
            
        Case "SAVE"
            SaveData
            
        Case "EXIT"
            Unload Me

    End Select

End Sub

Private Sub QueryData()

    If cbType.text = "" Then
        MsgBox "��ѡ��ά������", vbInformation, "��ʾ"
        Exit Sub

    End If

    Select Case cbType.ListIndex

        Case 0  ' ԭ����

            If TxtPN.text = "" Then
                MsgBox "��������Ҫ��ѯ���Ϻ�", vbInformation, "��ʾ"
                Exit Sub

            End If
            
            showData_PN

        Case 1  ' ��Ʒ

            If TxtPN.text = "" And txtLot.text = "" Then
                MsgBox "��������Ҫ��ѯ�Ĵ���Ż��������", vbInformation, "��ʾ"
                Exit Sub

            End If
            
            showData_QBOXNO

    End Select

End Sub

Private Sub showData_PN()

    Dim strSql As String

    Dim rs     As New ADODB.Recordset
    
    If txtLot.text = "" Then
        strSql = "select AA.�ֿ���,AA.��λ,BB.F_101 as �Ϻ�,AA.����,BB.FName as ��������,BB.FNumber as ���ϱ��, AA.��Ч����,AA.��������, AA.��ǰ����,AA.id from erpbase.dbo.tblStockNum AA  INNER JOIN AIS20141114094336.dbo.t_ICItem BB ON AA.���ϱ�� = BB.FNumber AND   BB.F_101 = '" & UCase(Trim(TxtPN.text)) & "' and ��ǰ���� > 0 "
    Else
        strSql = "select AA.�ֿ���,AA.��λ,BB.F_101 as �Ϻ�,AA.����,BB.FName as ��������,BB.FNumber as ���ϱ��, AA.��Ч����,AA.��������, AA.��ǰ����,AA.id from erpbase.dbo.tblStockNum AA  INNER JOIN AIS20141114094336.dbo.t_ICItem BB ON AA.���ϱ�� = BB.FNumber AND   BB.F_101 = '" & UCase(Trim(TxtPN.text)) & "' and AA.���� = '" & UCase(Trim(txtLot.text)) & "' and ��ǰ���� > 0"

    End If
        
    Set rs = Get_SqlserveRs(strSql)

    If Not rs.EOF Then

        With Fps(0)
            .MaxRows = 0
            Set .DataSource = rs

        End With

    Else
        MsgBox "��ѯ�������Ϻ�", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

Private Sub showData_QBOXNO()

    Dim strSql As String

    Dim rs     As New ADODB.Recordset
    
    If txtLot.text = "" Then
        strSql = "SELECT DX.��� as �����, XX.��� as С���,DX.��λ as ��λ FROM erpdata..tblStockNumTree DX inner join erpdata..tblStockNumTree XX on XX.�ϼ���� = DX.��� where DX.��� = '" & TxtPN.text & "'"
       
    Else
        strSql = "SELECT DX.��� as �����, XX.��� as С���,DX.��λ as ��λ FROM erpdata..tblStockNumTree DX inner join erpdata..tblStockNumTree XX on XX.�ϼ���� = DX.��� where DX.��� = '" & TxtPN.text & "'"

    End If

    Set rs = Get_SqlserveRs(strSql)

    If Not rs.EOF Then

        With Fps(0)
            .MaxRows = 0
            Set .DataSource = rs

        End With

    Else
        MsgBox "��ѯ�����ó�Ʒ�����", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

Private Sub showData_upload()

    On Error GoTo ErrHandle

    Dim VBExcel     As Excel.Application

    Dim xlBook      As Excel.Workbook

    Dim xlSheet     As Excel.Worksheet

    Dim strFileName As String

    Dim i           As Integer

    Dim j           As Integer

    Dim strChar     As String

    Dim strTmp(10)  As Variant
    
    MousePointer = 11

    Fps(0).MaxRows = 0

    If InStrRev(Trim(txtPath.text), "\") > 0 Then
        strFileName = Mid(Trim(txtPath.text), InStrRev(Trim(txtPath.text), "\") + 1)

        If InStr(strFileName, ".") > 0 Then
            strFileName = Mid(strFileName, 1, InStr(strFileName, ".") - 1)

        End If

    End If

    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(txtPath.text)
    Set xlSheet = xlBook.Worksheets(1)
  
    If xlSheet.Range("A1").CurrentRegion.Columns.count < 2 Then
        MousePointer = 0
        MsgBox "Excel�е��������趨��ģ��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        GoTo EXITPRO
        Exit Sub

    End If

    With Fps(0)

        For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.count
            strTmp(0) = Trim(xlSheet.Range("A" & i).Value)

            If Len(strTmp(0)) > 0 Then
                If i <> 1 Then .MaxRows = .MaxRows + 1

                For j = 1 To 4

                    If j > 26 Then
                        strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
                    Else
                        strChar = Chr(96 + j)

                    End If

                    If i = 1 Then
                        .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))
                    Else
                        .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))

                    End If

                Next

            End If

        Next

    End With

    MousePointer = 0
    
    xlBook.Close
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
    VBExcel.Quit
    
    Exit Sub
EXITPRO:

    On Error Resume Next

    MousePointer = 0

    If Not VBExcel Is Nothing Then
        xlBook.Close
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing
        VBExcel.Quit

    End If

    Exit Sub
ErrHandle:
    GoTo EXITPRO

End Sub

Private Sub UploadData()

    On Error GoTo ErrHandler

    Dim FName
    
    com.Filter = "�����ļ�(*.*)|*.*|Excel�ļ�(*.xls;*.xlsx)|*.xls;*.xlsx"
    com.ShowOpen

    FName = com.filename

    If FName <> "" Then
        txtPath.text = FName
        showData_upload

    End If
    
    Exit Sub
ErrHandler:
    MsgBox "ȡ��"
    Exit Sub

End Sub

Private Sub SaveData()

    If cbType.text = "" Then
        MsgBox "��ѡ��ά������", vbInformation, "��ʾ"
        Exit Sub

    End If

    saveData_PN

End Sub

Private Sub saveData_PN()

On Error GoTo ErrHandle

Dim strSql    As String
Dim i         As Integer
Dim strLH     As String, strKF As String, strPH As String, strCW As String
Dim strQboxNO As String, strno As String, strNO_2 As String

If Fps(0).MaxRows <= 0 Then
    MsgBox "û��Ҫ���������", vbInformation, "��ʾ"
    Exit Sub

End If

If MsgBox("�Ƿ�Ҫ������", vbInformation + vbYesNo, "��ʾ") = vbNo Then Exit Sub
MousePointer = 11

With Fps(0)

    For i = 1 To .MaxRows
        .Row = i
        If cbType.ListIndex = 0 Then
            .Col = 1
            strLH = Trim$(.text)    ' �Ϻ�
            .Col = 2
            strKF = Left(Trim$(.text), 2)   ' �ⷿ
            .Col = 3
            strPH = Trim$(.text)    ' ����
            .Col = 4
            strCW = Trim$(.text)    ' ��λ
            strSql = "update ERPBASE..tBLSTOCKNUM set ��λ = '" & strCW & "' where ���ϱ�� in (select aa.FNumber from  AIS20141114094336.dbo.t_ICItem aa where aa.F_101 = '" & strLH & "') and ���� = '" & strPH & "' and �ֿ��� = '" & strKF & "'  and ��ǰ���� > 0"
            '����ֵ���
            'strSql = "update ERPBASE..tBLSTOCKNUM set ��λ = '" & strCW & "' where CHARINDEX('" & strLH & "',���ϱ��) > 0 and ���� = '" & strPH & "' "
            If AddSql2(strSql) = 0 Then
                MsgBox "ԭ����:" & strLH & " û�и��µ��¿�λ", vbInformation, "��ʾ"

            End If

        Else
            .Col = 1
            strQboxNO = Trim$(.text)    ' �Ϻ�
            .Col = 2
            strCW = Trim$(.text)
            strSql = "select �ϼ���� from erpdata..tblStockNumTree where ��� = '" & strQboxNO & "' "
            strno = Get_SqlStr(strSql)
            If strno <> "0" Then
                ' С���
                strSql = "update erpdata..tblStockNumTree set ��λ = '" & strCW & "' where ��� = '" & strQboxNO & "' "
                If AddSql2(strSql) = 0 Then
                    MsgBox "С���:" & strCW & " û�и��µ���λ", vbInformation, "��ʾ"

                End If

                strSql = "update erpdata..tblStockNumTree set ��λ = '" & strCW & "' where ��� = '" & strno & "' "
                If AddSql2(strSql) = 0 Then
                   ' MsgBox "С���:" & strCW & " û�и��µ���λ", vbInformation, "��ʾ"

                End If

            Else
                ' �����
                strSql = "update erpdata..tblStockNumTree set ��λ = '" & strCW & "' where ��� = '" & strQboxNO & "' "
                If AddSql2(strSql) = 0 Then
                    MsgBox "�����:" & strCW & " û�и��µ���λ", vbInformation, "��ʾ"

                End If

                strSql = "select ��� from erpdata..tblStockNumTree where ��� = '" & strQboxNO & "'"
                strNO_2 = Get_SqlStr(strSql)
                If strNO_2 <> "" Then
                    strSql = "update erpdata..tblStockNumTree set ��λ = '" & strCW & "' where �ϼ���� = '" & strNO_2 & "' "
                    If AddSql2(strSql) = 0 Then
                        'MsgBox "�����:" & strCW & " û�и��µ���λ", vbInformation, "��ʾ"

                    End If

                End If

            End If

        End If

    Next

End With

MousePointer = 0
MsgBox "���ϱ���ɹ���", vbInformation, Me.Caption
Exit Sub
ErrHandle:
MousePointer = 0
MsgBox Err.DESCRIPTION, vbCritical + vbInformation, "����"

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "QUERY"
        queryWeightSize
    
    Case "SAVE"
        saveWeightSize
        
    Case "EXIT"
        Unload Me

End Select

End Sub

Private Sub queryWeightSize()

Dim strno As String, strCusCode As String, strSql As String
strno = Trim(txtNo.text)
If Len(strno) = 0 Then
    MsgBox "��������ⵥ��", vbInformation, "��ʾ"
    Exit Sub
End If

strSql = "SELECT �ͻ����� FROM erpdata..tblPackToHouse where ��ⵥ��� = '" & strno & "' "
strCusCode = Get_SqlStr(strSql)
If "" = strCusCode Then
    MsgBox "�������ⵥ�Ų���ȷ�򲻴���", vbInformation, "����"
    Exit Sub
End If

txtCusCode.text = strCusCode

showWeightSize

End Sub

Private Sub showWeightSize()

    Dim strSql As String, strno As String

    Dim rs     As New ADODB.Recordset

    strno = Trim(txtNo.text)

    strSql = "select AA.��ⵥ���,BB.���,BB.���� as ����KG,BB.�ߴ� as �ߴ�CM,BB.��λ  from erpdata..tblPackToHouseSub  AA " & "inner join erpdata..tblStockNumTree  BB on AA.��� = BB.��� and BB.������ = '1' " & "where AA.��ⵥ��� = '" & strno & "' "

    Set rs = Get_SqlserveRs(strSql)

    If Not rs.EOF Then

        With Fps(1)
            .MaxRows = 0
            Set .DataSource = rs

        End With

    Else
        MsgBox "��ѯ���������Ϣ", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

Private Sub saveWeightSize()

Dim strQboxNO As String, strWeight As String, strSize As String
Dim strSql As String
Dim i As Integer

With Fps(1)
    For i = 1 To .MaxRows
        .Row = i
        .Col = 2
        strQboxNO = Trim$(.text)
        
        .Col = 3
        strWeight = Trim$(.text)
        
        .Col = 4
        strSize = Trim$(.text)
        
        strSql = "update erpdata..tblStockNumTree set �ߴ� = '" & strSize & "', ���� = '" & strWeight & "' where ��� = '" & strQboxNO & "'"
        If AddSql2(strSql) = 0 Then
            MsgBox "���" & strQboxNO & "  û��ά���ɹ�", vbInformation, "��ʾ"
        End If
    Next
End With

MsgBox "ά���ɹ�", vbInformation, "��ʾ"

Call showWeightSize

End Sub
