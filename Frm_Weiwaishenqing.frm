VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Weiwaishenqing 
   Caption         =   "Form1"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18435
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
   ScaleHeight     =   10335
   ScaleWidth      =   18435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin FPSpreadADO.fpSpread fpS_Box 
      Height          =   6615
      Left            =   7560
      TabIndex        =   16
      Top             =   3240
      Width           =   10695
      _Version        =   524288
      _ExtentX        =   18865
      _ExtentY        =   11668
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
      MaxCols         =   10
      MaxRows         =   0
      SpreadDesigner  =   "Frm_Weiwaishenqing.frx":0000
      AppearanceStyle =   0
   End
   Begin FPSpreadADO.fpSpread fpS_Lot 
      Height          =   6615
      Left            =   480
      TabIndex        =   15
      Top             =   3240
      Width           =   6855
      _Version        =   524288
      _ExtentX        =   12091
      _ExtentY        =   11668
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
      MaxCols         =   10
      MaxRows         =   0
      SpreadDesigner  =   "Frm_Weiwaishenqing.frx":0422
      AppearanceStyle =   0
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7455
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   13150
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   15255
      Begin VB.ComboBox CbstockId_org 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox ComCbbond 
         Height          =   315
         ItemData        =   "Frm_Weiwaishenqing.frx":0844
         Left            =   4560
         List            =   "Frm_Weiwaishenqing.frx":084E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtCustLot 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原仓库"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3600
         TabIndex        =   13
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "保税"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "厂内机种"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label lblLOTID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOTID"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户机种"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户代码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   840
      End
      Begin MSForms.ComboBox cbCustCode 
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   360
         Width           =   1935
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3413;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18435
      _ExtentX        =   32517
      _ExtentY        =   1111
      ButtonWidth     =   2090
      ButtonHeight    =   1058
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  查  询"
            Key             =   "QUERY"
            ImageIndex      =   7
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  申请"
            Key             =   "REQUEST"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "A004"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "调拨"
            Key             =   "MOVE"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退回"
            Key             =   "CANCEL_PASS"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "导出申请记录"
            Key             =   "EXPORT_SOD"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  退  出"
            Key             =   "EXIT"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   12360
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":0860
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":299A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":5824
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":7FD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":A110
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":C8C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":F074
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":120F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":148A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":14BC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":1589C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":1891E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_Weiwaishenqing.frx":1B0D0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Frm_Weiwaishenqing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCustCode As String
Dim strCustLot As String

Enum E_LOT

    E_CHOOSE = 1
    E_CUSTCODE
    E_LOTID
    E_STOCKID
    E_PN
    E_TOTALQTY
    E_PASSQTY
    E_NGQTY1
    E_NGQTY2
    E_JOBID
    E_ID
    E_END

End Enum

Enum E_BOX

    E_CHOOSE = 1
    E_LOTID
    E_STOCKID
    E_BOXID
    E_ID
    E_END

End Enum

Enum E_WaferId

    E_CHOOSE = 1
    E_BOXID
    E_WaferId
    E_QTY
    
    E_END

End Enum


Private Sub cbCustCode_DropButtonClick()

    Dim rs          As New ADODB.Recordset
    Dim strSql      As String
       

    Set rs = Get_SqlserveRs("select distinct 客户代码  from erpbase..tblXCustomer ")
    
    cbCustCode.Clear
    If Not rs.EOF Then
        rs.MoveFirst
    
        Do While Not rs.EOF
            cbCustCode.AddItem Trim(rs("客户代码"))
            rs.MoveNext
        Loop
    
    End If
    

    Set rs = Nothing
    
End Sub


Private Sub CbstockId_org_DropDown()

    Dim rs          As New ADODB.Recordset

    Dim strSql      As String

    strSql = "SELECT DISTINCT a.库房代码 + Space(1) + a.库房名称 as 库房 FROM erpbase..tblstock a "

    Set rs = Get_SqlserveRs(strSql)
    CbstockId_org.Clear

    If Not rs.EOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            CbstockId_org.AddItem Trim(rs("库房"))
            rs.MoveNext
        Loop

    End If

    Set rs = Nothing
End Sub

Private Sub Query_Lot() '查lot

    Dim strSql    As String

    Dim rs        As New ADODB.Recordset

    Dim i         As Integer, strLotIDTmp As String
    
    Dim StockId    As String
    
    Dim strCustCode As String
    
    Dim strCustLot As String

    ' If Trim(txtCustLot.Text) = "" Then

        ' MsgBox "请填写LotID", vbExclamation, "提示"

    ' Exit Sub
    
    ' End If
    
    If Trim(cbCustCode.Text) = "" Then
    
        MsgBox "请选择客户代码", vbExclamation, "提示"
    
        Exit Sub
        
    End If
    
    
    
    ' If Trim(CbstockId_org.Text) = "" Then
    
        ' MsgBox "请选择仓库", vbExclamation, "提示"
    
        ' Exit Sub
    
    ' End If
    strCustCode = UCase(Trim$(cbCustCode.Text))

    strSql = "SELECT  0 as'选择',a.客户代码,a.工单号,a.库房编号,a.料号,a.库存数,a.合格数,a.不良数,a.制程不良数,a.大工单,a.ID from erpdata..tblStockNum a where a.客户代码='" & strCustCode & "' and isnull(a.库存数,0)>0 "
    
    If Trim$(txtCustLot.Text) <> "" Then
        strCustLot = Trim(txtCustLot.Text)
        strSql = strSql & " and a.工单号='" & strCustLot & "'"
    End If
    
    If Trim(CbstockId_org.Text) <> "" Then
        StockId = Left(Trim(CbstockId_org.Text), InStr(Trim(CbstockId_org.Text), " ") - 1)
        strSql = strSql & " and a.库房编号='" & StockId & "'"
    End If

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        'Call ListDataType(rs)
        With fpS_Lot
            Set .DataSource = Nothing
            .MaxRows = 0
            Set .DataSource = rs
            .Col = E_LOT.E_CHOOSE
            .CellType = CellTypeCheckBox
            .TypeHAlign = TypeVAlignCenter
            .TypeVAlign = TypeVAlignCenter
            '.Lock = False
        
            
        End With
        ' Do While Not rs.EOF
            ' With fps_Lot
            ' .MaxRows = .MaxRows + 1
            ' .SetText E_Lot.E_CHOOSE, .MaxRows, 0
            ' .SetText E_Lot.E_CUSTCODE, .MaxRows, Trim$("" & rs("客户代码").Value)
            ' .SetText E_Lot.E_LOTID, .MaxRows, Trim$("" & rs("工单号").Value)
            ' .SetText E_Lot.E_STOCKID, .MaxRows, Trim("" & rs("库房编号").Value)

            ' .SetText E_Lot.E_PN, .MaxRows, Trim$("" & rs("料号").Value)
            ' .SetText E_Lot.E_TOTALQTY, .MaxRows, Trim$("" & rs("库存数").Value)
            ' .SetText E_Lot.E_PASSQTY, .MaxRows, Trim("" & rs("合格数").Value)
            ' .SetText E_Lot.E_NGQTY1, .MaxRows, Trim$("" & rs("不良数").Value)
            ' .SetText E_Lot.E_NGQTY2, .MaxRows, Trim$("" & rs("制程不良数").Value)
            ' .SetText E_Lot.E_JOBID, .MaxRows, Trim("" & rs("大工单").Value)
            ' .SetText E_Lot.E_ID, .MaxRows, Trim("" & rs("ID").Value)
            
            
    
            ' End With
            ' rs.MoveNext
        ' Loop
    Else
        MsgBox "无数据", vbInformation, "提示"
        Exit Sub

    End If

End Sub

Private Sub cmdQuery_Click()
    Query_Lot
    
End Sub

Private Sub fpS_Lot_Change(ByVal Col As Long, ByVal Row As Long)
Exit Sub
Dim i       As Long
Dim J       As Integer
Dim strID As String

If Col <> 1 Then Exit Sub
If Row < 1 Then Exit Sub

'MsgBox "CHANGE"

With fpS_Lot
    .Row = Row
    .Col = E_WO.E_CHOOSE
    If .Value = 0 Then
        .Value = 1
        .Col = -1
        .ForeColor = &HFF8080
        .Col = E_LOT.E_ID
        strID = Trim$(.Text)
    
      
        Call SearchBoxID_ByID(strID, 1)
    Else
        .Value = 0
        .Col = -1
        .ForeColor = vbBlack
        .Col = E_LOT.E_ID
        strID = Trim$(.Text)
        
        Call SearchBoxID_ByID(strID, 2)

    End If

End With

End Sub




Private Sub fpS_lot_Click(ByVal Col As Long, ByVal Row As Long)
Dim i       As Long
Dim J       As Integer
Dim strID As String

If Col <> 1 Then Exit Sub
'If Row < 1 Then Exit Sub


With fpS_Lot
    .Row = Row
    .Col = E_WO.E_CHOOSE
    .Value = Abs(Val(.Value) - 1)
    If .Value = 1 Then
        .Col = -1
        .ForeColor = &HFF8080
        .Col = E_LOT.E_ID
        strID = Trim$(.Text)
         
        Call SearchBoxID_ByID(strID, 1)
    ElseIf .Value = 0 Then
        .Col = -1
        .ForeColor = vbBlack
        .Col = E_LOT.E_ID
        strID = Trim$(.Text)
        
        Call SearchBoxID_ByID(strID, 2)

    End If

End With
End Sub
Private Sub SearchBoxID_ByID(strID As String, intBJ As Integer)
Dim i      As Long
Dim strSql As String
Dim rs     As New ADODB.Recordset
Dim Lot_temp As String
Dim Stock_temp As String
If intBJ = 1 Then

    With fpS_Box

        For i = 1 To .MaxRows
            .Row = i
            .Col = E_BOX.E_ID

            If strID = Trim$(.Text) Then
                Exit Sub

            End If

        Next

    End With

    '查询资料
    strSql = "select distinct 0 as '√', a.工单号 ,a.库房编号 ,b.箱号,b.ID  from erpdata..tblStockNum a ,erpdata..tblStockNumSub b where a.ID=b.ID and a.Id='" & strID & "'"
    MsgBox strSql
    Set rs = Get_SqlserveRs(strSql)
    
    
    If rs.RecordCount > 0 Then
        With fpS_Box
            For i = 1 To rs.RecordCount
                .MaxRows = .MaxRows + 1
                .SetText E_BOX.E_CHOOSE, .MaxRows, 1
                .SetText E_BOX.E_LOTID, .MaxRows, Trim$("" & rs!工单号)
                .SetText E_BOX.E_STOCKID, .MaxRows, Trim$("" & rs!库房编号)
                .SetText E_BOX.E_BOXID, .MaxRows, Trim$("" & rs!箱号)
                .SetText E_BOX.E_ID, .MaxRows, Trim$("" & rs!ID)
                rs.MoveNext
            Next

        End With

    End If

    rs.Close
    Set rs = Nothing

End If

If intBJ = 2 Then

    With fpS_Box
        Set .DataSource = Nothing

        For i = .MaxRows To 1 Step -1
            .Row = i
            .Col = E_BOX.E_ID
            If strID = Trim$(.Text) Then
                .DeleteRows i, 1
                .MaxRows = .MaxRows - 1

            End If

        Next

    End With

End If



End Sub

Private Sub ListDataType(rs As ADODB.Recordset)
Dim i As Long

With fpS_Lot
    .MaxRows = 0
    Set .DataSource = rs
    
    '.MaxCols = .MaxCols + 1
   '  For i = 1 To .MaxRows
        ' .Row = i
      '   .Col = 1  '选择
      
        '.CellType = CellTypeCheckBox
        '.TypeHAlign = TypeVAlignCenter
       ' .TypeVAlign = TypeVAlignCenter
       ' .Lock = False
   ' Next
    
End With
End Sub



Private Sub Form_Load()
   InitFps
End Sub

Sub InitFps()

    With fpS_Lot
        .MaxCols = E_LOT.E_END - 1
        .Col = -1
        .Row = -1
        .Lock = True
        .SetText 0, 0, "序号"
        .ColWidth(0) = 4
        .SetText E_LOT.E_CHOOSE, 0, "√"
        .ColWidth(E_LOT.E_CHOOSE) = 2
    
        .SetText E_LOT.E_CUSTCODE, 0, "客户"
        .ColWidth(E_LOT.E_CUSTCODE) = 8
        .SetText E_LOT.E_LOTID, 0, "工单号"
        .ColWidth(E_LOT.E_LOTID) = 8
        .SetText E_LOT.E_STOCKID, 0, "库房"
        .ColWidth(E_LOT.E_STOCKID) = 4
        .SetText E_LOT.E_PN, 0, "料号"
        .ColWidth(E_LOT.E_PN) = 14
        .SetText E_LOT.E_TOTALQTY, 0, "库存"
        .ColWidth(E_LOT.E_TOTALQTY) = 4
        .SetText E_LOT.E_PASSQTY, 0, "合格"
        .ColWidth(E_LOT.E_PASSQTY) = 4
        .SetText E_LOT.E_NGQTY1, 0, "不良数"
        .ColWidth(E_LOT.E_NGQTY1) = 4
        .SetText E_LOT.E_NGQTY2, 0, "制程不良数"
        .ColWidth(E_LOT.E_NGQTY2) = 4
        .SetText E_LOT.E_JOBID, 0, "大工单"
        .ColWidth(E_LOT.E_JOBID) = 10
        .SetText E_LOT.E_ID, 0, "ID"
        .ColWidth(E_LOT.E_ID) = 4
        .Col = E_LOT.E_CHOOSE
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        '.Col = E_LOT.E_QTY
        '.BackColor = glColorInProcess
    End With
    With fpS_Box
        .MaxCols = E_BOX.E_END - 1
        .Col = -1
        .Row = -1
        .Lock = True
        .SetText 0, 0, "序号"
        .ColWidth(0) = 4
        .SetText E_BOX.E_CHOOSE, 0, "√"
    
        .ColWidth(E_BOX.E_CHOOSE) = 4
        .SetText E_BOX.E_LOTID, 0, "工单号"
        .ColWidth(E_BOX.E_LOTID) = 10
        .SetText E_BOX.E_STOCKID, 0, "库房"
        .ColWidth(E_BOX.E_STOCKID) = 4
        .SetText E_BOX.E_BOXID, 0, "箱号"
        .ColWidth(E_BOX.E_BOXID) = 10
        .SetText E_BOX.E_ID, 0, "ID"
        .ColWidth(E_BOX.E_ID) = 10
        .Col = E_BOX.E_CHOOSE
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
    End With
    With fpS_WaferId
        .MaxCols = E_WaferId.E_END - 1
        .MaxRows = 0
    End With
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key

        Case "QUERY"
            Toolbar1.Buttons("QUERY").Enabled = False
            Query_Lot
            Toolbar1.Buttons("QUERY").Enabled = True
        Case "EXIT"
            Unload Me

    End Select

End Sub



