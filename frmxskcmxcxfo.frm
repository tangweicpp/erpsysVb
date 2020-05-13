VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmxskcmxcxfo 
   Caption         =   "FO库房物料明细"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13710
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   13710
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   3255
      Begin MSComCtl2.DTPicker dtEndTime 
         Height          =   330
         Left            =   960
         TabIndex        =   24
         Top             =   4440
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   275513345
         CurrentDate     =   42801
      End
      Begin MSComCtl2.DTPicker dtStartTime 
         Height          =   330
         Left            =   960
         TabIndex        =   23
         Top             =   3960
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   275513345
         CurrentDate     =   42801
      End
      Begin VB.TextBox txtJob 
         Height          =   330
         Left            =   960
         TabIndex        =   22
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtBbox 
         Height          =   330
         Left            =   960
         TabIndex        =   19
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox txtSbox 
         Height          =   330
         Left            =   960
         TabIndex        =   17
         Top             =   2520
         Width           =   2175
      End
      Begin VB.ComboBox Combo5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   1  'Simple Combo
         TabIndex        =   13
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox Combo4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   1  'Simple Combo
         TabIndex        =   11
         Top             =   1605
         Width           =   2175
      End
      Begin VB.ComboBox Cmbcust 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   8
         Top             =   693
         Width           =   2175
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   7
         Top             =   1146
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "先选择库房，然后选择时间段..."
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox comBo2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   2
         Top             =   4975
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComctlLib.TreeView TreeView3 
         Height          =   2055
         Left            =   120
         TabIndex        =   15
         Top             =   5500
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3625
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblInTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入库时间"
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
         Left            =   120
         TabIndex        =   25
         Top             =   3960
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblJob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job  No"
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
         Left            =   120
         TabIndex        =   21
         Top             =   3380
         Width           =   735
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "大 箱 号"
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
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   2955
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "小 箱 号"
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
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   2600
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "仓    位"
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
         Left            =   120
         TabIndex        =   14
         Top             =   2100
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "料    号"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1659
         Width           =   840
      End
      Begin MSForms.Label Label3 
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   753
         Width           =   855
         ForeColor       =   0
         VariousPropertyBits=   276824091
         Caption         =   "客户代码"
         Size            =   "1508;370"
         FontName        =   "宋体"
         FontHeight      =   210
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label7 
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   750
         ForeColor       =   0
         VariousPropertyBits=   276824091
         Caption         =   "工单LOT"
         Size            =   "1323;370"
         FontName        =   "宋体"
         FontHeight      =   210
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房名称"
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
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "物料类型"
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
         Left            =   120
         TabIndex        =   4
         Top             =   5035
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   1535
      ButtonWidth     =   1455
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "  打  印   "
            Key             =   "A01"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  输  出  "
            Key             =   "A02"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   " 调  拨 "
            Key             =   "A03"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "  删  除"
            Key             =   "A04"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "修  改"
            Key             =   "A05"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  查  询"
            Key             =   "A06"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "  修  改"
            Key             =   "A07"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " 取  消"
            Key             =   "A08"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "  确  认"
            Key             =   "A09"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "A004"
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "  帮  助"
            Key             =   "A10"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  退  出"
            Key             =   "A11"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmxskcmxcxfo.frx":0000
      Begin VB.CheckBox chk 
         Caption         =   "全选/反选"
         Height          =   180
         Left            =   4320
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9600
         Top             =   0
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
               Picture         =   "frmxskcmxcxfo.frx":0162
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":229C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":5126
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":78D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":9A12
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":C1C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":E976
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":119F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":141AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":144C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":1519E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":18220
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmxskcmxcxfo.frx":1A9D2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog ee 
         Left            =   10200
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Excel导入"
         Filter          =   "*xls"
         InitDir         =   "D:\"
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7335
      Left            =   3360
      TabIndex        =   6
      ToolTipText     =   "双击进行仓位维护"
      Top             =   840
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12938
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmxskcmxcxfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adorst1         As New adodb.Recordset
Public strXHS          As String

Private Sub chk_Click()
Dim I           As Integer

    If ListView1.ListItems.Count < 1 Then Exit Sub
    If chk.Value = 1 Then   '全选
        For I = 1 To ListView1.ListItems.Count
            ListView1.ListItems(I).Checked = True
        Next I
    Else        '反选
        For I = 1 To ListView1.ListItems.Count
            ListView1.ListItems(I).Checked = False
        Next I
    End If
End Sub

Private Sub Cmbcust_DropDown()
Dim I As Integer
    Set adorst2 = New adodb.Recordset
    Set adorst2.ActiveConnection = INIadoCon2
    adorst2.Source = "select distinct 客户代码  from tblXCustomer "
    adorst2.Open , , , , adCmdText
    Cmbcust.Clear
    If adorst2.RecordCount > 0 Then
      For I = 1 To adorst2.RecordCount
        Cmbcust.AddItem Trim(adorst2("客户代码"))
        adorst2.MoveNext
      Next I
    Else
       Cmbcust.Text = ""
       Exit Sub
    End If
'    Combo5.Text = "不限"
End Sub

'Private Sub Combo1_Click()
'Combo5.Text = "不限"
'End Sub

Private Sub Combo2_Click()
  Dim adorst11 As adodb.Recordset
  Dim I As Integer
  Me.MousePointer = 11
  Set adorst11 = New adodb.Recordset
  adorst11.ActiveConnection = INIadoCon2
  adorst11.Source = "SELECT 层级,类型简码,物料类型,结构编码,上级编码  from tblSMTp1   WHERE (层级<  6) and (类型简码 like '01'+'%' ) order by 层级,类型简码"
  adorst11.Open , , adOpenStatic, adLockReadOnly, adCmdText
  TreeView3.Nodes.Clear
     If adorst11.RecordCount > 0 Then
        TreeView3.Top = comBo2.Top
        TreeView3.Left = Label8.Left
        TreeView3.Visible = True
        adorst11.MoveFirst
        For I = 0 To adorst11.RecordCount - 1
            If adorst11("层级") = 1 Then
               Set mNod = TreeView3.Nodes.Add(, , "K" + Trim(Str(adorst11("结构编码"))), Trim(adorst11("物料类型")))
            Else
               Set mNod = TreeView3.Nodes.Add("K" + Trim(Str(adorst11("上级编码"))), 4, "K" + Trim(Str(adorst11("结构编码"))), Trim(adorst11("类型简码")) & Space(1) & Trim(adorst11("物料类型")))
            End If
            adorst11.MoveNext
        Next I
     End If
  Me.MousePointer = 0
  adorst11.Close
  Set adorst11 = Nothing
End Sub

Private Sub Combo3_DropDown()
Dim intNext As Integer
   Set adorst1 = New adodb.Recordset
   adorst1.ActiveConnection = INIadoCon2
   adorst1.Source = "select distinct a.工单号 from tblStockNum a where a.库存数>0  " & _
   " and a.库房编号='" & Left(Combo1.Text, 2) & "'"
   adorst1.Open , , , , adCmdText
   If adorst1.RecordCount > 0 Then
      Combo3.Clear
      adorst1.MoveFirst
      For intNext = 1 To adorst1.RecordCount
          Combo3.AddItem Trim(adorst1("工单号"))
          adorst1.MoveNext
      Next intNext
   Else
   End If
  adorst1.Close
  Set adorst1 = Nothing
'  Combo5.Text = "不限"
End Sub

'Private Sub Combo4_Change()
'Combo5.Text = "不限"
'End Sub

Private Sub Form_Load()
  Set adorst11 = New adodb.Recordset
  adorst11.ActiveConnection = INIadoCon2
  adorst11.Source = "SELECT 库房代码+' '+库房名称 仓库名称 FROM erpbase..tblstock WHERE 仓库属性='成品仓'"
  adorst11.Open , , , , adCmdText
  Combo1.AddItem "不限"
  If adorst11.RecordCount > 0 Then
    adorst11.MoveFirst
    For intSubN = 1 To adorst11.RecordCount
      Combo1.AddItem Trim(adorst11.Fields(0))
    adorst11.MoveNext
    Next intSubN
  End If
  adorst11.Close
  Set adorst11 = Nothing
  Combo1.ListIndex = 0
  Combo1.Text = "不限"
'  comBo2.Text = "不限"
'  Combo5.Text = "不限"
  dtStartTime.Value = Format(Now, "yyyy-mm-dd")
  dtEndTime.Value = Format(DateAdd("d", 1, Now), "yyyy-mm-dd")
  
  Call listitem1
End Sub
Sub listitem1() '邦定ListView1表头
On Error GoTo EXITPRO
  Dim Clm As ColumnHeader
  ListView1.ListItems.Clear
  ListView1.ColumnHeaders.Clear
  ListView1.View = lvwReport
  ListView1.LabelEdit = lvwManual '
  Set Clm = ListView1.ColumnHeaders.Add(, , " ", 300)
  Set Clm = ListView1.ColumnHeaders.Add(, , "客户代码", 1800)
  Set Clm = ListView1.ColumnHeaders.Add(, , "库房编号", 1400)
  Set Clm = ListView1.ColumnHeaders.Add(, , "库房名称", 1400)
  Set Clm = ListView1.ColumnHeaders.Add(, , "工单_LOT", 1800)
  Set Clm = ListView1.ColumnHeaders.Add(, , "料号", 1600)
  Set Clm = ListView1.ColumnHeaders.Add(, , "物料名称", 1400)
  Set Clm = ListView1.ColumnHeaders.Add(, , "规格", 2000)
  Set Clm = ListView1.ColumnHeaders.Add(, , "型号", 1800)
  Set Clm = ListView1.ColumnHeaders.Add(, , "Wafer ID", 1800)
  Set Clm = ListView1.ColumnHeaders.Add(, , "小箱号", 1200)
'  Set Clm = ListView1.ColumnHeaders.Add(, , "大箱号", 1200)
'  Set Clm = ListView1.ColumnHeaders.Add(, , "合格数", 1600)
'  Set Clm = ListView1.ColumnHeaders.Add(, , "制程不良数", 1200)
'  Set Clm = ListView1.ColumnHeaders.Add(, , "来料不良数", 1400)
   Set Clm = ListView1.ColumnHeaders.Add(, , "单位", 1400)
'  Set Clm = ListView1.ColumnHeaders.Add(, , "仓位/货架号", 1400)
'  Set Clm = ListView1.ColumnHeaders.Add(, , "COMMENT", 1400)
'  Set Clm = ListView1.ColumnHeaders.Add(, , "REMARK", 1400)
'  Set Clm = ListView1.ColumnHeaders.Add(, , "CUSTOMERLOTID", 1400)
  Set Clm = ListView1.ColumnHeaders.Add(, , "入库时间", 1400)
'  Set Clm = ListView1.ColumnHeaders.ADD(, , "id", 0)
EXITSUB:
  Exit Sub
EXITPRO:
  On Error GoTo EXITSUB
  Resume Next
End Sub
Sub ListData()
Dim intnum As Integer
Dim itm As ListItem
Dim intSubTotal1 As Long
Dim intSubTotal2 As Long
Dim DoubleSubTotal As Double
Dim intIndex As Integer
Dim SingDQL As Double
Dim singJE As Double
  intSubTotal1 = 0
  intSubTotal2 = 0
  DoubleSubTotal = 0
  Me.MousePointer = 11

  Set adoCmd = New adodb.Command
  adoCmd.ActiveConnection = INIadoCon2
  adoCmd.CommandText = "uspCP_kcmxcx_BJ178"

  
  adoCmd.CommandType = adCmdStoredProc
  Set adoprm1 = New adodb.Parameter
  adoprm1.Type = adVarChar
  adoprm1.Size = 50
  adoprm1.Direction = adParamInput
  If InStr(Trim(Combo1.Text), " ") > 0 Then
  adoprm1.Value = Trim(Left(Trim(Combo1.Text), InStr(Trim(Combo1.Text), " ") - 1)) '库房编号
  Else
  adoprm1.Value = Trim(Combo1.Text)
  End If
  adoCmd.Parameters.Append adoprm1
  
  Set adoprm2 = New adodb.Parameter
  adoprm2.Type = adVarChar
  adoprm2.Size = 50
  adoprm2.Direction = adParamInput
  adoprm2.Value = Trim(Cmbcust.Text)  '客户代码
  adoCmd.Parameters.Append adoprm2
  
  Set adoPrm3 = New adodb.Parameter
  adoPrm3.Type = adVarChar
  adoPrm3.Size = 50
  adoPrm3.Direction = adParamInput
  adoPrm3.Value = Trim(Combo3.Text) '工单号
  adoCmd.Parameters.Append adoPrm3
  
  Set adoPrm4 = New adodb.Parameter
  adoPrm4.Type = adVarChar
  adoPrm4.Size = 50
  adoPrm4.Direction = adParamInput
  adoPrm4.Value = Trim(Combo4.Text) '料号
  adoCmd.Parameters.Append adoPrm4
  
  Set adoPrm5 = New adodb.Parameter
  adoPrm5.Type = adVarChar
  adoPrm5.Size = 50
  adoPrm5.Direction = adParamInput
  adoPrm5.Value = Trim(Combo5.Text) '仓位
  adoCmd.Parameters.Append adoPrm5
  
  Set adoPrm6 = New adodb.Parameter
  adoPrm6.Type = adVarChar
  adoPrm6.Size = 50
  adoPrm6.Direction = adParamInput
  adoPrm6.Value = Trim(txtSbox.Text) '小箱号
  adoCmd.Parameters.Append adoPrm6
  
  Set adoPrm7 = New adodb.Parameter
  adoPrm7.Type = adVarChar
  adoPrm7.Size = 50
  adoPrm7.Direction = adParamInput
  adoPrm7.Value = Trim(txtBbox.Text) '大箱号
  adoCmd.Parameters.Append adoPrm7
  
  Set adoPrm8 = New adodb.Parameter
  adoPrm8.Type = adVarChar
  adoPrm8.Size = 50
  adoPrm8.Direction = adParamInput
  adoPrm8.Value = Trim(txtJob.Text) 'JOBNO
  adoCmd.Parameters.Append adoPrm8
  
  Set adoPrm9 = New adodb.Parameter
  adoPrm9.Type = adVarChar
  adoPrm9.Size = 10
  adoPrm9.Direction = adParamInput
  adoPrm9.Value = Left(Trim(Format(dtStartTime.Value, "yyyy-mm-dd")), 10) '入库查询开始时间
  adoCmd.Parameters.Append adoPrm9
  
  
  Set adoprm10 = New adodb.Parameter
  adoprm10.Type = adVarChar
  adoprm10.Size = 10
  adoprm10.Direction = adParamInput
  adoprm10.Value = Left(Trim(Format(dtEndTime.Value, "yyyy-mm-dd")), 10) '入库查询结束时间
  adoCmd.Parameters.Append adoprm10
  
  
  Set adorst1 = adoCmd.Execute()
  
  If adorst1.RecordCount > 0 Then
      ListView1.ListItems.Clear
     For intIndex = 1 To adorst1.RecordCount
        Set itm = ListView1.ListItems.Add(, , " ")
        itm.SubItems(1) = Trim(adorst1("客户代码"))
        itm.SubItems(2) = Trim(adorst1("库房编号"))
        itm.SubItems(3) = Trim(adorst1("库房名称"))
        itm.SubItems(4) = Trim(adorst1("工单号"))
        itm.SubItems(5) = Trim(adorst1("料号"))
        itm.SubItems(6) = Trim(adorst1("物料名称"))
        itm.SubItems(7) = Trim(adorst1("规格型号"))
        itm.SubItems(8) = Trim(adorst1("型号"))
        itm.SubItems(9) = Trim(adorst1("WaferID"))

        itm.SubItems(10) = Trim(adorst1("小箱号"))
'        itm.SubItems(10 + i) = Trim("" & adorst1("大箱号"))
'        'mwl 2013.7.23 add 增加小箱大箱箱号相同就把大箱号列改为空
'        If Trim(adorst1("小箱号")) = Trim("" & adorst1("大箱号")) Then
'            itm.SubItems(10 + i) = ""
'        End If
'        itm.SubItems(11 + i) = Trim(adorst1("合格数"))
'        itm.SubItems(12 + i) = Trim(adorst1("制程不良数"))
'        itm.SubItems(13 + i) = Trim(adorst1("不良数"))
         itm.SubItems(11) = Trim(adorst1("计量单位名称"))
'        itm.SubItems(15 + i) = Trim("" & adorst1("仓位"))
'        itm.SubItems(16 + i) = Trim("" & adorst1("Comment"))
'        itm.SubItems(17 + i) = Trim("" & adorst1("Remark"))
'        itm.SubItems(18 + i) = Trim("" & adorst1("CUSTOMERLOTID"))
         itm.SubItems(12) = Trim("" & adorst1("入库时间"))
    '    itm.SubItems(15) = Trim(adoRst1("ID"))
        adorst1.MoveNext
     Next intIndex
     chk.Value = 0 '赋值到初始
 End If
     Me.MousePointer = 0
End Sub

Private Sub Form_Resize()
  On Error Resume Next
    ListView1.Width = Me.Width - Frame1.Width - 200
    ListView1.Height = Me.Height - 1500
    Frame1.Height = ListView1.Height + 70
End Sub


'Private Sub ListView1_DblClick()
'Dim intCount            As Integer
'Dim I                   As Integer
'
'intCount = 0
'strXHS = ""
''If FormLoad(strUserNum, "CP18") = True Then '库房调拨及仓位选择
'  If ListView1.ListItems.Count < 1 Then Exit Sub
'    For I = 1 To ListView1.ListItems.Count
'       If ListView1.ListItems(I).Checked = True Then
'            intCount = intCount + 1
'            strXHS = strXHS & Trim(ListView1.ListItems(I).ListSubItems(9).Text) & "★"
'       End If
'    Next I
'    If intCount <= 0 Then
'        MsgBox ("请先勾选要修改的箱号！")
'        Exit Sub
'    End If
'    '-------------------------------------------------------------
'    chkfbh = Left(Combo1.Text, InStr(Combo1.Text, " ") - 1) '库房编号
'    strPublicPrm = 5 '直接修改仓位
'    frmxscwsd.Show vbModal
''  End If
'End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
       Case "A02"
            If ListView1.ListItems.Count <= 0 Then Exit Sub
            If Trim(ListView1.ListItems(ListView1.selectedItem.Index).ListSubItems(1).Text) = "" Then Exit Sub
            Call RsExporToExcel(adorst1)
            Screen.MousePointer = 0
       Case "A08"
            Unload Me
       Case "A06"
            Call listitem1
            Call ListData
       Case "A11"
            Unload Me
  End Select
End Sub

'根据Rs数据集语句导出Excel
Public Sub RsExporToExcel(rs As adodb.Recordset)
Dim Irowcount       As Long
Dim Icolcount       As Integer

    Dim xlApp As New EXCEL.Application
    Dim xlBook As EXCEL.Workbook
    Dim xlSheet As EXCEL.Worksheet
    Dim xlQuery As EXCEL.QueryTable
    Screen.MousePointer = 11
    With rs
        If .RecordCount < 1 Then
            Screen.MousePointer = 0
            MsgBox ("没有可导出的资料")
            Exit Sub
        End If
        Irowcount = .RecordCount
        Icolcount = .Fields.Count
    End With
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlApp.Visible = True

    Set xlQuery = xlSheet.QueryTables.Add(rs, xlSheet.Range("a1"))
    
'    With xlQuery
'        .FieldNames = True
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = True
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = True
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .PreserveColumnInfo = True
'    End With
'    xlSheet.Name = Rptname
    xlQuery.FieldNames = True '
    xlQuery.Refresh
    
    With xlSheet
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Name = "宋体"
        '
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = True
        '
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
        '
'        .Range(.Cells(2, 1), .Cells(Irowcount + 1, Icolcount)).Font.Size = 9
    End With

    
    xlApp.Visible = True
    Set xlApp = Nothing  '
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Sub


Private Sub TreeView3_Click()
  If TreeView3.selectedItem.Parent Is Nothing Then Exit Sub
  comBo2.Text = Trim(TreeView3.selectedItem.Text)
  TreeView3.Visible = False
End Sub

