VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmPOPriceSys 
   Caption         =   "市场部订单价格维护"
   ClientHeight    =   10710
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
   ScaleHeight     =   10710
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12840
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPOPriceSys.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPOPriceSys.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPOPriceSys.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPOPriceSys.frx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPOPriceSys.frx":3148
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   1402
      ButtonWidth     =   2566
      ButtonHeight    =   1349
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "     查询      "
            Key             =   "SEARCH"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    增加    "
            Key             =   "ADD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    删除    "
            Key             =   "DEL"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    导出    "
            Key             =   "EXPORT"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "     退出     "
            Key             =   "EXIT"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9975
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   16695
      Begin VB.TextBox txtDiePrice 
         Height          =   300
         Left            =   5160
         TabIndex        =   36
         Top             =   2955
         Width           =   1335
      End
      Begin VB.TextBox txtRate 
         Height          =   300
         Left            =   8400
         TabIndex        =   33
         Top             =   1085
         Width           =   1815
      End
      Begin VB.TextBox txtPiecePrice 
         Height          =   300
         Left            =   5160
         TabIndex        =   30
         Top             =   2025
         Width           =   1335
      End
      Begin VB.TextBox txtFileName 
         Height          =   300
         Left            =   8400
         TabIndex        =   28
         Top             =   1540
         Width           =   1815
      End
      Begin VB.TextBox txtBJDH 
         Height          =   300
         Left            =   8400
         TabIndex        =   26
         Top             =   1995
         Width           =   1815
      End
      Begin VB.ComboBox cmbUnit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "FrmPOPriceSys.frx":3D9A
         Left            =   8400
         List            =   "FrmPOPriceSys.frx":3DA4
         TabIndex        =   24
         Top             =   630
         Width           =   1815
      End
      Begin VB.TextBox txtDIES 
         Height          =   300
         Left            =   5160
         TabIndex        =   22
         Top             =   2490
         Width           =   1335
      End
      Begin VB.TextBox txtPieces 
         Height          =   300
         Left            =   5160
         TabIndex        =   19
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cbPOType 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "FrmPOPriceSys.frx":3DB6
         Left            =   4680
         List            =   "FrmPOPriceSys.frx":3DC6
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1095
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DT1 
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   2520
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65280
         CalendarForeColor=   12632064
         CalendarTitleBackColor=   16576
         CalendarTitleForeColor=   12583104
         Format          =   173211649
         CurrentDate     =   43308
      End
      Begin MSDataListLib.DataCombo dcCusPOID 
         Height          =   315
         Left            =   4680
         TabIndex        =   9
         Top             =   623
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   12582912
         Text            =   ""
      End
      Begin VB.TextBox txtCusPN 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   1920
         Width           =   1935
      End
      Begin VB.ComboBox cbCusCode 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   630
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DT2 
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   3000
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12582912
         CalendarForeColor=   49152
         CalendarTitleBackColor=   255
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711935
         Format          =   173211649
         CurrentDate     =   43308
      End
      Begin MSComCtl2.DTPicker DT3 
         Height          =   375
         Left            =   8400
         TabIndex        =   35
         Top             =   2453
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12582912
         CalendarForeColor=   49152
         CalendarTitleBackColor=   255
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711935
         Format          =   173211649
         CurrentDate     =   43308
      End
      Begin VB.Line Line7 
         X1              =   3720
         X2              =   6600
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line6 
         BorderStyle     =   3  'Dot
         X1              =   3840
         X2              =   6600
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line5 
         BorderStyle     =   3  'Dot
         X1              =   3840
         X2              =   6600
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line3 
         X1              =   6600
         X2              =   6600
         Y1              =   480
         Y2              =   3360
      End
      Begin VB.Line Line2 
         X1              =   6600
         X2              =   3720
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         X1              =   3720
         X2              =   3720
         Y1              =   480
         Y2              =   3360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提供订单日期"
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
         Index           =   11
         Left            =   7140
         TabIndex        =   34
         Top             =   2535
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "     汇  率"
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
         Index           =   16
         Left            =   7125
         TabIndex        =   32
         Top             =   1140
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "订单价格(DIE)"
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
         Index           =   15
         Left            =   3840
         TabIndex        =   31
         Top             =   3000
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "订单价格(片)"
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
         Index           =   14
         Left            =   3840
         TabIndex        =   29
         Top             =   2070
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "返回文件名"
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
         Index           =   13
         Left            =   7260
         TabIndex        =   27
         Top             =   1605
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报价单号"
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
         Index           =   12
         Left            =   7500
         TabIndex        =   25
         Top             =   2070
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单价单位"
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
         Index           =   10
         Left            =   7560
         TabIndex        =   23
         Top             =   675
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "订单数量(DIE)"
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
         Index           =   9
         Left            =   3840
         TabIndex        =   21
         Top             =   2535
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   8
         Left            =   720
         TabIndex        =   20
         Top             =   3075
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "订单数量(片)"
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
         Index           =   7
         Left            =   3840
         TabIndex        =   18
         Top             =   1605
         Width           =   1260
      End
      Begin MSForms.TextBox txtCusSName 
         Height          =   300
         Left            =   960
         TabIndex        =   17
         Top             =   1035
         Width           =   1935
         VariousPropertyBits=   -1400879081
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "3413;529"
         SpecialEffect   =   0
         FontName        =   "宋体"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户简写"
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
         TabIndex        =   16
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "订单类型"
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
         Left            =   3840
         TabIndex        =   13
         Top             =   1140
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   2595
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "订单号码"
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
         Index           =   3
         Left            =   3840
         TabIndex        =   8
         Top             =   675
         Width           =   840
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   840
      End
      Begin MSForms.TextBox txtCusName 
         Height          =   495
         Left            =   960
         TabIndex        =   5
         Top             =   1365
         Width           =   1935
         VariousPropertyBits=   -1400879081
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "3413;873"
         SpecialEffect   =   0
         FontName        =   "宋体"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户名称"
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
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   675
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmPOPriceSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbCusCode_Change()
  ListCusName
ListCusPOID
End Sub

Private Sub cbCusCode_Click()
    ListCusName
    ListCusPOID

End Sub

Private Sub ListCusName()

    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    rs.Source = "select distinct 客户名称 from tblxcustomer where 客户代码='" & Trim(cbCusCode.Text) & "'"
    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        txtCusName.Text = Trim(rs!客户名称)
    End If
    
    txtCusSName.Text = GetCustomerNameSqlServer1(Trim(cbCusCode.Text))
    
    If GetCustomerNameSqlServer2(Trim(cbCusCode.Text)) = "01" Then
        cmbUnit.Text = "人名币"
    Else
        cmbUnit.Text = "美元"
    End If

End Sub

Private Sub ListCusPOID()

    Dim rs As ADODB.Recordset

    Set rs = GetCustomerPONum(Trim(cbCusCode.Text))
    Set dcCusPOID.RowSource = rs
    dcCusPOID.ListField = rs("productname").Name
    dcCusPOID.BoundColumn = rs("PID").Name

End Sub

Private Sub InitCuscode()

    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    rs.Source = "select distinct 客户代码 from tblxcustomer"
    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    cbCusCode.Clear

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        For I = 1 To rs.RecordCount
            cbCusCode.AddItem Trim(rs("客户代码"))
            rs.MoveNext
        Next I

    End If
    
End Sub

Private Sub cbCusCode_LostFocus()

    Dim rs As ADODB.Recordset

    If txtCusName.Text <> "" Then

        Exit Sub

    End If

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    rs.Source = "select distinct 客户名称 from tblxcustomer where 客户代码='" & Trim(cbCusCode.Text) & "'"
    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        txtCusName.Text = Trim(rs!客户名称)
    End If

End Sub

Private Sub dcCusPOID_Change()

    Dim potemp   As String

    Dim custTemp As String

    potemp = UCase(Trim(dcCusPOID.Text))
    custTemp = UCase(Trim(cbCusCode.Text))

    Set oiRS = GetOIDataPONum(custTemp, potemp)

    If (oiRS.RecordCount > 0) Then

        DT3.Value = oiRS.Fields("qtech_created_date").Value
        txtCusPN.Text = oiRS.Fields("mpn_desc").Value
        txtDies.Text = oiRS.Fields("qty").Value
        txtPieces.Text = oiRS.Fields("qty2").Value
        
    End If
     
End Sub

Private Sub Form_Load()
    InitCuscode
    InitDate
    InitListView
End Sub

Private Sub InitDate()

    DT1.Value = Now - 1
    DT2.Value = Now
    DT3.Value = Now

End Sub

Private Sub InitListView()

    On Error GoTo EXITPRO

    Dim Clm As ColumnHeader

    ListView1.ColumnHeaders.Clear
    ListView1.View = lvwReport
    ListView1.LabelEdit = lvwAutomatic
    
    Set Clm = ListView1.ColumnHeaders.Add(, , "", 500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "记录号", 1000)
    Set Clm = ListView1.ColumnHeaders.Add(, , "客户代码", 1000)
    Set Clm = ListView1.ColumnHeaders.Add(, , "客户全称", 1500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "订单号", 1500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "提供订单日期", 1500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "录入日期", 1500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "订单类型", 1200)
    Set Clm = ListView1.ColumnHeaders.Add(, , "机种", 1500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "订单片数", 1500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "汇率", 1000)
    Set Clm = ListView1.ColumnHeaders.Add(, , "订单数量", 1500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "单片价格", 1500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "单DIE价格", 1500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "单价单位", 1500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "返点文件名", 1000)
    Set Clm = ListView1.ColumnHeaders.Add(, , "客户简写", 1500)
    Set Clm = ListView1.ColumnHeaders.Add(, , "报价单号", 1500)
 
EXITSUB:

    Exit Sub

EXITPRO:

    On Error GoTo EXITSUB

    Resume Next

End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key

        Case "SEARCH"
            ForSearch

        Case "EXIT"
            ForExit

        Case "EXPORT"
            ForExport

        Case "ADD"
            ForAdd

        Case "DEL"
            ForDel
    End Select

End Sub

Private Sub ForDel()

    Dim I    As Integer

    Dim id   As String

    Dim PO   As String

    Dim pt   As String
    
    Dim bRtn As Boolean

    bRtn = False
    
    For I = 1 To ListView1.ListItems.Count

        If ListView1.ListItems(I).Checked Then
            bRtn = True

        End If

    Next I
    
    If bRtn = False Then
        MsgBox "请勾选要删除的条目", vbCritical, "警告"

        Exit Sub

    End If

    For I = 1 To ListView1.ListItems.Count
        
        If ListView1.ListItems(I).Checked Then
            id = ListView1.ListItems(I).SubItems(1)
            PO = ListView1.ListItems(I).SubItems(4)
            pt = ListView1.ListItems(I).SubItems(8)

            If MsgBox("是否删除", vbYesNoCancel, "提醒") = vbYes Then
                Call DelPoPrice(id, PO, pt)
            End If

            Exit For

        End If

    Next I
    
    ForSearch
End Sub

Private Sub DelPoPrice(no As String, PO As String, pt As String)

    Dim strSql As String

    strSql = "delete from TSV_MD_POPrice where id = '" & no & "'"
    AddSql (strSql)

    strSql = "delete from erptemp..tblBB_CSRPO where PO_NUM = '" & PO & "' and FAB_DEVICE = '" & pt & "'"
    AddSql2 (strSql)

    MsgBox "PO: " & PO & " 机种: " & pt & " ,删除成功", vbInformation, "提示"

End Sub

Private Sub ForAdd()

    Dim nPOTemp As POPrice

    Dim rs      As ADODB.Recordset

    If UCase(Trim(cbCusCode.Text)) = "" Or UCase(Trim(txtCusName.Text)) = "" Then
        MsgBox "客户代码或客户名称不可以为空！", vbExclamation, "警告"

        Exit Sub

    End If

    If Trim(txtPiecePrice.Text) = "" And Trim(txtDiePrice.Text) = "" Then
        MsgBox "单价不可以为空!", vbExclamation, "警告"

        Exit Sub

    End If

    If Trim(txtDies.Text) = "" Or Trim(cmbUnit.Text) = "" Then
        MsgBox "数量,单位 不可以为空！"

        Exit Sub

    End If

    If UCase(Trim(cbCusCode.Text)) = "KR001" Or UCase(Trim(cbCusCode.Text)) = "81" Then
        If UCase(Trim(txtBJDH.Text)) = "" Then
            MsgBox "报价单号不能为空！"

            Exit Sub
 
        End If
    End If

    nPOTemp.CreateBy = UCase(gUserName)
    nPOTemp.id = GetPOPriceID()
    nPOTemp.customerName = UCase(Trim(txtCusName.Text))
    nPOTemp.CUSTOMERSHORTNAME = UCase(Trim(cbCusCode.Text))
    nPOTemp.PODATE = Format(DT3.Value, "YYYY-MM-DD")
    nPOTemp.PONo = UCase(Trim(dcCusPOID.Text))
    nPOTemp.POType = UCase(Trim(cbPOType.Text))
    nPOTemp.pt = UCase(Trim(txtCusPN.Text))
    nPOTemp.qty = UCase(Trim(txtDies.Text))
    nPOTemp.SingDie = Trim$(txtDiePrice.Text)
    nPOTemp.peaseQty = UCase(Trim(txtPieces.Text))
    nPOTemp.Price = UCase(Trim(txtPiecePrice.Text))
    nPOTemp.unit = UCase(Trim(cmbUnit.Text))
    nPOTemp.File = UCase(Trim(txtFileName.Text))
    nPOTemp.bj = UCase(Trim(txtBJDH.Text))
    nPOTemp.custAA = UCase(Trim(txtCusSName))
    nPOTemp.SingWafer = Trim$(txtRate.Text)

    Set rs = GetPJPOData(UCase(Trim(dcCusPOID.Text)), UCase(Trim(txtCusPN.Text)))

    If rs.RecordCount > 0 Then
        MsgBox "Mes系统中已存在此采购单单号，请确认采购单 ！"

        Exit Sub

    End If

    Call AddPOPrice(nPOTemp)

    MsgBox "新增成功!", vbInformation, "提示"
 
    ForSearch

End Sub

Private Sub ForExport()

    If Me.ListView1.ListItems.Count = 0 Then
        MsgBox "没有数据", vbExclamation, "提示"

        Exit Sub

    End If
    
    On Error Resume Next

    Dim xlApp As Object, Wb As Object, I&, j&, T&, h&, ar()

    T = ListView1.ListItems.Count
    h = ListView1.ColumnHeaders.Count

    If T = 0 Then Exit Sub
    DoEvents
    Set xlApp = CreateObject("Excel.Application")

    If xlApp Is Nothing Then
        MsgBox "pls install Microsoft Excel."

        Exit Sub

    Else

        ReDim ar(1 To T + 1, 1 To (h + 2))

        For I = 1 To h
            ar(1, I) = ListView1.ColumnHeaders(I)
        Next

        For I = 2 To T + 1

            If ListView1.ListItems(I - 1).Checked Then
                ar(I, 1) = ListView1.ListItems(I - 1)

                For j = 1 To h
                    ar(I, j + 1) = ListView1.ListItems(I - 1).SubItems(j)
                Next
                
            End If

        Next
        
        Set Wb = xlApp.Workbooks.Add
        Wb.ActiveSheet.Range("A:AA").NumberFormatLocal = "@"
        Wb.ActiveSheet.Range("a1").Resize(UBound(ar), h + 1) = ar
        Wb.ActiveSheet.Cells.Columns.AutoFit
        xlApp.Visible = True
        Set Wb = Nothing
        Set xlApp = Nothing

    End If

End Sub

Private Sub ForSearch()
   
    Dim rs        As ADODB.Recordset

    Dim strSql    As String

    Dim date1Temp As String

    Dim date2Temp As String

    Dim I         As Long

    Dim ITEM      As ListItem

    If Trim(dcCusPOID.Text) <> "" Then
        strSql = "select ID ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE ,qtech_created_date, PO_TYPE ,PT , PeaceQty, rate, QTY ,PRICE ,DIE_PRICE,UNIT, FILENAME,CUSTAA,BJ from TSV_MD_POPrice where flag='Y' and PO_NUM ='" & UCase$(Trim(dcCusPOID.Text)) & "'order by id desc "
        Set rs = GetAAMPNData(strSql)
    
    Else
        date1Temp = Format(DT1.Value, "YYYY-MM-DD")
        date2Temp = Format(DT2.Value + 1, "YYYY-MM-DD")
 
        strSql = " select   ID ,CUSTOMERSHORTNAME,CUSTOMERNAME ,PO_NUM ,PO_DATE ,qtech_created_date, PO_TYPE ,PT ,PeaceQty, rate,QTY ,PRICE ,DIE_PRICE,UNIT , " & "  FILENAME,CUSTAA,BJ   from  TSV_MD_POPrice where flag='Y'  and  PO_DATE >=to_date('" + date1Temp + "','YYYY-MM-DD')  and  PO_DATE <to_date('" + date2Temp + "','YYYY-MM-DD')  "
 
        If cbCusCode.Text <> "" Then

            strSql = strSql & " and  CUSTOMERSHORTNAME = '" + UCase(Trim(cbCusCode.Text)) + "' "

        End If

        If cbPOType.Text <> "" Then

            strSql = strSql & " and  PO_TYPE = '" + UCase(Trim(cbPOType.Text)) + "' "

        End If

        If txtCusPN.Text <> "" Then

            strSql = strSql & " and  PT = '" + UCase(Trim(txtCusPN.Text)) + "' "

        End If

        Set rs = GetAAMPNData(strSql)
 
    End If
    
    ListView1.ListItems.Clear

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        For I = 1 To rs.RecordCount
            Set ITEM = ListView1.ListItems.Add(, , Str(I))
            ITEM.SubItems(1) = GetRsData2(rs, "ID")
            ITEM.SubItems(2) = GetRsData2(rs, "CUSTOMERSHORTNAME")
            ITEM.SubItems(3) = GetRsData2(rs, "CUSTOMERNAME")
            ITEM.SubItems(4) = GetRsData2(rs, "PO_NUM")
            ITEM.SubItems(5) = GetRsData2(rs, "PO_DATE")
            ITEM.SubItems(6) = GetRsData2(rs, "qtech_created_date")
            ITEM.SubItems(7) = GetRsData2(rs, "PO_TYPE")
            ITEM.SubItems(8) = GetRsData2(rs, "PT")
            ITEM.SubItems(9) = GetRsData2(rs, "PeaceQty")
            ITEM.SubItems(10) = GetRsData2(rs, "rate")
            ITEM.SubItems(11) = GetRsData2(rs, "QTY")
            ITEM.SubItems(12) = GetRsData2(rs, "PRICE")
            ITEM.SubItems(13) = GetRsData2(rs, "DIE_PRICE")
            ITEM.SubItems(14) = GetRsData2(rs, "UNIT")
            ITEM.SubItems(15) = GetRsData2(rs, "FILENAME")
            ITEM.SubItems(16) = GetRsData2(rs, "CUSTAA")
            ITEM.SubItems(17) = GetRsData2(rs, "BJ")
            
            rs.MoveNext
        Next
        
    Else
        MsgBox "没有查询到数据!", vbInformation, "提示"
        Screen.MousePointer = 0
        
        Exit Sub
        
    End If

    With ListView1

        For I = 1 To .ListItems.Count
            .ListItems(I).Checked = True
                
        Next

    End With

End Sub

Private Sub ForExit()

    Unload Me
End Sub

Private Function GetRsData2(rs As ADODB.Recordset, data As String) As String

    If IsNull(rs(data)) = True Then
        GetRsData2 = ""
    Else
        GetRsData2 = Trim$(rs(data))
    End If

End Function
