VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_CGVH 
   Caption         =   "采购维护"
   ClientHeight    =   11595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15195
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11595
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   10935
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   15255
      Begin VB.CommandButton Command7 
         Caption         =   "还原"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "作废"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "还原"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6000
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "作废"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5160
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "更改"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9840
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin FPSpreadADO.fpSpread fpss 
         Height          =   3855
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   6960
         Width           =   15255
         _Version        =   524288
         _ExtentX        =   26908
         _ExtentY        =   6800
         _StockProps     =   64
         DAutoCellTypes  =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   0
         MaxRows         =   0
         SpreadDesigner  =   "For4.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   6000
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "确认"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9000
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "For4.frx":0404
         Left            =   1200
         List            =   "For4.frx":040E
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   6000
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   2895
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   4455
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   1800
         Width           =   15255
         _Version        =   524288
         _ExtentX        =   26908
         _ExtentY        =   7858
         _StockProps     =   64
         DAutoCellTypes  =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   0
         MaxRows         =   0
         SpreadDesigner  =   "For4.frx":0426
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label lb5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "采购修改记录"
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
         Left            =   0
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lb6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请购修改记录"
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
         Left            =   0
         TabIndex        =   15
         Top             =   6480
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lb4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "供应商"
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
         Left            =   5040
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lb3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "料号"
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
         Left            =   5040
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lb2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询单号"
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
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lb1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改方式"
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
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
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
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":082A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":147C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":20CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":2D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":3972
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":3CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":4916
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":5568
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":61BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":6E0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1535
      ButtonWidth     =   1455
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查询"
            Key             =   "QUE"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "资料修改"
            Key             =   "MOD"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "重新下单"
            Key             =   "DEL"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "EXIT"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改记录"
            Key             =   "QUERY"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "供应商"
            Key             =   "Modify"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "请购作废"
            Key             =   "Part"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "未交货"
            Key             =   "Delay"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "Frm_CGVH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_click()

    Select Case Combo1.Text

        Case "请购单号"
            lb2 = "请购单号"

        Case "采购单号"
            lb2 = "采购单号"
            
    End Select

End Sub

Private Sub Command1_Click()
    
    ForDe1

End Sub

Private Sub Command2_Click()
    
    ForMod2

End Sub
Private Sub Command3_Click()

    PartN

End Sub

Private Sub Command4_Click()

    ForQuery

End Sub

Private Sub Command5_Click()

    PartN1

End Sub

Private Sub Command6_Click()

    PartN2

End Sub

Private Sub Command7_Click()

    PartN3

End Sub


'初始化
Private Sub Initial()

    fps(0).MaxRows = 0
    fps(0).MaxCols = 0
    fpss(0).MaxRows = 0
    fpss(0).MaxCols = 0
    
    Combo1.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""

End Sub



Private Sub Form_Load()

    With fps(0)

        .Col = -1
        .Row = -1
        .Lock = True
        
        .DAutoSizeCols = DAutoSizeColsBest

    End With
    
    With fpss(0)

        .Col = -1
        .Row = -1
        .Lock = True
        
        .DAutoSizeCols = DAutoSizeColsBest

    End With

End Sub

Private Sub Datapro(strsql As String)

    Dim rs As New ADODB.Recordset
      
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType1(rs)
    Else
        
        MsgBox "查询不到该采购信息", vbInformation, "提示"
        Exit Sub

    End If

End Sub

Private Sub Datapro1(strsql As String)

    Dim rs As New ADODB.Recordset
      
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType3(rs)
    Else
        
        MsgBox "查询不到该采购信息", vbInformation, "提示"
        Exit Sub

    End If

End Sub

'料号检查
Private Sub DataDetect(strPartno1 As String)

     If Get_SqlserverCnt("select b.料号 from erpdata..tblSmainM2 b  WHERE b.料号 = '" & strPartno1 & "'") = 0 Then
        MsgBox "没有此料号,请重新输入", vbInformation, "提示"
        Exit Sub

    End If

End Sub

'请购单检查
Private Sub DataDetect1(strBuy1 As String)

    If Get_SqlserverCnt("select a.请购单编号 from erpbase..tblCRequest a  WHERE a.请购单编号 = '" & strBuy1 & "'") = 0 Then
        MsgBox "没有此请购单号,请重新输入", vbInformation, "提示"
        Exit Sub
    End If

End Sub

'采购单检查
Private Sub DataDetect2(strBuy2 As String)

    If Get_SqlserverCnt("select a.采购单编号 from erpbase..tblCPurDataSub a  WHERE a.采购单编号 = '" & strBuy2 & "'") = 0 Then
        MsgBox "没有此采购单号,请重新输入", vbInformation, "提示"
        Exit Sub

    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key

        Case "QUE"
            Initial
            
            Toolbar1.Buttons(3).Enabled = True
            Toolbar1.Buttons(3).Caption = "资料修改"
            Toolbar1.Buttons(3).Image = 6
            Toolbar1.Buttons(3).Visible = True
            
            Combo1.Visible = True
            
            lb1.Visible = True
            lb2.Visible = True
            lb3.Visible = True
            lb4.Visible = False
            lb5.Visible = False
            lb6.Visible = False
            
            Text1.Visible = True
            Text2.Visible = True
            Text3.Visible = False
            
            Command1.Visible = False
            Command2.Visible = False
            Command3.Visible = False
            '查询
            Command4.Visible = True
            Command5.Visible = False
            Command6.Visible = False
            Command7.Visible = False
        
        Case "MOD"
            
            Combo1.Visible = False
            
            lb1.Visible = False
            
            Command4.Visible = False
            
            ForMod1
               
        Case "DEL"
           
            Initial
            
            Toolbar1.Buttons(3).Enabled = False
            
            Combo1.Visible = True
            
            lb1.Visible = True
            lb2.Visible = True
            lb3.Visible = True
            lb4.Visible = False
            lb5.Visible = False
            lb6.Visible = False
            
            Text1.Visible = True
            Text2.Visible = True
            Text3.Visible = False
            
            '重新下单
            Command1.Visible = True
            Command2.Visible = False
            Command3.Visible = False
            Command4.Visible = False
            Command5.Visible = False
            Command6.Visible = False
            Command7.Visible = False

        Case "EXIT"
            Unload Me
            
        Case "QUERY"
            Initial
            
            Toolbar1.Buttons(3).Enabled = False
           
            Combo1.Visible = False
            
            lb5 = "采购修改记录"
            lb6 = "请购修改记录"
            lb1.Visible = False
            lb2.Visible = False
            lb3.Visible = False
            lb4.Visible = False
            lb5.Visible = True
            lb6.Visible = True
            
            Text1.Visible = False
            Text2.Visible = False
            Text3.Visible = False
            
            Command1.Visible = False
            Command2.Visible = False
            Command3.Visible = False
            Command4.Visible = False
            Command5.Visible = False
            Command6.Visible = False
            Command7.Visible = False
        
            Query1
            
        Case "Modify"
            Initial
            
            Toolbar1.Buttons(3).Enabled = False
            
            Combo1.Visible = False
            lb2 = "采购单号"
            lb1.Visible = False
            lb2.Visible = True
            lb3.Visible = False
            lb4.Visible = True
            lb5.Visible = False
            lb6.Visible = False
          
            Text1.Visible = True
            Text2.Visible = False
            Text3.Visible = True
            
            Command1.Visible = False
            '供应商修改
            Command2.Visible = True
            Command3.Visible = False
            Command4.Visible = False
            Command5.Visible = False
            Command6.Visible = False
            Command7.Visible = False

        Case "Part"
            Initial
            
            Toolbar1.Buttons(3).Enabled = False
           
            Combo1.Visible = False
            
            lb2 = "请购单号"
            lb1.Visible = False
            lb2.Visible = True
            lb3.Visible = True
            lb4.Visible = False
            lb5.Visible = False
            lb6.Visible = False
            
            Text1.Visible = True
            Text2.Visible = True
            Text3.Visible = False
            
            Command1.Visible = False
            Command2.Visible = False
            '请购单作废
            Command3.Visible = True
            Command4.Visible = False
            '还原
            Command5.Visible = True
            Command6.Visible = False
            Command7.Visible = False
        
        Case "Delay"
        
            Initial
            
            Toolbar1.Buttons(3).Enabled = False
           
            Combo1.Visible = True
        
            lb1.Visible = True
            lb2.Visible = True
            lb3.Visible = True
            lb4.Visible = False
            lb5.Visible = False
            lb6.Visible = False
            
            Text1.Visible = True
            Text2.Visible = True
            Text3.Visible = False
            
            Command1.Visible = False
            Command2.Visible = False
            Command3.Visible = False
            Command4.Visible = False
            Command5.Visible = False
            '未交货
            Command6.Visible = True
            '还原
            Command7.Visible = True

    End Select

End Sub

Private Sub ForQuery()

    Dim rs        As New ADODB.Recordset

    Dim strBuy    As String
    
    Dim strPartno As String

    Dim strsql    As String
    
    If Combo1.Text = "" Then
    
        MsgBox "请选择修改方式", vbInformation, "提示"
        Exit Sub

    End If
    
    If Text1.Text = "" Then
    
        MsgBox "请输入资料", vbInformation, "提示"
        Exit Sub

    End If
    
    If Text2.Text = "" Then
    
        MsgBox "请输入料号", vbInformation, "提示"
        Exit Sub
    Else
        strBuy = Trim$(Text1.Text)
     
        strPartno = Trim$(Text2.Text)

        If lb2 = "请购单号" Then
            
            Call DataDetect1(strBuy)
            
            Call DataDetect(strPartno)
            
            strsql = "select '' as  '√',a.采购单编号,a.采购单项次,a.请购单编号,a.请购单项次,a.物料编号,b.料号,a.采购数量,a.批准采购数量,a.单价,a.金额,a.币别 from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0'"
        Else
            
            Call DataDetect2(strBuy)
            
            Call DataDetect(strPartno)
            
            strsql = "select '' as  '√',a.采购单编号,a.采购单项次,a.请购单编号,a.请购单项次,a.物料编号,b.料号,a.采购数量,a.批准采购数量,a.单价,a.金额,a.币别 from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.采购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0'"
    
        End If

    End If
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        lb5 = "查询结果"
        Call ListDataType2(rs)
    Else
        
        MsgBox "查询不到资料", vbInformation, "提示"
        Exit Sub

    End If

End Sub

'料号修改后资料呈现
Private Sub ForQuery1(strBuy As String, strPartno As String)

    Dim rs        As New ADODB.Recordset
    
    Dim strsql    As String
            
    strsql = "select '' as  '√',a.采购单编号,a.采购单项次,a.请购单编号,a.请购单项次,a.物料编号,b.料号,a.采购数量,a.批准采购数量,a.单价,a.金额,a.币别 from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.采购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0'"
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        lb5 = "查询结果"
        Call ListDataType2(rs)
    Else
        
        MsgBox "查询不到资料", vbInformation, "提示"
        Exit Sub

    End If
    
End Sub

'修改记录
Private Sub Query1()

    Dim rs     As New ADODB.Recordset

    Dim strsql As String
    
    strsql = "select a.修改人,a.修改方式,a.修改状态,a.修改时间,a.采购单编号,a.采购单项次,a.采购单编号,a.请购单项次,a.物料编号,a.约定交货日期,a.请购日期,a.采购数量,a.批准采购数量,a.单价,a.金额,a.币别 from erptemp.dbo.ksrequisition a order by a.修改时间"
    
    fps(0).MaxRows = 0
    
    If rs.State = adStateOpen Then rs.Close
    
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType1(rs)
    Else
        
        MsgBox "查询不到修改记录", vbInformation, "提示"
        Exit Sub

    End If
    
    strsql = "select b.修改人,b.修改方式,b.修改状态,b.修改时间,b.请购单编号,b.请购单项次,b.物料编号,b.请购人,b.请购数量,b.交货日期,b.请购日期,b.订购标记,b.是否禁用,b.批准数量,b.订购数量,b.采购员 from erptemp.dbo.ksbuy b order by b.修改时间"
    
    fpss(0).MaxRows = 0
    
    If rs.State = adStateOpen Then rs.Close
    
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType3(rs)
    Else
        
        MsgBox "查询不到修改记录", vbInformation, "提示"
        Exit Sub

    End If
      
    
End Sub

Private Sub ListDataType1(rs As ADODB.Recordset)
   
    With fps(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With

End Sub

Private Sub ListDataType2(rs As ADODB.Recordset)

    Dim i As Long
   
    With fps(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
    With fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .ColWidth(1) = 2
            .CellType = CellTypeCheckBox
        Next

    End With

End Sub

Private Sub ListDataType3(rs As ADODB.Recordset)
   
    With fpss(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With

End Sub

'采购明细表备份
Private Sub Databackup1(strstyle As String, _
                        strstyle1 As String, _
                        strBuy As String, _
                        strPartno As String, _
                        strno As Integer)

    Dim strid      As Integer
    
    Dim strid1     As Integer
    
    Dim Userrecord As String
    
    Userrecord = gUserName '获取登录人

    strid = Get_SqlStr("select isnull(max(id),0) from erptemp.dbo.ksrequisition where 修改方式 = '" & strstyle & "' ")
    
    strid1 = strid + 1
    
    AddSql2 (" insert into erptemp.dbo.ksrequisition(id,修改人,修改方式,修改状态,修改时间,是否禁用,采购单编号,采购单项次,请购单编号,请购单项次,物料编号,约定交货日期,请购日期,采购数量,批准采购数量,单价,金额,币别) select '" & strid1 & "','" & Userrecord & "','" & strstyle & "','" & strstyle1 & "',GetDate(),a.是否禁用,a.采购单编号,a.采购单项次,a.请购单编号,a.请购单项次,a.物料编号,a.约定交货日期,a.请购日期,a.采购数量,a.批准采购数量,a.单价,a.金额,a.币别  from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.采购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '" & strno & "'")

End Sub

'请购明细表备份
Private Sub Databackup2(strstyle As String, _
                        strstyle1 As String, _
                        strBuy As String, _
                        strPartno As String, _
                        strno As Integer)

    Dim strid      As Integer
    
    Dim strid1     As Integer

    Dim Userrecord As String
    
    Userrecord = gUserName '获取登录人
    
    strid = Get_SqlStr("select isnull(max(id),0) from erptemp.dbo.ksbuy where 修改方式 = '" & strstyle & "' ")
        
    strid1 = strid + 1
        
    AddSql2 (" insert into erptemp.dbo.ksbuy(id,修改人,修改方式,修改状态,修改时间,请购单编号,请购单项次,物料编号,请购人,请购数量,交货日期,请购日期,订购标记,是否禁用,批准数量,订购数量,采购员) select '" & strid1 & "','" & Userrecord & "','" & strstyle & "','" & strstyle1 & "',GetDate(),a.请购单编号,a.请购单项次,a.物料编号,a.请购人,a.请购数量,a.交货日期,a.请购日期,a.订购标记,a.是否禁用,a.批准数量,a.订购数量,a.采购员 from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号  = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '" & strno & "' ")

End Sub

Private Sub ForMod1()

    Dim i         As Integer

    Dim m         As Integer

    Dim J         As Integer

    Dim strstyle  As String

    Dim strstyle1 As String

    Dim strstyle2 As String
    
    Dim strpu1    As String

    Dim strpu2    As String

    Dim strpu3    As String

    Dim strpu4    As String

    Dim strpu5    As String

    Dim strpu6    As String

    Dim strpu7    As String

    Dim strpu8    As String

    Dim strpu9    As String

    Dim strpu10   As String

    Dim strpu11   As String

    Dim strpua    As String
    
    Dim strPartno As String
    
    Dim bFlag     As Boolean
    
    Dim strsql    As String
    
    strPartno = Trim$(Text2.Text)

    If Toolbar1.Buttons(3).Caption <> "确认修改" Then

        With fps(0)

            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                .Lock = False
      
                For m = 7 To 10
            
                    .Col = m
                    .Lock = False
      
                Next
    
            Next
        
        End With
    
        Toolbar1.Buttons(3).Caption = "确认修改"
        Toolbar1.Buttons(3).Image = 10
        
        Exit Sub

    End If

    bFlag = False
    
    With fps(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
    
            J = 0

            If .Text = "1" Then
            
                J = J + 1
                bFlag = True
    
                .Col = 2
                strpu1 = Trim$(.Text)
                
                .Col = 3
                strpu2 = Trim$(.Text)
                
                .Col = 4
                strpu3 = Trim$(.Text)
                
                .Col = 5
                strpu4 = Trim$(.Text)
                
                .Col = 6
                strpu5 = Trim$(.Text)
                
                '料号modify
                .Col = 7
                strpu6 = Trim$(.Text)
                
                If Trim$(strpu6) <> Trim$(strPartno) Then
                    
                    If Get_SqlserverCnt("select distinct 料号 from erpdata..tblSmainM2 where 料号 = '" & strpu6 & "'") = 0 Then
                                                
                        MsgBox "没有此料号请确认", vbInformation, "提示"
                        Exit Sub

                    End If
                    
                    '获取新的物料编号
                    strpua = Get_SqlStr("select distinct 物料编号 from erpdata..tblSmainM2 where 料号 = '" & strpu6 & "'")
                    Else
                 
                    strpua = Get_SqlStr("select distinct 物料编号 from erpdata..tblSmainM2 where 料号 = '" & strPartno & "'")

                    
                End If
                
                '采购数量modify
                
                .Col = 8
                strpu7 = Trim$(.Text)
                
                '批准采购数量modify
                .Col = 9
                strpu8 = Trim$(.Text)
                
                '单价modify
                .Col = 10
                strpu9 = Trim$(.Text)
                
                .Col = 11
                strpu10 = Trim$(.Text)
                
                .Col = 12
                strpu11 = Trim$(.Text)
                
                strstyle = "资料修改"
                
                strstyle1 = "修改前"
                
                strstyle2 = "修改后"
                
                Call Databackup1(strstyle, strstyle1, strpu1, strPartno, 0)
                
                AddSql2 (" update a set a.物料编号 = '" & strpua & "',a.采购数量 = '" & strpu7 & "',a.批准采购数量 = '" & strpu8 & "',a.单价 = '" & strpu9 & "' from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.采购单编号 = '" & strpu1 & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0'")
                
                Call Databackup1(strstyle, strstyle2, strpu1, strpu6, 0)
                
                Call Databackup2(strstyle, strstyle1, strpu3, strPartno, 0)

                AddSql2 ("update a set a.物料编号 = '" & strpua & "',a.批准数量 = '" & strpu8 & "',a.请购数量 = '" & strpu8 & "',a.订购数量 = '" & strpu8 & "' from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strpu3 & "' and  b.料号 = '" & strPartno & "' and a.是否禁用 = '0' ")
                
                Call Databackup2(strstyle, strstyle2, strpu3, strpu6, 0)
            
            End If
            
        Next
        
        If bFlag = False And J = 0 Then
            MsgBox "请选择要修改的行", vbInformation, "提示"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "修改成功", vbInformation, "提示"

    Toolbar1.Buttons(3).Caption = "资料修改"
    Toolbar1.Buttons(3).Image = 6
    
    If Trim$(strpu6) = Trim$(strPartno) Then
    
        lb5 = "采购明细资料"
        lb5.Visible = True
        
        ForQuery
        
        strsql = "select a.是否禁用,a.请购单编号,a.请购单项次,a.物料编号,a.请购人,a.请购数量,a.交货日期,a.请购日期,a.订购标记,a.批准数量,a.订购数量,a.采购员 from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号  = '" & strpu3 & "' and  a.是否禁用 = '0' and b.料号 = '" & strPartno & "'"
        
        lb6 = "请购明细资料"
        lb6.Visible = True
        
        Call Datapro1(strsql)
    Else
        lb5 = "采购明细资料"
        lb5.Visible = True
        
        Call ForQuery1(strpu1, strpu6)
        
        strsql = "select a.是否禁用,a.请购单编号,a.请购单项次,a.物料编号,a.请购人,a.请购数量,a.交货日期,a.请购日期,a.订购标记,a.批准数量,a.订购数量,a.采购员 from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号  = '" & strpu3 & "' and  a.是否禁用 = '0' and b.料号 = '" & strpu6 & "'"
        
        lb6 = "请购明细资料"
        lb6.Visible = True
        Call Datapro1(strsql)

    End If

End Sub

Private Sub ForDe1()
    
    Dim strBuy    As String
    
    Dim strBuy1   As String

    Dim strPartno As String
    
    Dim strstyle  As String

    Dim strstyle1 As String
    
    Dim strstyle2 As String
 
    Dim strsql    As String
    
    If Combo1.Text = "" Then
        MsgBox "请选择修改方式", vbInformation, "提示"
        Exit Sub
            
    End If
            
    If Combo1.Text <> "采购单号" And Combo1.Text <> "请购单号" Then
            
        MsgBox "请选择正确的修改方式", vbInformation, "提示"
            
        Exit Sub
            
    End If
    
    If Text1.Text = "" Then
        MsgBox "请输入单号", vbInformation, "提示"
        Exit Sub

    End If
       
    strBuy = Trim$(Text1.Text)
    
    If lb2 = "采购单号" Then
    
        Call DataDetect2(strBuy)
    
    Else

        Call DataDetect1(strBuy)

    End If

    If Text2.Text = "" Then
        MsgBox "请输入料号", vbInformation, "提示"
        Exit Sub

    End If
    
    strPartno = Trim$(Text2.Text)
    
    Call DataDetect(strPartno)

    If lb2 = "采购单号" Then
        If Get_SqlserverCnt("select a.采购单编号,a.请购单编号,a.物料编号 from erpbase..tblCPurDataSub a inner join erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 WHERE a.采购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0'") = 0 Then
            MsgBox "没有此笔数据,请重新输入", vbInformation, "提示"
            Exit Sub

        End If
    
        strsql = "select a.采购单编号,a.请购单编号,a.物料编号 from erpbase..tblCPurDataSub a inner join erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 WHERE a.采购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0'"
    Else

        If Get_SqlserverCnt("select a.采购单编号,a.请购单编号,a.物料编号 from erpbase..tblCPurDataSub a inner join erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 WHERE a.请购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0'") = 0 Then
            MsgBox "没有此笔数据,请重新输入", vbInformation, "提示"
            Exit Sub

        End If

        strsql = "select a.采购单编号,a.请购单编号,a.物料编号 from erpbase..tblCPurDataSub a inner join erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 WHERE a.请购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0'"

    End If
    
    lb5 = "修改中的资料"
    lb5.Visible = True
    
    Call Datapro(strsql)
    
    strstyle = "重新下单"
    
    strstyle1 = "修改前"
        
    strstyle2 = "修改后"
    
    If lb2 = "采购单号" Then
        
        '获取请购单编号
        strBuy1 = Get_SqlStr("select distinct c.请购单编号 from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.物料编号 = d.物料编号 WHERE c.采购单编号 = '" & strBuy & "' and d.料号 = '" & strPartno & "' and c.是否禁用 = '0'")
                
        Call Databackup1(strstyle, strstyle1, strBuy, strPartno, 0)
        '采购明细表信息更改'
        AddSql2 (" update a set a.是否禁用 = '1'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.采购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0'")
    
        '修改后的资料backup
             
        Call Databackup1(strstyle, strstyle2, strBuy, strPartno, 1)
    
        '请购明细表资料backup
                
        Call Databackup2(strstyle, strstyle1, strBuy1, strPartno, 0)
        
        '请购明细表信息更改
        AddSql2 ("update a set a.订购数量 = '0',a.订购标记 = '0' from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy1 & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0'")
        
        '修改后的资料backup

        Call Databackup2(strstyle, strstyle2, strBuy1, strPartno, 0)
        
        strsql = "select a.请购单编号,a.请购单项次,a.物料编号,a.订购标记,a.订购数量 from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where b.料号 = '" & strPartno & "' and a.请购单编号 = '" & strBuy1 & "' and a.是否禁用 = '0'"
    Else
        
        '获取采购单编号
        strBuy1 = Get_SqlStr("select distinct c.采购单编号 from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.物料编号 = d.物料编号 WHERE c.请购单编号 = '" & strBuy & "' and d.料号 = '" & strPartno & "' and c.是否禁用 = '0'")
        
        Call Databackup1(strstyle, strstyle1, strBuy1, strPartno, 0)
        
        '采购明细表信息更改'
        AddSql2 (" update a set a.是否禁用 = '1'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.采购单编号 = '" & strBuy1 & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0' ")
    
        '修改后的资料backup
        
        Call Databackup1(strstyle, strstyle2, strBuy1, strPartno, 1)
        
        '请购明细表资料backup
                
        Call Databackup2(strstyle, strstyle1, strBuy, strPartno, 0)
        
        '请购明细表信息更改
        AddSql2 ("update a set a.订购数量 = '0',a.订购标记 = '0' from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0' ")
        
        '修改后的资料backup
                
        Call Databackup2(strstyle, strstyle2, strBuy, strPartno, 0)
        
        strsql = "select a.请购单编号,a.请购单项次,a.物料编号,a.订购标记,a.订购数量 from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0' "

    End If
    
    lb6 = "修改后的请购资料"
    lb6.Visible = True
    
    Call Datapro1(strsql)
    
    MsgBox "重新下单成功", vbInformation, "提示"

End Sub

Private Sub ForMod2()

    Dim strBuy      As String
    
    Dim strprovider As String
    
    Dim strsupply   As String

    Dim strsupply1  As String

    Dim strsql      As String
    
    Dim Userrecord  As String
    
    
    
    Userrecord = gUserName '获取登录人
    
    If Text1.Text = "" Then
        MsgBox "请输入采购单号", vbInformation, "提示"
        Exit Sub

    End If
    
    strBuy = Trim$(Text1.Text)
    
    Call DataDetect2(strBuy)
    
    If Text3.Text = "" Then
        MsgBox "请输入供应商名称", vbInformation, "提示"
        Exit Sub

    End If
    
    strprovider = Trim$(Text3.Text)
    
    If Get_SqlserverCnt("SELECT distinct 供应商编号 FROM ERPBASE..tblSupplierData where 供应商名称 = '" & strprovider & "'") = 0 Then
        MsgBox "没有此供应商信息,请重新输入", vbInformation, "提示"
        Exit Sub

    End If
    
    strsupply = Get_SqlStr("SELECT distinct 供应商编号 FROM ERPBASE..tblSupplierData where 供应商名称 = '" & strprovider & "'")
    
    strsupply1 = Get_SqlStr("SELECT distinct 供应商编号 FROM erpbase..tblcpurdata where 采购单编号 = '" & strBuy & "' and 是否禁用 = '0'")

    If Trim$(strsupply) = Trim$(strsupply1) Then
        MsgBox "供应商编号已经保持一致无需修改！", vbInformation, "提示"
        Exit Sub

    End If
    
    '资料backup
    AddSql2 ("insert into erptemp.dbo.kspur(修改人,修改方式,修改状态,修改时间,采购单号,供应商编号) select '" & Userrecord & "','供应商更改','修改前',GetDate(),采购单编号,供应商编号 from erpbase..tblcpurdata where 采购单编号 = '" & strBuy & "' and 是否禁用 = '0'")
    
    AddSql2 ("update erpbase..tblcpurdata set 供应商编号 = '" & strsupply & "' where 采购单编号 = '" & strBuy & "' and 是否禁用 = '0'")
    
    AddSql2 ("insert into erptemp.dbo.kspur(修改人,修改方式,修改状态,修改时间,采购单号,供应商编号) select '" & Userrecord & "','供应商更改','修改后',GetDate(),采购单编号,供应商编号 from erpbase..tblcpurdata where 采购单编号 = '" & strBuy & "' and 是否禁用 = '0'")
    '修改后资料呈现
    
    strsql = "select distinct m.采购单号,m.供应商编号 as 修改前供应商编号,h.供应商名称 as 修改前供应商名称,n.供应商编号 as 修改后供应商编号,'" & strprovider & "' as  修改后供应商名称 from erptemp.dbo.kspur m inner join erpbase..tblcpurdata n on m.采购单号 = n.采购单编号 left join  ERPBASE..tblSupplierData h on h.供应商编号 = m.供应商编号 where m.采购单号 = '" & strBuy & "' and m.修改状态 = '修改前' and n.是否禁用 = '0' "

    '修改前资料呈现
    lb5 = "修改状态资料"
    lb5.Visible = True
    
    Call Datapro(strsql)

    '修改记录
    strsql = "select * from erptemp.dbo.kspur where 1 = 1  "
    
    lb6 = "历史修改记录"
    lb6.Visible = True
    Call Datapro1(strsql)
    
    MsgBox "供应商修改成功", vbInformation, "提示"
 
End Sub

Private Sub PartN()

    Dim strBuy    As String
    
    Dim strPartno As String

    Dim strstyle  As String

    Dim strstyle1 As String
    
    Dim strstyle2 As String
    
    Dim strsql    As String
    
    Dim strid1 As String

    strstyle = "请购作废"
    
    strstyle1 = "修改前"
    
    strstyle2 = "修改后"
    
    If Text1.Text = "" Then
        MsgBox "请输入请购单号", vbInformation, "提示"
        Exit Sub

    End If
    
    strBuy = Trim$(Text1.Text)
    
    Call DataDetect1(strBuy)

    If Text2.Text = "" Then
        MsgBox "请输入料号", vbInformation, "提示"
        Exit Sub

    End If
    
    strPartno = Trim$(Text2.Text)
    
    Call DataDetect(strPartno)
    
    If Get_SqlserverCnt("select a.请购单编号 from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 WHERE a.请购单编号 = '" & strBuy & "' and  a.是否禁用 = '0' and b.料号 = '" & strPartno & "'") = 0 Then
        MsgBox "此笔请购单已经作废,请确认", vbInformation, "提示"
        Exit Sub
        
    End If
    
    If Get_SqlserverCnt("select a.采购单编号 from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号  WHERE a.请购单编号 = '" & strBuy & "' and  a.是否禁用 = '0' and b.料号 = '" & strPartno & "' ") <> 0 Then
        MsgBox "此笔请购单已经进行采购动作,请确认", vbInformation, "提示"
        Exit Sub
        
    End If
    
    Call Databackup2(strstyle, strstyle1, strBuy, strPartno, 0)
    
    strid1 = Get_SqlStr("select isnull(max(id),0) from erptemp.dbo.ksbuy where 修改方式 = '" & strstyle & "' ")
    
    strsql = "select 是否禁用,请购单编号,请购单项次,物料编号,请购人,请购数量,交货日期,请购日期,订购标记,批准数量,订购数量,采购员 from erptemp.dbo.ksbuy where 请购单编号  = '" & strBuy & "'and 修改状态 = '修改前' and id = '" & strid1 & "'  order by 请购单项次 "
    
    lb5 = "修改前的资料"
    lb5.Visible = True
    Call Datapro(strsql)
    
    AddSql2 (" update a set a.是否禁用 = '1'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0' ")
    
    Call Databackup2(strstyle, strstyle2, strBuy, strPartno, 1)
    
    strsql = "select a.是否禁用,a.请购单编号,a.请购单项次,a.物料编号,a.请购人,a.请购数量,a.交货日期,a.请购日期,a.订购标记,a.批准数量,a.订购数量,a.采购员 from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号  = '" & strBuy & "' and  a.是否禁用 = '1' and b.料号 = '" & strPartno & "' order by a.请购单项次 "
        
    lb6 = "修改后的资料"
    lb6.Visible = True
     
    Call Datapro1(strsql)
     
    MsgBox "已经作废", vbInformation, "提示"
    
End Sub

'还原

Private Sub PartN1()

    Dim strBuy    As String
    
    Dim strPartno As String
    
    Dim strstyle  As String

    Dim strstyle1 As String
    
    Dim strstyle2 As String
    
    Dim strsql    As String
    
    strstyle = "请购作废还原"
    
    strstyle1 = "修改前"
    
    strstyle2 = "修改后"
    
    If Text1.Text = "" Then
        MsgBox "请输入请购单号", vbInformation, "提示"
        Exit Sub

    End If
    
    strBuy = Trim$(Text1.Text)
    
    Call DataDetect1(strBuy)

    If Text2.Text = "" Then
        MsgBox "请输入料号", vbInformation, "提示"
        Exit Sub

    End If
    
    strPartno = Trim$(Text2.Text)
    
    Call DataDetect(strPartno)
    
    If Get_SqlserverCnt("select a.请购单编号 from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 WHERE a.请购单编号 = '" & strBuy & "' and  a.是否禁用 = '1' and b.料号 = '" & strPartno & "'") = 0 Then
        MsgBox "此笔请购单没有作废记录,请确认", vbInformation, "提示"
        Exit Sub
        
    End If
    
    If Get_SqlserverCnt("select a.采购单编号 from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号  WHERE a.请购单编号 = '" & strBuy & "' and  a.是否禁用 = '0' and b.料号 = '" & strPartno & "' ") <> 0 Then
        MsgBox "此笔请购单已经进行采购动作,请确认", vbInformation, "提示"
        Exit Sub
        
    End If
    
    Call Databackup2(strstyle, strstyle1, strBuy, strPartno, 1)
    
    Dim strid1    As String
strid1 = Get_SqlStr("select isnull(max(id),0) from erptemp.dbo.ksbuy where 修改方式 = '" & strstyle & "' ")
    
    strsql = "select 是否禁用,请购单编号,请购单项次,物料编号,请购人,请购数量,交货日期,请购日期,订购标记,批准数量,订购数量,采购员 from erptemp.dbo.ksbuy where 请购单编号  = '" & strBuy & "'and 修改状态 = '修改前' and id = '" & strid1 & "'  order by 请购单项次 "
    
    lb5 = "修改前的资料"
    lb5.Visible = True
    Call Datapro(strsql)
    
    AddSql2 (" update a set a.是否禁用 = '0'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '1' ")
     
    Call Databackup2(strstyle, strstyle2, strBuy, strPartno, 0)
     
    strsql = "select a.是否禁用,a.请购单编号,a.请购单项次,a.物料编号,a.请购人,a.请购数量,a.交货日期,a.请购日期,a.订购标记,a.批准数量,a.订购数量,a.采购员 from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号  = '" & strBuy & "' and  a.是否禁用 = '0' and b.料号 = '" & strPartno & "' order by a.请购单项次 "
        
    lb6 = "修改后的资料"
    lb6.Visible = True
     
    Call Datapro1(strsql)
        
    MsgBox "还原成功", vbInformation, "提示"

End Sub

Private Sub PartN2()
    
    Dim strBuy    As String
    
    Dim strBuy1   As String
    
    Dim strPartno As String
    
    Dim strstyle  As String

    Dim strstyle1 As String
    
    Dim strstyle2 As String
     
    strstyle = "未交货作废"
    
    strstyle1 = "修改前"
    
    strstyle2 = "修改后"
     
    If Combo1.Text = "" Then
        MsgBox "请选择修改方式", vbInformation, "提示"
        Exit Sub
            
    End If
            
    If Combo1.Text <> "采购单号" And Combo1.Text <> "请购单号" Then
            
        MsgBox "请选择正确的修改方式", vbInformation, "提示"
            
        Exit Sub
            
    End If
    
    If Text1.Text = "" Then
        MsgBox "请输入单号", vbInformation, "提示"
        Exit Sub

    End If
       
    strBuy = Trim$(Text1.Text)
    
    If lb2 = "采购单号" Then
    
        Call DataDetect2(strBuy)
    
    Else

        Call DataDetect1(strBuy)

    End If

    If Text2.Text = "" Then
        MsgBox "请输入料号", vbInformation, "提示"
        Exit Sub

    End If
    
    strPartno = Trim$(Text2.Text)
    
    Call DataDetect(strPartno)
    
    If lb2 = "请购单号" Then
    
        If Get_SqlserverCnt("select a.请购单编号 from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 WHERE a.请购单编号 = '" & strBuy & "' and  a.是否禁用 = '0' and b.料号 = '" & strPartno & "'") = 0 Then
            MsgBox "此笔请购单已经作废,请确认", vbInformation, "提示"
            Exit Sub
        
        End If
    
        If Get_SqlserverCnt("select a.采购单编号 from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号  WHERE a.请购单编号 = '" & strBuy & "' and  a.是否禁用 = '0' and b.料号 = '" & strPartno & "' ") = 0 Then
            MsgBox "此笔采购单已经作废,请确认", vbInformation, "提示"
            Exit Sub
    
        End If
                
        '获取采购单号
        strBuy1 = Get_SqlStr("select distinct c.采购单编号 from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.物料编号 = d.物料编号 WHERE c.请购单编号 = '" & strBuy & "' and d.料号 = '" & strPartno & "'")
        
        Call Databackup2(strstyle, strstyle1, strBuy, strPartno, 0)
        
        AddSql2 (" update a set a.是否禁用 = '1'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0' ")
        
        Call Databackup2(strstyle, strstyle2, strBuy, strPartno, 1)
        
        Call Databackup1(strstyle, strstyle1, strBuy1, strPartno, 0)
        
        AddSql2 (" update a set a.是否禁用 = '1'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0' ")
           
        Call Databackup1(strstyle, strstyle2, strBuy1, strPartno, 1)
    
    Else
    
        '获取请购单号
        strBuy1 = Get_SqlStr("select distinct c.请购单编号 from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.物料编号 = d.物料编号 WHERE c.采购单编号 = '" & strBuy & "' and d.料号 = '" & strPartno & "'")

        If Get_SqlserverCnt("select a.请购单编号 from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 WHERE a.请购单编号 = '" & strBuy1 & "' and  a.是否禁用 = '0' and b.料号 = '" & strPartno & "'") = 0 Then
            MsgBox "此笔请购单已经作废,请确认", vbInformation, "提示"
            Exit Sub
        
        End If
    
        If Get_SqlserverCnt("select a.采购单编号 from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号  WHERE a.请购单编号 = '" & strBuy1 & "' and  a.是否禁用 = '0' and b.料号 = '" & strPartno & "' ") = 0 Then
            MsgBox "此笔采购单已经作废,请确认", vbInformation, "提示"
            Exit Sub
    
        End If

        Call Databackup2(strstyle, strstyle1, strBuy1, strPartno, 0)
        
        AddSql2 (" update a set a.是否禁用 = '1'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy1 & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0' ")
      
        Call Databackup2(strstyle, strstyle2, strBuy1, strPartno, 1)
        
        Call Databackup1(strstyle, strstyle1, strBuy, strPartno, 0)
        
        AddSql2 (" update a set a.是否禁用 = '1'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy1 & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '0' ")
        
        Call Databackup1(strstyle, strstyle2, strBuy, strPartno, 1)

    End If
    
    MsgBox "已经作废", vbInformation, "提示"

End Sub

Private Sub PartN3()

    Dim strBuy    As String
    
    Dim strBuy1   As String
    
    Dim strPartno As String
    
    Dim strstyle  As String

    Dim strstyle1 As String
    
    Dim strstyle2 As String
     
    strstyle = "未交货还原"
    
    strstyle1 = "修改前"
    
    strstyle2 = "修改后"
     
    If Combo1.Text = "" Then
        MsgBox "请选择修改方式", vbInformation, "提示"
        Exit Sub
            
    End If
            
    If Combo1.Text <> "采购单号" And Combo1.Text <> "请购单号" Then
            
        MsgBox "请选择正确的修改方式", vbInformation, "提示"
            
        Exit Sub
            
    End If
    
    If Text1.Text = "" Then
        MsgBox "请输入单号", vbInformation, "提示"
        Exit Sub

    End If
       
    strBuy = Trim$(Text1.Text)
    
    If lb2 = "采购单号" Then
    
        Call DataDetect2(strBuy)
    
    Else

        Call DataDetect1(strBuy)

    End If

    If Text2.Text = "" Then
        MsgBox "请输入料号", vbInformation, "提示"
        Exit Sub

    End If
    
    strPartno = Trim$(Text2.Text)
    
    Call DataDetect(strPartno)
    
    If lb2 = "请购单号" Then
        
        '获取采购单号
        strBuy1 = Get_SqlStr("select distinct c.采购单编号 from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.物料编号 = d.物料编号 WHERE c.请购单编号 = '" & strBuy & "' and d.料号 = '" & strPartno & "'")
   
        If Get_SqlserverCnt("select a.请购单编号 from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 WHERE a.请购单编号 = '" & strBuy & "' and  a.是否禁用 = '1' and b.料号 = '" & strPartno & "'") = 0 Then
            MsgBox "此笔请购单未作废无需还原,请确认", vbInformation, "提示"
            Exit Sub
        
        End If
    
        If Get_SqlserverCnt("select a.采购单编号 from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号  WHERE a.请购单编号 = '" & strBuy & "' and  a.是否禁用 = '1' and b.料号 = '" & strPartno & "' ") = 0 Then
            MsgBox "此笔采购单未作废无需还原,请确认", vbInformation, "提示"
            Exit Sub
    
        End If
        
        Call Databackup2(strstyle, strstyle1, strBuy, strPartno, 1)
        
        AddSql2 (" update a set a.是否禁用 = '0'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '1' ")
      
        Call Databackup2(strstyle, strstyle2, strBuy, strPartno, 0)
        
        Call Databackup1(strstyle, strstyle1, strBuy1, strPartno, 1)
      
        AddSql2 (" update a set a.是否禁用 = '0'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '1' ")
        
        Call Databackup1(strstyle, strstyle2, strBuy1, strPartno, 0)
    Else
    
        '获取请购单号
        strBuy1 = Get_SqlStr("select distinct c.请购单编号 from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.物料编号 = d.物料编号 WHERE c.采购单编号 = '" & strBuy & "' and d.料号 = '" & strPartno & "'")

        If Get_SqlserverCnt("select a.请购单编号 from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 WHERE a.请购单编号 = '" & strBuy1 & "' and  a.是否禁用 = '1' and b.料号 = '" & strPartno & "'") = 0 Then
            MsgBox "此笔请购单未作废无需还原,请确认", vbInformation, "提示"
            Exit Sub
        
        End If
    
        If Get_SqlserverCnt("select a.采购单编号 from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号  WHERE a.请购单编号 = '" & strBuy1 & "' and  a.是否禁用 = '1' and b.料号 = '" & strPartno & "' ") = 0 Then
            MsgBox "此笔采购单未作废无需还原,请确认", vbInformation, "提示"
            Exit Sub
    
        End If

        Call Databackup2(strstyle, strstyle1, strBuy1, strPartno, 1)
         
        AddSql2 (" update a set a.是否禁用 = '0'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy1 & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '1' ")
        
        Call Databackup2(strstyle, strstyle2, strBuy1, strPartno, 0)
        
        Call Databackup1(strstyle, strstyle1, strBuy, strPartno, 1)
      
        AddSql2 (" update a set a.是否禁用 = '0'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.物料编号 = b.物料编号 where a.请购单编号 = '" & strBuy1 & "' and b.料号 = '" & strPartno & "' and a.是否禁用 = '1' ")

        Call Databackup1(strstyle, strstyle2, strBuy, strPartno, 0)
        
    End If
    
    MsgBox "还原成功", vbInformation, "提示"

End Sub
