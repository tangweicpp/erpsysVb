VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_SetData 
   Caption         =   "信息维护"
   ClientHeight    =   10845
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
   ScaleHeight     =   10845
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtKey2 
      Height          =   285
      Left            =   6840
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_SetData.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_SetData.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_SetData.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_SetData.frx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_SetData.frx":3148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_SetData.frx":3D9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   1535
      ButtonWidth     =   1032
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查询"
            Key             =   "QUE"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "ADD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "MOD"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除"
            Key             =   "DEL"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "EXIT"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   9975
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   19935
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Frm_SetData.frx":49EC
         Left            =   12720
         List            =   "Frm_SetData.frx":49EE
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdCommand1 
         Caption         =   "除账"
         Height          =   600
         Left            =   6960
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtKey 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   3975
      End
      Begin VB.ComboBox cmbCombo1 
         Height          =   315
         ItemData        =   "Frm_SetData.frx":49F0
         Left            =   1200
         List            =   "Frm_SetData.frx":4A06
         TabIndex        =   1
         Top             =   645
         Width           =   3975
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   7455
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   18615
         _Version        =   524288
         _ExtentX        =   32835
         _ExtentY        =   13150
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
         SpreadDesigner  =   "Frm_SetData.frx":4A67
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lab2 
         Caption         =   "类别"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12120
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询条件2"
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
         Index           =   2
         Left            =   5760
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询条件"
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
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "维护类型"
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
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   660
         Width           =   960
      End
   End
End
Attribute VB_Name = "Frm_SetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCombo1_Click()

    Select Case cmbCombo1.Text

        Case "合格供应商合格物料"
            lbl(1) = "查询料号"
            lbl(2).Visible = False
            txtKey2.Visible = False
            'lbl(3).Visible = False
            'DTPicker1.Visible = False
            
        Case "客户出货调仓配置"
            lbl(1) = "查询客户代码"
            
            lbl(2).Visible = False
            txtKey2.Visible = False
            'lbl(3).Visible = False
            'DTPicker1.Visible = False
            
        Case "ERP物料有效期更新"
            lbl(1) = "查询料号"
            lbl(2) = "批号"
            lbl(2).Visible = True
            txtKey2.Visible = True
        

            
        Case "待验仓除账"
            lbl(1) = "查询客户代码"
            lbl(2) = "批号"
            lbl(2).Visible = True
            txtKey2.Visible = True
            cmdCommand1.Visible = True
            
        Case "出口明细表"
            lbl(1) = "出货单据"
'           lbl(2) = "料号"
'           lbl(2).Visible = True
'           txtKey2.Visible = True
            lbl(3).Visible = False
            DTPicker1.Visible = False
            lab2 = "类别"
            Combo1.Clear
            Combo1.AddItem ("保税原材料")
            Combo1.AddItem ("非保税原材料")
            Combo1.AddItem ("零部件")
            Combo1.AddItem ("保税设备")
            Combo1.AddItem ("非保税设备")
            Combo1.AddItem ("暂时进出口")
            Combo1.AddItem ("成品复进")
        
        Case "进口明细表"
            lbl(1) = "采购单号"
'            lbl(2) = "料号"
'            lbl(2).Visible = True
'            txtKey2.Visible = True
            lbl(3).Visible = False
            DTPicker1.Visible = False
            lab2 = "类别"
'            Combo1.Visible = True
            Combo1.Clear
            Combo1.AddItem ("保税成品")
            Combo1.AddItem ("非保税成品")
            Combo1.AddItem ("料件退运")
            Combo1.AddItem ("零部件退运")
            Combo1.AddItem ("设备退运")
            Combo1.AddItem ("暂时进出口退运")

    End Select

End Sub

Private Sub Form_Load()

    With fpS(0)
    
        .Col = -1
        .Row = -1
        .Lock = True

    End With

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key

        Case "QUE"
            ForQuery

        Case "ADD"
            ForAdd
        
        Case "MOD"

            Select Case cmbCombo1.Text

                Case "合格供应商合格物料"
                    ForMod1
        
                Case "客户出货调仓配置"
                    ForMod2
                
                Case "ERP物料有效期更新"
                    'MsgBox "暂时下线"
                    ForMod3

                Case "待验仓除账"
                    ForMod3
                    
                Case "出口明细表"
                    ForMod5
                
                Case "进口明细表"
                    ForMod6
                   
            End Select
        
        Case "DEL"
            
            Select Case cmbCombo1.Text

                Case "合格供应商合格物料"
                    ForDel1
        
                Case "客户出货调仓配置"
                    ForDel2
                    
                Case "ERP物料有效期更新"
                    MsgBox "该维护类型不支持删除操作", vbInformation, "提示"
                    Exit Sub
                    
                Case "出口明细表"
                    ForDel5
                
                Case "进口明细表"
                    ForDel6

            End Select

        Case "EXIT"
            Unload Me

    End Select

End Sub

Private Sub ForQuery()

    If cmbCombo1.Text = "" Then
        MsgBox "请选择维护类型", vbInformation, "提示"
        Exit Sub

    End If

    Select Case cmbCombo1.Text

        Case "合格供应商合格物料"
            QueType1
        
        Case "客户出货调仓配置"
            QueType2
        
        Case "ERP物料有效期更新"
            QueType3
            
        Case "待验仓除账"
            QueType4
        
        Case "出口明细表"
            QueType5
                
        Case "进口明细表"
            QueType6
        

    End Select

End Sub

Private Sub QueType1()

    Dim rs     As New ADODB.Recordset

    Dim strMat As String

    Dim strSql As String
    
    strMat = Trim$(txtKey.Text)

    If txtKey.Text = "" Then
        strSql = "select 序号,供应商编号,供应商名称,料号, 物料编号,有效否,创建时间,'' as '√' from ERPBASE..tblCG_PassSupplier where 供应商名称 <> ''"
    Else
        strSql = "select 序号,供应商编号,供应商名称,料号, 物料编号,有效否,创建时间,'' as '√' from ERPBASE..tblCG_PassSupplier where 料号 = '" & strMat & "' and 供应商名称 <> ''"

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType1(rs)
    Else
        
        MsgBox "查询不到该料号", vbInformation, "提示"
        Exit Sub

    End If

End Sub

Private Sub QueType2()

    Dim rs         As New ADODB.Recordset

    Dim strCusCode As String

    Dim strSql     As String

    If txtKey.Text = "" Then
        strSql = "select CUSTOMER as 客户代码,WAREHOUSE as 出货单号, FLAG as 是否有效,'' as '√' from erptemp..tbltransfer"
    Else
        strCusCode = Trim$(txtKey.Text)
        strSql = "select CUSTOMER as 客户代码,WAREHOUSE as 出货单号, FLAG as 是否有效,'' as '√' from erptemp..tbltransfer where CUSTOMER = '" & strCusCode & "' "

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType2(rs)
    Else
        
        MsgBox "查询不到该客户代码", vbInformation, "提示"
        Exit Sub

    End If

End Sub

Private Sub QueType3()

    Dim rs         As New ADODB.Recordset

    Dim strCusCode As String

    Dim strSql     As String
    
    If txtKey.Text = "" Then
        MsgBox "请输入物料的料号", vbInformation, "提示"
        Exit Sub

    End If
    
    If txtKey2.Text = "" Then
        strSql = "select '' as '√',AA.id,AA.仓库编号,BB.F_101 as 料号,BB.FName as 物料名称,AA.物料编号,AA.批号, AA.有效期至,AA.建立日期, AA.库位, AA.当前存量 from erpbase.dbo.tblStockNum AA INNER JOIN  AIS20141114094336.dbo.t_ICItem BB ON AA.物料编号 = BB.FNumber AND   BB.F_101 = '" & UCase(Trim(txtKey.Text)) & "' and AA.当前存量 > 0 "
    Else
        strSql = "select '' as '√',AA.id,AA.仓库编号,BB.F_101 as 料号,BB.FName as 物料名称,AA.物料编号,AA.批号, AA.有效期至,AA.建立日期, AA.库位, AA.当前存量 from erpbase.dbo.tblStockNum AA INNER JOIN  AIS20141114094336.dbo.t_ICItem BB ON AA.物料编号 = BB.FNumber AND   BB.F_101 = '" & UCase(Trim(txtKey.Text)) & "' and AA.当前存量 > 0 and AA.批号 = '" & txtKey2.Text & "' "
        
    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType3(rs)
    Else
        
        MsgBox "查询不到该物料信息", vbInformation, "提示"
        Exit Sub

    End If

End Sub

Private Sub QueType4()

    Dim rs         As New ADODB.Recordset

    Dim strCusCode As String

    Dim strSql     As String
    
    If txtKey.Text = "" Then
        MsgBox "请输入客户代码", vbInformation, "提示"
        Exit Sub

    End If
    
    If txtKey2.Text = "" Then
        strSql = "      select '' as '√',cc.客户代码,bb.供应商编号,cc.客户名称,AA.仓库编号,aa.批号,aa.当前存量 ,'' AS 除账数量 FROM erpbase..tblStockNum AA  INNER JOIN  tblSupplierData  bb   ON  bb.供应商编号 = aa.供应商编号  " & "    LEFT JOIN erpdata..tblXCustomer cc ON cc.客户名称 = bb.供应商名称  WHERE aa.仓库编号 = '54'  AND  cc.客户代码 like '%" & UCase(Trim(txtKey.Text)) & "%'  AND aa.当前存量 > 0   "
    Else

        strSql = "      select '' as '√',cc.客户代码,bb.供应商编号,cc.客户名称,AA.仓库编号,aa.批号,aa.当前存量 ,'' AS 除账数量 FROM erpbase..tblStockNum AA  INNER JOIN  tblSupplierData  bb   ON  bb.供应商编号 = aa.供应商编号  " & "    LEFT JOIN erpdata..tblXCustomer cc ON cc.客户名称 = bb.供应商名称  WHERE aa.仓库编号 = '54'  AND  cc.客户代码 like '%" & UCase(Trim(txtKey.Text)) & "%'  AND aa.当前存量 > 0  and  AA.批号 = '" & UCase(Trim(txtKey2.Text)) & "' "

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType4(rs)
    Else
        
        MsgBox "查询不到信息", vbInformation, "提示"
        Exit Sub

    End If

End Sub
Private Sub QueType5()

    Dim rs         As New ADODB.Recordset

    Dim strInv As String
    
    Dim strInv1 As String

    Dim strSql     As String
    
    strInv = Trim$(txtKey.Text)
    
    strInv1 = Trim$(txtKey2.Text)
    
    If txtKey.Text = "" Then
        MsgBox "请输入出货单据", vbInformation, "提示"
        Exit Sub
    End If
    
    If txtKey2.Text = "" Then
        strSql = "select 出货单据,料号,发票号,出货日期,数量,类别,报关单号,品名,手册项号,单位,总价,手册号,AWB#,目的地,货代,退单日期,备注,'' as '√' from erptemp.dbo.ksexport where 出货单据 = '" & strInv & "' and flag = '0' "
    Else

        strSql = "select 出货单据,料号,发票号,出货日期,数量,类别,报关单号,品名,手册项号,单位,总价,手册号,AWB#,目的地,货代,退单日期,备注,'' as '√' from erptemp.dbo.ksexport where 出货单据 = '" & strInv & "' and 料号 = '" & strInv1 & "' and flag = '0'"

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType5(rs)
    Else
        
        MsgBox "查询不到该出口单据信息", vbInformation, "提示"
        Exit Sub

    End If

End Sub

Private Sub QueType6()

    Dim rs         As New ADODB.Recordset

    Dim strInv As String
    
    Dim strInv1 As String

    Dim strSql     As String
    
    strInv = Trim$(txtKey.Text)
    
    strInv1 = Trim$(txtKey2.Text)
    
    If txtKey.Text = "" Then
        MsgBox "请输入采购单号", vbInformation, "提示"
        Exit Sub
    End If
    
    If txtKey2.Text = "" Then
        strSql = "select 采购单号,料号,类别,关务到货数量,标准die,入场日期,发票号,品名,项号,件数,手册号,关税,增值税,报关单号,AWB#,货代,退单日期,备注,id,'' as '√' from erptemp.dbo.ksimport where 采购单号 = '" & strInv & "' and flag = '0' "
    Else

        strSql = "select 采购单号,料号,类别,关务到货数量,标准die,入场日期,发票号,品名,项号,件数,手册号,关税,增值税,报关单号,AWB#,货代,退单日期,备注,id,'' as '√' from erptemp.dbo.ksimport where 采购单号 = '" & strInv & "' and 料号 = '" & strInv1 & "' and flag = '0'"

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '表示有数据了
        Call ListDataType6(rs)
    Else
        
        MsgBox "查询不到该出口单据信息", vbInformation, "提示"
        Exit Sub

    End If

End Sub

Private Sub ListDataType2(rs As ADODB.Recordset)

    Dim i As Long

    With fpS(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4
            .ColWidth(4) = 2
            .CellType = CellTypeCheckBox
        Next
        
    End With

End Sub

Private Sub ListDataType3(rs As ADODB.Recordset)

    Dim i As Long
   
    With fpS(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .ColWidth(1) = 2
            .CellType = CellTypeCheckBox
        Next

    End With

End Sub

Private Sub ListDataType4(rs As ADODB.Recordset)

    Dim i As Long
   
    With fpS(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .ColWidth(1) = 2
            .CellType = CellTypeCheckBox
            
            .Col = 1
            .Lock = False
            
            .Col = 8
            .Lock = False
            
        Next

    End With
    
End Sub


Private Sub ListDataType5(rs As ADODB.Recordset)

 Dim i As Long
  
   
    With fpS(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
     With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 18
            .ColWidth(18) = 2
            .CellType = CellTypeCheckBox
        Next

    End With
    

End Sub


Private Sub ListDataType6(rs As ADODB.Recordset)

 Dim i As Long
  
   
    With fpS(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
     With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 19
            .ColWidth(19) = 4
            .Col = 20
            .ColWidth(20) = 2
            .CellType = CellTypeCheckBox
        Next

    End With
    
 
End Sub
Private Sub ListDataType1(rs As ADODB.Recordset)

    Dim i As Long

    With fpS(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 8
            .ColWidth(8) = 2
            .CellType = CellTypeCheckBox
        Next
        
    End With

End Sub

Private Sub ForAdd()

    If Toolbar1.Buttons(3).Caption = "提交" Then
        
        Select Case cmbCombo1.Text

            Case "合格供应商合格物料"
                ForCommit1
        
            Case "客户出货调仓配置"
                ForCommit2
             
            Case "出口明细表"
                ForCommit5
                
            Case "进口明细表"
                ForCommit6

        End Select
        
        Exit Sub

    End If

    If cmbCombo1.Text = "" Then
        MsgBox "请选择维护类型", vbInformation, "提示"
        Exit Sub

    End If

    Select Case cmbCombo1.Text

        Case "合格供应商合格物料"
            AddType1
        
        Case "客户出货调仓配置"
            AddType2
            
        Case "ERP物料有效期更新"
            MsgBox "该维护类型不支持新增操作", vbInformation, "提示"
            Exit Sub

        Case "出口明细表"
            AddType5

        Case "进口明细表"
            AddType6

    End Select

End Sub

Private Sub AddType1()

    Dim rs     As New ADODB.Recordset

    Dim strMat As String, strMatNo As String

    Dim strSql As String

    If txtKey.Text = "" Then
        MsgBox "请填写要维护的料号", vbInformation, "提示"
        Exit Sub

    End If
    
    fpS(0).MaxRows = 0

    strMat = Trim$(txtKey.Text)
    strMatNo = Get_SqlStr("select 物料编号 from dbo.tblSmainM2 where 料号 = '" & strMat & "'")
    
    If strMatNo = "" Then
        MsgBox "查询不到该料号信息,是否有误", vbInformation, "提示"
        Exit Sub

    End If
    
    strSql = "select '' as 供应商编号,'' as 供应商名称,'" & strMat & "' as 料号, '" & strMatNo & "' as 物料编号"

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    Call ListDataType1(rs)
    
    Toolbar1.Buttons(3).Caption = "提交"
    Toolbar1.Buttons(3).Image = 6
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    
    With fpS(0)
        .Col = 1
        .Lock = False
        .CellType = CellTypeEdit
      
        .Col = 2
        .Lock = False
        .CellType = CellTypeEdit
      
    End With
    
End Sub

Private Sub AddType2()

    Dim rs         As New ADODB.Recordset

    Dim strCusCode As String, strHouse As String

    Dim strSql     As String

    If txtKey.Text = "" Then
        MsgBox "请填写要维护的客户代码", vbInformation, "提示"
        Exit Sub

    End If
    
    fpS(0).MaxRows = 0

    strCusCode = Trim$(txtKey.Text)
 
    strSql = "select '" & strCusCode & "' as 客户代码,'' as 出货单号"

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    Call ListDataType2(rs)
    
    Toolbar1.Buttons(3).Caption = "提交"
    Toolbar1.Buttons(3).Image = 6
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    
    With fpS(0)
        .Col = 2
        .Lock = False
        .CellType = CellTypeEdit
    
    End With
    
End Sub

Private Sub AddType5()

    Dim rs     As New ADODB.Recordset

    Dim i      As Integer

    Dim m      As Integer
    
    Dim strInv As String

    Dim strSql As String

    If txtKey.Text = "" Then
        MsgBox "请填写要维护的出货单据", vbInformation, "提示"
        Exit Sub

    End If
    
    
    fpS(0).MaxRows = 0

    strInv = Trim$(txtKey.Text)
    

    If Get_SqlserverCnt("SELECT * FROM erpdata..tblStockMove A WHERE A.单据编号 = '" & strInv & "'") = 0 Then
        MsgBox "没有此出货单据,请重新输入", vbInformation, "提示"
        Exit Sub

    End If

    strSql = "select b.单据编号 as 单据号码,c.料号,(select distinct 销售发票 from erptemp.dbo.tblBB_CPFH_Invoice  a where  a.发货单号 = b.单据编号) as 发票号,CONVERT(varchar(100), b.操作日期, 23) as 出货日期,SUM(b.实发良品数+b.实发不良数+b.实发制程不良数) as 数量,'' as 类别,'' as 报关单号,'' as 品名,'' as 手册项号,'' as 单位,'' as 总价,'' as 手册号,'' as AWB#,'' as 目的地,'' as 货代,'' as 退单日期,'' as 备注,'' as '√' from   erpdata..tblStockMove b,erpdata..tblSmainM2 c  where    b.单据编号 = '" & strInv & "' and  c.物料编号 = b.物料编号 and c.料号 not in (select distinct 料号 from erptemp.dbo.ksexport where 出货单据 = '" & strInv & "' and flag = '0') group by b.单据编号,c.料号,CONVERT(varchar(100), b.操作日期, 23) "
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    Call ListDataType5(rs)
    
    Toolbar1.Buttons(3).Caption = "提交"
    Toolbar1.Buttons(3).Image = 6
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
    
            For m = 7 To 18
            
                .Col = m
                .Lock = False
      
            Next
    
        Next

        '        For i = 6 To 16
        '            .Col = i
        '            .Lock = False
        '            .CellType = CellTypeEdit
        '        Next
        '
    End With
    
End Sub

Private Sub AddType6()

    Dim rs     As New ADODB.Recordset

    Dim i      As Integer

    Dim m      As Integer
    
    Dim ID     As Integer
    
    Dim strInv As String

    Dim strSql As String

    If txtKey.Text = "" Then
        MsgBox "请填写要维护的采购单号", vbInformation, "提示"
        Exit Sub

    End If
    
    ID = 1
    
    fpS(0).MaxRows = 0

    strInv = Trim$(txtKey.Text)

    If Get_SqlserverCnt("SELECT * FROM erpbase..tblCPurDataSub WHERE 采购单编号 = '" & strInv & "'") = 0 Then
        MsgBox "没有此采购单号,请重新输入", vbInformation, "提示"
        Exit Sub

    End If

    strSql = "SELECT a.采购单编号,b.料号,'' AS 类别,ceiling(sum(a.批准采购数量) - isnull(SUM(c.关务到货数量),0)) as 关务到货数量,'' as 标准die,'' as 入场日期,'' as 发票号,'' as 品名,'' as 项号,'' as 件数,'' as 手册号,'' as 关税,'' as 增值税,'' as 报关单号,'' as AWB#,'' as 货代,'' as 退单日期,'' as 备注,'' as id ,'' as '√' FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b  left join erptemp.dbo.ksimport c on c.料号 = b.料号 and flag = '0' and c.采购单号 = '" & strInv & "' WHERE a.采购单编号 = '" & strInv & "' and a.物料编号 = b.物料编号 GROUP by a.采购单编号,b.料号 "
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    Call ListDataType6(rs)
    
    Toolbar1.Buttons(3).Caption = "提交"
    Toolbar1.Buttons(3).Image = 6
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
    
            For m = 4 To 18
            
                .Col = m
                .Lock = False
      
            Next
            .Col = 20
            .Lock = False
        Next

    End With
    
End Sub

Private Sub ForCommit1()

    Dim strGYSNo   As String

    Dim strGYSName As String

    Dim strMat     As String

    Dim strSql     As String

    Dim strMatNo   As String

    With fpS(0)
        .Row = 1
        .Col = 1

        If .Text = "" Then
            MsgBox "请输入供应商编号", vbInformation, "提示"
            Exit Sub

        End If
    
        strGYSNo = Trim$(.Text)
    
        .Col = 2

        If .Text = "" Then
            MsgBox "请输入供应商名称", vbInformation, "提示"
            Exit Sub

        End If
    
        strGYSName = Trim$(.Text)
    
        .Col = 3
        strMat = Trim$(.Text)
    
        .Col = 4
        strMatNo = Trim$(.Text)

    End With

    AddSql2 ("insert into ERPBASE..tblCG_PassSupplier( 供应商编号,供应商名称,料号, 物料编号,有效否,创建时间) values('" & strGYSNo & "','" & strGYSName & "','" & strMat & "','" & strMatNo & "','1',GetDate())")

    MsgBox "新增成功", vbInformation, "提示"

    Toolbar1.Buttons(3).Caption = "新增"
    Toolbar1.Buttons(3).Image = 2
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery

End Sub

Private Sub ForCommit2()

    Dim strCusCode As String

    Dim strHouse   As String

    Dim strSql     As String
    
    With fpS(0)
        .Row = 1
        
        .Col = 1
        strCusCode = Trim(.Text)
        
        .Col = 2

        If .Text = "" Then
            MsgBox "请输入出货单号", vbInformation, "提示"
            Exit Sub

        End If
    
        strHouse = Trim$(.Text)

        If Get_SqlserverCnt("select * from erpdata..tblStockmovesub where 单据编号 = '" & strHouse & "'") = 0 Then
            MsgBox "查不到该出货单号", vbInformation, "提示"
            Exit Sub

        End If
        
    End With
    
    Dim rs As New ADODB.Recordset
    
    Set rs = Get_SqlserveRs("SELECT b.客户代码,A.单据编号,a.料号,SUM(a.数量) as DIE数量,COUNT(DISTINCT a.流程卡编号)  as 片数量 FROM erpdata..tblStockmovesub A,erpdata..tblStockmove B " & " WHERE A.单据编号 = '" & strHouse & "' AND b.单据编号 = a.单据编号 AND b.序号 = a.单据项次 GROUP BY  A.单据编号, b.客户代码,a.料号")

    With fpS(0)
        
        .MaxRows = 0

        If rs.RecordCount > 0 Then
            Set .DataSource = rs

        End If

    End With
    
    If MsgBox("确认信息是否无误", vbYesNoCancel, "提示") = vbNo Then
        Exit Sub

    End If
    
    AddSql2 ("insert into erptemp..tbltransfer(CUSTOMER,WAREHOUSE,CREATE_DATE,CREATE_BY,LAST_UPDATE_DATE,LAST_UPDATE_BY, FLAG) values('" & strCusCode & "','" & strHouse & "',CONVERT(varchar(100), GETDATE(), 23),'" & gUserName & "','','',1)")

    MsgBox "新增成功", vbInformation, "提示"

    Toolbar1.Buttons(3).Caption = "新增"
    Toolbar1.Buttons(3).Image = 2
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery

End Sub

Private Sub ForCommit5()

    Dim strInv1  As String

    Dim strInv2  As String

    Dim strInv3  As String
    
    Dim strInv4  As String

    Dim strInv5  As String

    Dim strInv6  As String

    Dim strInv7  As String

    Dim strInv8  As String

    Dim strInv9  As String

    Dim strInv10 As String

    Dim strInv11 As String

    Dim strInv12 As String

    Dim strInv13 As String

    Dim strInv14 As String

    Dim strInv15 As String

    Dim strInv16 As String
    
    Dim strInv17 As String

    Dim strSql   As String

    Dim i        As Integer

    Dim j        As Integer

    Dim bFlag    As Boolean

    bFlag = False

    With fpS(0)
    
        '        For i = 1 To .MaxRows
        '            .Row = i
        '
        '            For m = 6 To 17
        '
        '                .Col = m
        '                .Lock = False
        '            Next
        '
        '        Next

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
        
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 18
    
            j = 0

            If .Text = "1" Then
                j = j + 1
                bFlag = True
                .Col = 1

                If .Text = "" Then
                    MsgBox "请输入出货单据", vbInformation, "提示"
                    Exit Sub

                End If
    
                strInv1 = Trim$(.Text)
    
                .Col = 2

                If .Text = "" Then
                    MsgBox "请输入料号", vbInformation, "提示"
                    Exit Sub

                End If
    
                strInv2 = Trim$(.Text)
                
                If Get_SqlserverCnt("select * from erptemp.dbo.ksexport where 出货单据 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and flag = '0'") > 0 Then
                    MsgBox "该笔资料已经新增过", vbInformation, "提示"
                    Exit Sub

                End If
    
                .Col = 3
                strInv3 = Trim$(.Text)
                
                .Col = 4
                strInv4 = Trim$(.Text)
        
                .Col = 5
                
                strInv5 = Trim$(.Text)
        
                .Col = 6
                lab2.Visible = True
                Combo1.Visible = True

                If Combo1.Text = "" Then
                    MsgBox "请输入类别", vbInformation, "提示"
                    Exit Sub

                End If
                
                .Text = Combo1.Text
                
                If .Text = "" Then
                    MsgBox "请输入类别", vbInformation, "提示"
                    Exit Sub

                End If

                strInv6 = Trim$(.Text)
        
                .Col = 7
                strInv7 = Trim$(.Text)
        
                .Col = 8
                strInv8 = Trim$(.Text)
        
                .Col = 9
                strInv9 = Trim$(.Text)
        
                .Col = 10
                strInv10 = Trim$(.Text)
        
                .Col = 11
                strInv11 = Trim$(.Text)
        
                .Col = 12
                strInv12 = Trim$(.Text)
        
                .Col = 13
                strInv13 = Trim$(.Text)
        
                .Col = 14
                strInv14 = Trim$(.Text)
        
                .Col = 15
                strInv15 = Trim$(.Text)
        
                .Col = 16
                strInv16 = Trim$(.Text)
                
                .Col = 17
                strInv17 = Trim$(.Text)

                AddSql2 ("insert into erptemp.dbo.ksexport( 出货单据,料号,发票号,出货日期,数量,类别,报关单号,品名,手册项号,单位,总价,手册号,AWB#,目的地,货代,退单日期,备注,键入时间,修改状态,修改时间,删除时间,flag) values('" & strInv1 & "','" & strInv2 & "','" & strInv3 & "','" & strInv4 & "','" & strInv5 & "','" & strInv6 & "','" & strInv7 & "','" & strInv8 & "','" & strInv9 & "','" & strInv10 & "','" & strInv11 & "','" & strInv12 & "','" & strInv13 & "','" & strInv14 & "','" & strInv15 & "','" & strInv16 & "','" & strInv17 & "',GetDate(),NULL,NULL,NULL,'0')")

            End If

            '
            '            If bFlag = False And j = 0 Then
            '                MsgBox "请选择要新增的行", vbInformation, "提示"
            '                Exit Sub
            '
            '            End If

        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "请选择要新增的行", vbInformation, "提示"
            Exit Sub
            
        End If

    End With
    
    MsgBox "新增成功", vbInformation, "提示"
    Toolbar1.Buttons(3).Caption = "新增"
    Toolbar1.Buttons(3).Image = 2
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery

End Sub

Private Sub ForCommit6()

    Dim strInv1  As String

    Dim strInv2  As String

    Dim strInv3  As String
    
    Dim strInv4  As Integer

    Dim strInv5  As String

    Dim strInv6  As String

    Dim strInv7  As String

    Dim strInv8  As String

    Dim strInv9  As String

    Dim strInv10 As String

    Dim strInv11 As String

    Dim strInv12 As String

    Dim strInv13 As String

    Dim strInv14 As String

    Dim strInv15 As String

    Dim strInv16 As String
    
    Dim strInv17 As String
    
    Dim strInv18 As String

    Dim strInv19 As Integer
    
    Dim strID    As Integer
    
    Dim strid1   As Integer

    Dim strNo1   As Integer

    Dim strNo2   As Integer

    Dim strNo3   As Integer

    Dim strSql   As String

    Dim i        As Integer

    Dim j        As Integer

    Dim bFlag    As Boolean

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        strID = 1

        For i = 1 To .MaxRows
            .Row = i
            .Col = 20
    
            j = 0

            If .Text = "1" Then
                j = j + 1
                bFlag = True
                .Col = 1

                If .Text = "" Then
                    MsgBox "请输入采购单号", vbInformation, "提示"
                    Exit Sub

                End If
    
                strInv1 = Trim$(.Text)
    
                .Col = 2

                If .Text = "" Then
                    MsgBox "请输入料号", vbInformation, "提示"
                    Exit Sub

                End If
    
                strInv2 = Trim$(.Text)
    
                .Col = 3
                
                lab2.Visible = True
                Combo1.Visible = True

                If Combo1.Text = "" Then
                    MsgBox "请输入类别", vbInformation, "提示"
                    Exit Sub

                End If
                
                .Text = Combo1.Text
                
                If .Text = "" Then
                    MsgBox "请输入类别", vbInformation, "提示"
                    Exit Sub

                End If
                
                strInv3 = Trim$(.Text)
                              
                .Col = 4
                
                strInv4 = Trim$(.Text)
                
                strNo1 = Get_SqlStr("SELECT ceiling(isnull(SUM(a.批准采购数量),0)) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.采购单编号 = '" & strInv1 & "' and a.物料编号 = b.物料编号 and b.料号 = '" & strInv2 & "' ")
                
                strNo2 = Get_SqlStr("SELECT ceiling(isnull(SUM(关务到货数量),0)) FROM erptemp.dbo.ksimport where 采购单号 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and flag = '0'")
                
                strNo3 = strNo1 - strNo2
                
                If strInv4 > strNo3 Then
                    MsgBox "该笔料号" & strInv2 & "批准采购数量: " & strNo1 & ",已经维护关务数量：" & strNo2 & ",最大数量只能维护：" & strNo3 & "", vbInformation, "提示"
                    Exit Sub

                End If
                
                .Col = 5
                strInv5 = Trim$(.Text)
        
                .Col = 6
                strInv6 = Trim$(.Text)
        
                .Col = 7
                strInv7 = Trim$(.Text)
        
                .Col = 8
                strInv8 = Trim$(.Text)
        
                .Col = 9
                strInv9 = Trim$(.Text)
        
                .Col = 10
                strInv10 = Trim$(.Text)
        
                .Col = 11
                strInv11 = Trim$(.Text)
        
                .Col = 12
                strInv12 = Trim$(.Text)
        
                .Col = 13
                strInv13 = Trim$(.Text)
        
                .Col = 14
                strInv14 = Trim$(.Text)
        
                .Col = 15
                strInv15 = Trim$(.Text)
        
                .Col = 16
                strInv16 = Trim$(.Text)
                
                .Col = 17
                strInv17 = Trim$(.Text)
                
                .Col = 18
                strInv18 = Trim$(.Text)
                
                .Col = 19
                
                If Get_SqlserverCnt("select * from erptemp.dbo.ksimport where 采购单号 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and flag = '0'") > 0 Then
                
                    strid1 = Get_SqlStr(" select MAX(id) from erptemp.dbo.ksimport where 采购单号 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and flag = '0'")
                    
                    strID = strid1 + 1
              
                End If
        
                .Text = strID
                
                strInv19 = Trim$(.Text)

                AddSql2 ("insert into erptemp.dbo.ksimport( 采购单号,料号,类别,关务到货数量,标准die,入场日期,发票号,品名,项号,件数,手册号,关税,增值税,报关单号,AWB#,货代,退单日期,备注,id,键入时间,修改状态,修改时间,删除时间,flag) values('" & strInv1 & "','" & strInv2 & "','" & strInv3 & "','" & strInv4 & "','" & strInv5 & "','" & strInv6 & "','" & strInv7 & "','" & strInv8 & "','" & strInv9 & "','" & strInv10 & "','" & strInv11 & "','" & strInv12 & "','" & strInv13 & "','" & strInv14 & "','" & strInv15 & "','" & strInv16 & "','" & strInv17 & "','" & strInv18 & "','" & strInv19 & "',GetDate(),NULL,NULL,NULL,'0')")

            End If

            '           strid = strid + 1
        Next
        
        'j = 0 获取不到用户需要输入的资料
        If bFlag = False And j = 0 Then
            MsgBox "请选择要新增的行", vbInformation, "提示"
            Exit Sub
            
        End If
        
        lab2.Visible = Flase
        Combo1.Visible = Flase
        '        Combol.Text = ""

    End With
    
    MsgBox "新增成功", vbInformation, "提示"
    Toolbar1.Buttons(3).Caption = "新增"
    Toolbar1.Buttons(3).Image = 2
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery

End Sub


Private Sub ForMod1()

    Dim i As Integer

    If Toolbar1.Buttons(5).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                
                .Col = 2
                .Lock = False
        
                .Col = 3
                .Lock = False
        
                .Col = 8
                .Lock = False
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "提交"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    Dim bFlag As Boolean

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 8

            If .Text = "1" Then
                bFlag = True
           
            End If

        Next

    End With

    If bFlag = False Then
        MsgBox "请选择要修改的行", vbInformation, "提示"
        Exit Sub

    End If
    
    Dim strGYSNo   As String

    Dim strMat     As String

    Dim strGYSName As String

    Dim strno      As String
    
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 8

            If .Text = "1" Then
                .Col = 1
                strno = Trim$(.Text)
                
                .Col = 2
                strGYSNo = Trim$(.Text)
                
                .Col = 3
                strGYSName = Trim$(.Text)
                
                AddSql2 ("update ERPBASE..tblCG_PassSupplier set 供应商编号 = '" & strGYSNo & "', 供应商名称 = '" & strGYSName & "' where 序号 = '" & strno & "'     ")
            
            End If
            
        Next
    
    End With
    
    MsgBox "修改成功", vbInformation, "提示"

    Toolbar1.Buttons(5).Caption = "修改"
    Toolbar1.Buttons(5).Image = 3
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery
    
End Sub

Private Sub ForMod2()

    Dim i As Integer

    If Toolbar1.Buttons(5).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                
                .Col = 2
                .Lock = False
                
                .Col = 3
                .Lock = False
                
                .Col = 4
                .Lock = False
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "提交"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    Dim bFlag As Boolean

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4

            If .Text = "1" Then
                bFlag = True
           
            End If

        Next

    End With

    If bFlag = False Then
        MsgBox "请选择要修改的行", vbInformation, "提示"
        Exit Sub

    End If
    
    Dim strCusCode As String

    Dim strHouse   As String
    
    Dim strflag    As String

    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4

            If .Text = "1" Then
                .Col = 1
                strCusCode = Trim$(.Text)
                
                .Col = 2
                strHouse = Trim$(.Text)
                
                .Col = 3
                strflag = Trim$(.Text)
                
                AddSql2 ("update erptemp..tbltransfer set Warehouse = '" & strHouse & "' , last_update_date = CONVERT(varchar(100), GETDATE(), 23), last_update_by = '" & gUserName & "', flag = '" & strflag & "' where Customer = '" & strCusCode & "' ")
            
            End If
            
        Next
    
    End With
    
    MsgBox "修改成功", vbInformation, "提示"

    Toolbar1.Buttons(5).Caption = "修改"
    Toolbar1.Buttons(5).Image = 3
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery
    
End Sub

Private Sub ForMod3()

    Dim i As Integer

    If Toolbar1.Buttons(5).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                
                .Col = 1
                .Lock = False
                .BackColor = vbGreen
                
                .Col = 8
                .Lock = False
                .BackColor = vbGreen
            
                .Col = 10
                .Lock = False
                .BackColor = vbGreen
            
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "提交"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    Dim bFlag As Boolean

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Text = "1" Then
                bFlag = True
           
            End If

        Next

    End With

    If bFlag = False Then
        MsgBox "请选择要修改的行", vbInformation, "提示"
        Exit Sub

    End If
    
    Dim strID      As String

    Dim strNewDate As String
    Dim strKW As String
   
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Text = "1" Then
                .Col = 2
                strID = Trim$(.Text)
                
                .Col = 8
                strNewDate = UCase(Trim$(.Text))
                
                .Col = 10
                strKW = UCase(Trim$(.Text))
                
                AddSql2 ("update erpbase.dbo.tblStockNum  set 有效期至 = '" & strNewDate & "', 库位 = '" & strKW & "'  where id = '" & strID & "' ")
            
            End If
            
        Next
    
    End With
    
    MsgBox "修改成功", vbInformation, "提示"

    Toolbar1.Buttons(5).Caption = "修改"
    Toolbar1.Buttons(5).Image = 3
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery
    
End Sub

Private Sub ForMod5()

    Dim i        As Integer

    Dim m        As Integer

    Dim j        As Integer

    Dim strInv1  As String

    Dim strInv2  As String

    Dim strInv3  As String

    Dim strInv4  As String

    Dim strInv5  As String

    Dim strInv6  As String

    Dim strInv7  As String

    Dim strInv8  As String

    Dim strInv9  As String

    Dim strInv10 As String

    Dim strInv11 As String

    Dim strInv12 As String

    Dim strInv13 As String

    Dim strInv14 As String

    Dim strInv15 As String

    Dim strInv16 As String
    
    Dim strInv17 As String

    Dim strtime  As String
    
    Dim bFlag    As Boolean

    If Toolbar1.Buttons(5).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
    
                For m = 7 To 18
            
                    .Col = m
                    .Lock = False
      
                Next
    
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "提交"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 18
    
            j = 0

            If .Text = "1" Then
            
                j = j + 1
                bFlag = True
                .Col = 1
                strInv1 = Trim$(.Text)
    
                .Col = 2
                strInv2 = Trim$(.Text)
    
                .Col = 3
                strInv3 = Trim$(.Text)
                
                .Col = 4
                strInv4 = Trim$(.Text)
        
                .Col = 5
                strInv5 = Trim$(.Text)
        
                .Col = 6
                strInv6 = Trim$(.Text)
        
                .Col = 7
                strInv7 = Trim$(.Text)
        
                .Col = 8
                strInv8 = Trim$(.Text)
        
                .Col = 9
                strInv9 = Trim$(.Text)
        
                .Col = 10
                strInv10 = Trim$(.Text)
        
                .Col = 11
                strInv11 = Trim$(.Text)
        
                .Col = 12
                strInv12 = Trim$(.Text)
        
                .Col = 13
                strInv13 = Trim$(.Text)
        
                .Col = 14
                strInv14 = Trim$(.Text)
        
                .Col = 15
                strInv15 = Trim$(.Text)
        
                .Col = 16
                strInv16 = Trim$(.Text)
                
                .Col = 17
                strInv17 = Trim$(.Text)
    
                strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
                
                
                AddSql2 ("insert into erptemp.dbo.ksexport (出货单据,料号,发票号,出货日期,数量,类别,报关单号,品名,手册项号,单位,总价,手册号,AWB#,目的地,货代,退单日期,备注,键入时间,修改状态,修改时间,删除时间,flag) SELECT 出货单据,料号,发票号,出货日期,数量,类别,报关单号,品名,手册项号,单位,总价,手册号,AWB#,目的地,货代,退单日期,备注,键入时间,'修改前',修改时间,删除时间,'2' FROM erptemp.dbo.ksexport WHERE 出货单据 = '" & strInv1 & "'  AND 料号 =  '" & strInv2 & "' AND  flag = '0'")
                AddSql2 ("update erptemp.dbo.ksexport set 报关单号 =  '" & strInv7 & "',品名 =  '" & strInv8 & "',手册项号 =  '" & strInv9 & "',单位 =  '" & strInv10 & "',总价 =  '" & strInv11 & "',手册号 =  '" & strInv12 & "',AWB# =  '" & strInv13 & "',目的地 =  '" & strInv14 & "',货代 =  '" & strInv15 & "',退单日期 =  '" & strInv16 & "',备注 =  '" & strInv17 & "',修改状态 = '修改后',修改时间 = '" & strtime & "' where 出货单据 = '" & strInv1 & "' and flag = '0' and 料号  = '" & strInv2 & "' ")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "请选择要修改的行", vbInformation, "提示"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "修改成功", vbInformation, "提示"

    Toolbar1.Buttons(5).Caption = "修改"
    Toolbar1.Buttons(5).Image = 3
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery
    
End Sub

'Private Sub ForMod6()
'
'    Dim i        As Integer
'
'    Dim m        As Integer
'
'    Dim j        As Integer
'
'    Dim strInv1  As String
'
'    Dim strInv2  As String
'
'    Dim strInv3  As String
'
'    Dim strInv4  As Integer
'
'    Dim strInv5  As String
'
'    Dim strInv6  As String
'
'    Dim strInv7  As String
'
'    Dim strInv8  As String
'
'    Dim strInv9  As String
'
'    Dim strInv10 As String
'
'    Dim strInv11 As String
'
'    Dim strInv12 As String
'
'    Dim strInv13 As String
'
'    Dim strInv14 As String
'
'    Dim strInv15 As String
'
'    Dim strInv16 As String
'
'    Dim strInv17 As String
'
'    Dim strInv18 As String
'
'    Dim strInv19 As Integer
'
'    Dim strTime  As String
'
'    Dim bFlag    As Boolean
'
'    Dim strNo1   As Integer
'
'    Dim strNo2   As Integer
'
'    Dim strNo3   As Integer
'
'    If Toolbar1.Buttons(5).Caption <> "提交" Then
'
'        With fps(0)
'
'            For i = 1 To .MaxRows
'                .Row = i
'
'                For m = 4 To 18
'
'                    .Col = m
'                    .Lock = False
'
'                Next
'                .Col = 20
'                .Lock = False
'
'            Next
'
'        End With
'
'        Toolbar1.Buttons(5).Caption = "提交"
'        Toolbar1.Buttons(5).Image = 6
'        Toolbar1.Buttons(1).Enabled = False
'        Toolbar1.Buttons(3).Enabled = False
'        Toolbar1.Buttons(7).Enabled = False
'        Exit Sub
'
'    End If
'
'    bFlag = False
'
'    With fps(0)
'
'        If .MaxRows = 0 Then
'            MsgBox "没有数据", vbInformation, "提示"
'            Exit Sub
'
'        End If
'
'        For i = 1 To .MaxRows
'            .Row = i
'            .Col = 20
'
'            j = 0
'
'            If .Text = "1" Then
'
'                j = j + 1
'                bFlag = True
'                .Col = 1
'                strInv1 = Trim$(.Text)
'
'                .Col = 2
'                strInv2 = Trim$(.Text)
'
'                .Col = 3
'                strInv3 = Trim$(.Text)
'
'                .Col = 4
'                strInv4 = Trim$(.Text)
'
'                strNo1 = Get_SqlStr("SELECT ceiling(isnull(SUM(a.批准采购数量),0)) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.采购单编号 = '" & strInv1 & "' and a.物料编号 = b.物料编号 and b.料号 = '" & strInv2 & "' ")
'
'                strNo2 = Get_SqlStr("SELECT ceiling(isnull(SUM(关务到货数量),0)) FROM erptemp.dbo.ksimport where 采购单号 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and flag = '0'")
'
'                strNo3 = strNo1 - strNo2
'
'                If strInv4 > strNo3 Then
'                    MsgBox "该笔料号" & strInv2 & "批准采购数量: " & strNo1 & ",已经维护关务数量：" & strNo2 & ",最大数量只能维护：" & strNo3 & "", vbInformation, "提示"
'                    Exit Sub
'
'                End If
'
'                .Col = 5
'                strInv5 = Trim$(.Text)
'
'                .Col = 6
'                strInv6 = Trim$(.Text)
'
'                .Col = 7
'                strInv7 = Trim$(.Text)
'
'                .Col = 8
'                strInv8 = Trim$(.Text)
'
'                .Col = 9
'                strInv9 = Trim$(.Text)
'
'                .Col = 10
'                strInv10 = Trim$(.Text)
'
'                .Col = 11
'                strInv11 = Trim$(.Text)
'
'                .Col = 12
'                strInv12 = Trim$(.Text)
'
'                .Col = 13
'                strInv13 = Trim$(.Text)
'
'                .Col = 14
'                strInv14 = Trim$(.Text)
'
'                .Col = 15
'                strInv15 = Trim$(.Text)
'
'                .Col = 16
'                strInv16 = Trim$(.Text)
'
'                .Col = 17
'                strInv17 = Trim$(.Text)
'
'                .Col = 18
'                strInv18 = Trim$(.Text)
'
'                .Col = 19
'                strInv19 = Trim$(.Text)
'
'                strTime = Format(Now, "yyyy-mm-dd hh:mm:ss")
'
'                AddSql2 ("insert into erptemp.dbo.ksimport(采购单号,料号,类别,关务到货数量,标准die,入场日期,发票号,品名,项号,件数,手册号,关税,增值税,报关单号,AWB#,货代,退单日期,备注,id,键入时间,修改状态,修改时间,删除时间,flag) SELECT 采购单号,料号,类别,关务到货数量,标准die,入场日期,发票号,品名,项号,件数,手册号,关税,增值税,报关单号,AWB#,货代,退单日期,备注,id,键入时间,'修改前',修改时间,删除时间,'2' FROM erptemp.dbo.ksimport WHERE 采购单号 = '" & strInv1 & "'  AND 料号 =  '" & strInv2 & "' AND id =  '" & strInv19 & "'  AND  flag = '0'")
'
'                AddSql2 ("update erptemp.dbo.ksimport set 关务到货数量 = '" & strInv4 & "',标准die =  '" & strInv5 & "',入场日期 =  '" & strInv6 & "',发票号 =  '" & strInv7 & "',品名 =  '" & strInv8 & "',项号 =  '" & strInv9 & "',件数 =  '" & strInv10 & "',手册号 =  '" & strInv11 & "',关税 =  '" & strInv12 & "',增值税 =  '" & strInv13 & "',报关单号 =  '" & strInv14 & "',AWB#  =  '" & strInv15 & "',货代 =  '" & strInv16 & "', 退单日期 =  '" & strInv17 & "',备注 =  '" & strInv18 & "',修改状态 = '修改后',修改时间 = '" & strTime & "' where 采购单号 = '" & strInv1 & "' and flag = '0' and 料号  = '" & strInv2 & "' and id =  '" & strInv19 & "' ")
'
'            End If
'
'        Next
'
'        If bFlag = False And j = 0 Then
'            MsgBox "请选择要修改的行", vbInformation, "提示"
'            Exit Sub
'
'        End If
'
'    End With
'
'    MsgBox "修改成功", vbInformation, "提示"
'
'    Toolbar1.Buttons(5).Caption = "修改"
'    Toolbar1.Buttons(5).Image = 3
'    Toolbar1.Buttons(1).Enabled = True
'    Toolbar1.Buttons(3).Enabled = True
'    Toolbar1.Buttons(7).Enabled = True
'
'    ForQuery
'
'End Sub

Private Sub ForMod6()

    Dim i        As Integer

    Dim m        As Integer

    Dim j        As Integer

    Dim strInv1  As String

    Dim strInv2  As String

    Dim strInv3  As String

    Dim strInv4  As Integer

    Dim strInv5  As String

    Dim strInv6  As String

    Dim strInv7  As String

    Dim strInv8  As String

    Dim strInv9  As String

    Dim strInv10 As String

    Dim strInv11 As String

    Dim strInv12 As String

    Dim strInv13 As String

    Dim strInv14 As String

    Dim strInv15 As String

    Dim strInv16 As String

    Dim strInv17 As String

    Dim strInv18 As String

    Dim strInv19 As Integer

    Dim strtime  As String
    
    Dim bFlag    As Boolean
    
    Dim strNo1   As Integer

    Dim strNo2   As Integer

    Dim strNo3   As Integer

    If Toolbar1.Buttons(5).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
    
                For m = 4 To 18
            
                    .Col = m
                    .Lock = False
      
                Next
                .Col = 20
                .Lock = False
    
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "提交"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 20
    
            j = 0

            If .Text = "1" Then
            
                j = j + 1
                bFlag = True
                .Col = 1
                strInv1 = Trim$(.Text)
    
                .Col = 2
                strInv2 = Trim$(.Text)
    
                .Col = 3
                strInv3 = Trim$(.Text)
                
                .Col = 4
                strInv4 = Trim$(.Text)
           
        
                .Col = 5
                strInv5 = Trim$(.Text)
        
                .Col = 6
                strInv6 = Trim$(.Text)
        
                .Col = 7
                strInv7 = Trim$(.Text)
        
                .Col = 8
                strInv8 = Trim$(.Text)
        
                .Col = 9
                strInv9 = Trim$(.Text)
        
                .Col = 10
                strInv10 = Trim$(.Text)
        
                .Col = 11
                strInv11 = Trim$(.Text)
        
                .Col = 12
                strInv12 = Trim$(.Text)
        
                .Col = 13
                strInv13 = Trim$(.Text)
        
                .Col = 14
                strInv14 = Trim$(.Text)
        
                .Col = 15
                strInv15 = Trim$(.Text)
        
                .Col = 16
                strInv16 = Trim$(.Text)
                
                .Col = 17
                strInv17 = Trim$(.Text)
                
                .Col = 18
                strInv18 = Trim$(.Text)
                
                .Col = 19
                strInv19 = Trim$(.Text)
                
                    
                strNo1 = Get_SqlStr("SELECT ceiling(isnull(SUM(a.批准采购数量),0)) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.采购单编号 = '" & strInv1 & "' and a.物料编号 = b.物料编号 and b.料号 = '" & strInv2 & "' ")
                
                strNo2 = Get_SqlStr("SELECT ceiling(isnull(SUM(关务到货数量),0)) FROM erptemp.dbo.ksimport where 采购单号 = '" & strInv1 & "' and 料号 = '" & strInv2 & "' and id <> '" & strInv19 & "' and flag = '0'")
                
                strNo3 = strNo1 - strNo2
                
                If strInv4 > strNo3 Then
                    MsgBox "该笔料号" & strInv2 & "批准采购数量: " & strNo1 & ",已经维护关务数量：" & strNo2 & ",最大数量只能维护：" & strNo3 & "", vbInformation, "提示"
                    Exit Sub

                End If
                
                strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
                
                AddSql2 ("insert into erptemp.dbo.ksimport(采购单号,料号,类别,关务到货数量,标准die,入场日期,发票号,品名,项号,件数,手册号,关税,增值税,报关单号,AWB#,货代,退单日期,备注,id,键入时间,修改状态,修改时间,删除时间,flag) SELECT 采购单号,料号,类别,关务到货数量,标准die,入场日期,发票号,品名,项号,件数,手册号,关税,增值税,报关单号,AWB#,货代,退单日期,备注,id,键入时间,'修改前',修改时间,删除时间,'2' FROM erptemp.dbo.ksimport WHERE 采购单号 = '" & strInv1 & "'  AND 料号 =  '" & strInv2 & "' AND id =  '" & strInv19 & "'  AND  flag = '0'")
                
                AddSql2 ("update erptemp.dbo.ksimport set 关务到货数量 = '" & strInv4 & "',标准die =  '" & strInv5 & "',入场日期 =  '" & strInv6 & "',发票号 =  '" & strInv7 & "',品名 =  '" & strInv8 & "',项号 =  '" & strInv9 & "',件数 =  '" & strInv10 & "',手册号 =  '" & strInv11 & "',关税 =  '" & strInv12 & "',增值税 =  '" & strInv13 & "',报关单号 =  '" & strInv14 & "',AWB#  =  '" & strInv15 & "',货代 =  '" & strInv16 & "', 退单日期 =  '" & strInv17 & "',备注 =  '" & strInv18 & "',修改状态 = '修改后',修改时间 = '" & strtime & "' where 采购单号 = '" & strInv1 & "' and flag = '0' and 料号  = '" & strInv2 & "' and id =  '" & strInv19 & "' ")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "请选择要修改的行", vbInformation, "提示"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "修改成功", vbInformation, "提示"

    Toolbar1.Buttons(5).Caption = "修改"
    Toolbar1.Buttons(5).Image = 3
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery
    
End Sub

Private Sub ForDel1()

    Dim i As Integer

    If Toolbar1.Buttons(7).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
            
                .Col = 8
                .Lock = False
            Next
        
        End With
    
        Toolbar1.Buttons(7).Caption = "提交"
        Toolbar1.Buttons(7).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Exit Sub

    End If
    
    Dim bFlag As Boolean

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 8

            If .Text = "1" Then
                bFlag = True
           
            End If

        Next

    End With

    If bFlag = False Then
        MsgBox "请选择要删除的行", vbInformation, "提示"
        Exit Sub

    End If
    
    Dim strno As String
    
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 8

            If .Text = "1" Then
                .Col = 1
                strno = Trim$(.Text)
                
                AddSql2 ("delete from ERPBASE..tblCG_PassSupplier where 序号 = '" & strno & "'  ")
            
            End If
            
        Next
    
    End With
    
    MsgBox "删除成功", vbInformation, "提示"

    Toolbar1.Buttons(7).Caption = "删除"
    Toolbar1.Buttons(7).Image = 5
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(5).Enabled = True

    ForQuery
    
End Sub

Private Sub ForDel2()

    Dim i As Integer

    If Toolbar1.Buttons(7).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
            
                .Col = 4
                .Lock = False
            Next
        
        End With
    
        Toolbar1.Buttons(7).Caption = "提交"
        Toolbar1.Buttons(7).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Exit Sub

    End If
    
    Dim bFlag As Boolean

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4

            If .Text = "1" Then
                bFlag = True
           
            End If

        Next

    End With

    If bFlag = False Then
        MsgBox "请选择要删除的行", vbInformation, "提示"
        Exit Sub

    End If
    
    Dim strCusCode As String

    Dim strFhdh    As String
    
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4

            If .Text = "1" Then
                .Col = 1
                strCusCode = Trim$(.Text)
                
                .Col = 2
                strFhdh = Trim(.Text)
                
                AddSql2 ("delete from erptemp..tbltransfer where customer = '" & strCusCode & "'   and  warehouse = '" & strFhdh & "'    ")
            
            End If
            
        Next
    
    End With
    
    MsgBox "删除成功", vbInformation, "提示"

    Toolbar1.Buttons(7).Caption = "删除"
    Toolbar1.Buttons(7).Image = 5
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(5).Enabled = True

    ForQuery
    
End Sub

Private Sub cmdCommand1_Click()

    Dim i     As Integer

    Dim bFlag As Boolean

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Text = "1" Then
                bFlag = True
                .Row = i
                .Col = 8
                If Trim(.Text) = "" Then
                    MsgBox "请填写除帐数量", vbInformation, "提示"
                    Exit Sub
                    
                End If
           
            End If

        Next

    End With

    If bFlag = False Then
        MsgBox "请选择要除账的行", vbInformation, "提示"
        Exit Sub

    End If
    
    Dim intQtyN As Long

    Dim intQtyU As Long

    Dim strPt   As String

    Dim strSup  As String
   
    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Text = "1" Then
            
                .Col = 3
                strSup = Trim$(.Text)
            
                .Col = 6
                strPt = Trim$(.Text)
                
                .Col = 7
                intQtyN = Trim$(.Text)
                
                .Col = 8
                intQtyU = Trim$(.Text)
                
                If intQtyU <= intQtyN Then
                    
                    AddSql2 ("UPDATE erpbase..tblStockNum SET 当前存量 =  " & intQtyN & " -  " & intQtyU & "  WHERE 仓库编号 = '54'  AND 供应商编号 = '" & strSup & " '  AND  批号 = '" & strPt & "'")
                Else
                    MsgBox "除账数量大于库存数量", vbInformation, "提示"
                    Exit Sub

                End If

            End If
            
        Next
    
    End With
    
    MsgBox "除账成功", vbInformation, "提示"

    ForQuery
 
End Sub

Private Sub ForDel5()

    Dim i       As Integer

    Dim j       As Integer
    
    Dim bFlag   As Boolean
    
    Dim strInv1 As String

    Dim strInv2 As String

    Dim strtime As String

    If Toolbar1.Buttons(7).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                
                .Col = 18
                .Lock = False
              
            Next
        
        End With
    
        Toolbar1.Buttons(7).Caption = "提交"
        Toolbar1.Buttons(7).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Exit Sub

    End If

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 18
    
            j = 0

            If .Text = "1" Then
            
                j = j + 1
                bFlag = True
                
                .Col = 1
                strInv1 = Trim$(.Text)
                .Col = 2
                
                strInv2 = Trim$(.Text)
                
                strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
                AddSql2 ("update erptemp.dbo.ksexport set flag = '1',删除时间  = '" & strtime & "' where 出货单据 = '" & strInv1 & "' and  料号 = '" & strInv2 & "' and flag = '0'")

            End If

        Next

    End With

    If bFlag = False And j = 0 Then
        MsgBox "请选择要删除的行", vbInformation, "提示"
        Exit Sub

    End If
    
    MsgBox "删除成功", vbInformation, "提示"

    Toolbar1.Buttons(7).Caption = "删除"
    Toolbar1.Buttons(7).Image = 5
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(5).Enabled = True

    ForQuery
    
End Sub


Private Sub ForDel6()

    Dim i       As Integer

    Dim j       As Integer
    
    Dim bFlag   As Boolean
    
    Dim strInv1 As String

    Dim strInv2 As String
    Dim strInv3 As String

    Dim strtime As String

    If Toolbar1.Buttons(7).Caption <> "提交" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                
                .Col = 20
                .Lock = False
              
            Next
        
        End With
    
        Toolbar1.Buttons(7).Caption = "提交"
        Toolbar1.Buttons(7).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Exit Sub

    End If

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "没有数据", vbInformation, "提示"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 20
    
            j = 0

            If .Text = "1" Then
            
                j = j + 1
                bFlag = True
                
                .Col = 1
                strInv1 = Trim$(.Text)
                .Col = 2
                
                strInv2 = Trim$(.Text)
                
                .Col = 19
                
                strInv19 = Trim$(.Text)
                
                strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
                AddSql2 ("update erptemp.dbo.ksimport set flag = '1',删除时间  = '" & strtime & "' where 采购单号 = '" & strInv1 & "' and  料号 = '" & strInv2 & "' and id = '" & strInv19 & "' and flag = '0'")

            End If

        Next

    End With

    If bFlag = False And j = 0 Then
        MsgBox "请选择要删除的行", vbInformation, "提示"
        Exit Sub

    End If
    
    MsgBox "删除成功", vbInformation, "提示"

    Toolbar1.Buttons(7).Caption = "删除"
    Toolbar1.Buttons(7).Image = 5
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(5).Enabled = True

    ForQuery
    
End Sub
















