VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmApplyWO 
   Caption         =   "PMC 下工单"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   16155
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0FF&
      Caption         =   "To金碟"
      Height          =   480
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdM 
      BackColor       =   &H0080C0FF&
      Caption         =   "手工建立工单"
      Height          =   480
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdBom 
      BackColor       =   &H00FFC0FF&
      Caption         =   "尾批绑定"
      Height          =   480
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "清空数据"
      Height          =   480
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "导出Detail"
      Height          =   480
      Left            =   10770
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "导出Header"
      Height          =   480
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton ComSave 
      BackColor       =   &H00FFC0FF&
      Caption         =   "保存工单"
      Height          =   480
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   9480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "工单Detail"
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   18615
      Begin VB.CommandButton CmdNewQuery 
         Caption         =   "搜索"
         Height          =   360
         Left            =   2040
         TabIndex        =   74
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtQueryLotID 
         Height          =   285
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox ChkAll 
         Height          =   255
         Left            =   13320
         TabIndex        =   57
         Top             =   120
         Width           =   255
      End
      Begin VB.ListBox Lst 
         Height          =   5235
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   52
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">>"
         Height          =   360
         Left            =   2040
         TabIndex        =   51
         Top             =   1920
         Width           =   615
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5535
         Index           =   0
         Left            =   2760
         TabIndex        =   50
         Top             =   360
         Width           =   15855
         _Version        =   524288
         _ExtentX        =   27966
         _ExtentY        =   9763
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
         SpreadDesigner  =   "Form2.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "工单Header"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   18615
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox TxtTradeType 
         Height          =   285
         Left            =   11040
         TabIndex        =   75
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox TxtDateCode 
         Height          =   285
         Left            =   3960
         TabIndex        =   71
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox TxtWoDept 
         Height          =   285
         Left            =   7080
         TabIndex        =   68
         Top             =   3000
         Width           =   2295
      End
      Begin VB.ComboBox CmbCheckCustomer 
         Height          =   315
         Left            =   1080
         TabIndex        =   55
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox TxtShipSite 
         Height          =   285
         Left            =   16800
         TabIndex        =   53
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox TxtRequestDate 
         Height          =   285
         Left            =   13800
         TabIndex        =   48
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox TxtMpn 
         Height          =   285
         Left            =   7080
         TabIndex        =   46
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox TxtLotStatus 
         Height          =   285
         Left            =   3960
         TabIndex        =   44
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox TxtFilmApld 
         Height          =   285
         Left            =   11040
         TabIndex        =   42
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox TxtPoItem 
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox TxtMMaterial 
         Height          =   285
         Left            =   16800
         TabIndex        =   39
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox TxtCounFab 
         Height          =   285
         Left            =   13800
         TabIndex        =   37
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   11040
         TabIndex        =   35
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox TxtMarkingcode 
         Height          =   285
         Left            =   7680
         TabIndex        =   33
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   3960
         TabIndex        =   31
         Text            =   "Y"
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Txt260 
         Height          =   285
         Left            =   1080
         TabIndex        =   29
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   16800
         TabIndex        =   27
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtDesignId 
         Height          =   285
         Left            =   13800
         TabIndex        =   25
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtCusRev 
         Height          =   285
         Left            =   11040
         TabIndex        =   23
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtFab 
         Height          =   285
         Left            =   7680
         TabIndex        =   21
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtCustomerPT 
         Height          =   285
         Left            =   3960
         TabIndex        =   19
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox TxtPo 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   1440
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   16800
         TabIndex        =   16
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   269025281
         CurrentDate     =   40882
      End
      Begin VB.TextBox TxtDate 
         Height          =   285
         Left            =   13800
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtNum 
         Enabled         =   0   'False
         Height          =   285
         Left            =   11040
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form2.frx":0470
         Left            =   3960
         List            =   "Form2.frx":0472
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "查询OI"
         Height          =   360
         Left            =   6360
         TabIndex        =   4
         Top             =   360
         Width           =   990
      End
      Begin VB.TextBox TxtSourceBatchId 
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo Text3 
         Height          =   315
         Left            =   7680
         TabIndex        =   59
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo CmbCustomer 
         Height          =   315
         Left            =   960
         TabIndex        =   66
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblPRI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "37PRI"
         Height          =   315
         Left            =   8760
         TabIndex        =   77
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "贸易类型"
         Height          =   195
         Left            =   10080
         TabIndex        =   76
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DateCode"
         Height          =   195
         Left            =   3120
         TabIndex        =   72
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PoItem"
         Height          =   195
         Left            =   480
         TabIndex        =   70
         Top             =   2520
         Width           =   510
      End
      Begin VB.Label Label30 
         Caption         =   "加工地："
         Height          =   255
         Left            =   360
         TabIndex        =   69
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label LblWoDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单部门"
         Height          =   195
         Left            =   6120
         TabIndex        =   67
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "接口中的客户"
         Height          =   435
         Left            =   480
         TabIndex        =   56
         Top             =   3000
         Width           =   600
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ShipSite"
         Height          =   195
         Left            =   16200
         TabIndex        =   54
         Top             =   2400
         Width           =   570
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户需求日"
         Height          =   195
         Left            =   12840
         TabIndex        =   49
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mpn(OPN)"
         Height          =   195
         Left            =   6240
         TabIndex        =   47
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LotStatus"
         Height          =   195
         Left            =   3120
         TabIndex        =   45
         Top             =   2520
         Width           =   690
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ProtectiveFilmApld"
         Height          =   195
         Left            =   9600
         TabIndex        =   43
         Top             =   2400
         Width           =   1320
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MicronMaterial"
         Height          =   195
         Left            =   15720
         TabIndex        =   40
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CountryFab"
         Height          =   195
         Left            =   12840
         TabIndex        =   38
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "比率(*)"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   10440
         TabIndex        =   36
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MarkingCode"
         Height          =   195
         Left            =   6720
         TabIndex        =   34
         Top             =   1920
         Width           =   930
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NG标志"
         Height          =   195
         Left            =   3240
         TabIndex        =   32
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level260"
         Height          =   195
         Left            =   360
         TabIndex        =   30
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level235(CS)"
         Height          =   195
         Left            =   15840
         TabIndex        =   28
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DesignId"
         Height          =   195
         Left            =   13080
         TabIndex        =   26
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ImagerCustomerRev"
         Height          =   195
         Left            =   9480
         TabIndex        =   24
         Top             =   1440
         Width           =   1485
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAB设备"
         Height          =   195
         Left            =   6960
         TabIndex        =   22
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户料号"
         Height          =   195
         Left            =   3120
         TabIndex        =   20
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "订单单号"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预计完工日"
         Height          =   195
         Left            =   15840
         TabIndex        =   15
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预计开工日"
         Height          =   195
         Left            =   12840
         TabIndex        =   14
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生产数量"
         Height          =   195
         Left            =   10200
         TabIndex        =   12
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "产品料号"
         Height          =   195
         Left            =   6840
         TabIndex        =   10
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单类型"
         Height          =   195
         Left            =   3120
         TabIndex        =   8
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客户"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source_batch_id"
         Height          =   195
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   1200
      End
   End
End
Attribute VB_Name = "FrmApplyWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail
    E_ID = 1                 'id
    E_WaferID                'Waferid
    E_CompleteFlag           '完成标志
    E_TotalDie               '总数量
    E_GoodDie                'good数量
    E_WaferLot               'wafer
    E_MarkingCode            'markingcode
    E_OK                     '选择
    E_End
    
End Enum

Private Enum E_FPS1          'Bom
    E_ID = 0                 'id
    E_BomID                  '材料规范编号
    E_PT                     '料号
    E_Mt                     '物料编号
    E_Name                   '名称
    e_Qty                    '每只用量
    E_Unit                   '单位
    
    E_Pt2                     '料号2
    E_Mt2                     '物料编号2
    E_Name2                   '名称2
    E_Qty2                    '每只用量2
    E_Unit2                   '单位2
    
    E_End
    
End Enum


Dim oiRS        As New ADODB.Recordset
Dim listRS        As New ADODB.Recordset
Dim bomRS2        As New ADODB.Recordset
Dim bomRSC        As New ADODB.Recordset

Dim bomRS        As New ADODB.Recordset

Dim mainItemRS As New ADODB.Recordset



Private Sub ChkAll_Click()

Dim i As Integer
    If ChkAll.Value = 1 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .Text = 1
            End With
        Next i
        
    ElseIf ChkAll.Value = 0 Then
        For i = 1 To fps(0).MaxRows
            With fps(0)
                .Row = i
                .Col = E_FPS0.E_OK
                .Text = 0
            End With
        Next i
        
    End If


End Sub

Private Sub CmbCustomer_Click(Area As Integer)
If UCase(Trim(CmbCustomer.Text)) = "AA(ON)" Then
TxtMpn.SetFocus
End If
End Sub

Private Sub CmdBom_Click()
'ComSave.Enabled = True
'
'Call addLogTxt(UCase(Trim(Text2.Text)), " 点击Bom料设定 " & "料号：" & Text3.Text)
'
'openWOID = UCase(Trim(Text2.Text))

FrmSemtechWeiPi.Show

End Sub

Private Sub CmdM_Click()
FormM.Show

Unload Me

End Sub

Private Sub CmdNewQuery_Click()
'搜索
Dim lotIDTemp As String
Dim i As Integer
Dim resultTemp As Boolean
resultTemp = False

lotIDTemp = UCase(Trim(TxtQueryLotID.Text))

'查一下，这个lotID在不在List中，如果在，则选中，如不在，则弹出提示


With Lst
        '开始查找赋值
        For i = 0 To .ListCount - 1
        
            If GetLot003(.List(i)) = lotIDTemp Then
                .Selected(i) = True
                
                resultTemp = True
                
                Exit For
            
            
            End If
            
        Next
 End With


If resultTemp = False Then

    MsgBox "请确认LotID是否确认，搜索不到！", vbInformation, "友情提示"
    Exit Sub

End If


TxtQueryLotID.Text = ""
TxtQueryLotID.SetFocus

End Sub

Private Sub Command1_Click()
Dim hyChar As String
'7468859023
'TxtSourceBatchId.Text = "7468859023"

'ComSave.Enabled = False

Dim customerNameTemp As String

If CmbCustomer.Text = "" Or TxtSourceBatchId.Text = "" Then
    MsgBox "请先选择客户，再输入Lot号。请确认!", vbInformation, "友情提示"
    Exit Sub
Else
    customerNameTemp = Trim(CmbCustomer.Text)
    
    If customerNameTemp = "37(ICI)" Then
    customerNameTemp = "37"
    End If

    'Set oiRS = GetOIData(UCase(customerNameTemp), UCase(Trim(TxtSourceBatchId.Text)))
    Set oiRS = GetOIData((customerNameTemp), (Trim(TxtSourceBatchId.Text))) 'GD02客户ID里面有小写字母20161116CCSADD
    If (oiRS.RecordCount > 0) Then

        '2014-05-08 jiayunzhang add
        '查询一下，这个LotId是否存在多个客户机种号
        If JudgeCustomerPTNum(UCase(Trim(TxtSourceBatchId.Text))) Then
        MsgBox "此LotID： " + UCase(Trim(TxtSourceBatchId.Text)) + " 客户WO上有多笔客户机种号，请注意确认目前带出信息是否正确！"

        End If
        
        TxtPo.Text = getStr(oiRS.fields("po_num").Value)
        TxtCustomerPT.Text = getStr(oiRS.fields("mpn_desc").Value)
        
        '2014-07-30 jiayun add 添加HY客户机种
        If UCase(Trim(CmbCustomer.Text)) = "HY" Then
            hyChar = Mid(UCase(Trim(TxtSourceBatchId.Text)), 1, 3)
            
            If hyChar = "SAQ" Then
                TxtCustomerPT.Text = "Hi-257"
                
             ElseIf hyChar = "SAZ" Then
             
               TxtCustomerPT.Text = "Hi-258"
            
            End If
            
            
        
        End If
        
        
        TxtFab.Text = getStr(oiRS.fields("fabrication_facility").Value)
        TxtCusRev.Text = getStr(oiRS.fields("imager_customer_rev").Value)
        TxtDesignId.Text = getStr(oiRS.fields("design_id").Value)
        Txt260.Text = getStr(oiRS.fields("shipping_mst_260").Value)
        Text11.Text = getStr(oiRS.fields("shipping_mst_level").Value)
        TxtMarkingcode.Text = getStr(oiRS.fields("encoded_mark_id").Value)
        TxtCounFab.Text = getStr(oiRS.fields("country_of_fab").Value)
        TxtMMaterial.Text = getStr(oiRS.fields("micron_material").Value)
        TxtPoItem.Text = getStr(oiRS.fields("po_item").Value)
        TxtLotStatus.Text = getStr(oiRS.fields("lot_status").Value)
        TxtMpn.Text = getStr(oiRS.fields("mpn").Value)
        TxtTradeType.Text = getStr(oiRS.fields("PROBE_SHIP_PART_TYPE").Value)
        
        
        If getStr(oiRS.fields("protective_film_apld").Value) = "YES" Then
            TxtFilmApld.Text = "PF"
        Else
            TxtFilmApld.Text = getStr(oiRS.fields("protective_film_apld").Value)
        End If
        
        TxtRequestDate.Text = getStr(oiRS.fields("lot_priority").Value)
        TxtShipSite.Text = getStr(oiRS.fields("ship_site").Value)
        
        If TxtShipSite.Text = "Qtech" And UCase(Trim(CmbCustomer.Text)) = "AA" Then
            CmbCheckCustomer.Text = "WLC"
            
        ElseIf TxtShipSite.Text = "SG" And UCase(Trim(CmbCustomer.Text)) = "AA" Then
            CmbCheckCustomer.Text = "AA"
            
        ElseIf UCase(Trim(CmbCustomer.Text)) = "GC" Then
             CmbCheckCustomer.Text = "GC"
        End If
        
        Call IniProductTwo(UCase(Trim(CmbCustomer.Text)))
        
        '初始化左边的Lot明细表
        
'        Call InitListBox(UCase(Trim(CmbCustomer.Text)))

        
        Call InitListBoxForSoNew(UCase(Trim(CmbCustomer.Text)), TxtPo.Text, TxtCustomerPT.Text)
        
        
        
        
        
        '2012-12-18 jiayun add  根据source_batch_id ，自动带出料号
        '2013-05-13 jiayun add GC
'        Text3.Text = "aa"
         If CmbCustomer.Text = "AA" Then
        
            Call getAutoWo(UCase(Trim(TxtSourceBatchId.Text)))
        
        ElseIf CmbCustomer.Text = "PT" Or CmbCustomer.Text = "SY" Then
        
             Call getOtherCustomerPt(UCase(Trim(TxtSourceBatchId.Text)))
             
        '2016-01-28 jiayun add GC,SX
        ElseIf CmbCustomer.Text = "GC" Or CmbCustomer.Text = "SX" Then
        
             Call getGCSXCustomerPt(UCase(Trim(TxtSourceBatchId.Text)))
        
        Else
          'Call getOtherCustomerPt(UCase(Trim(TxtSourceBatchId.Text)))
        Call getOtherCustomerPt((Trim(TxtSourceBatchId.Text))) 'GD02客户ID里面有小写字母20161116CCSADD
        End If
        
        '2014-04-09 jiayunzhang add DataCode
        
        TxtDateCode.Text = GetWODataCode()
        
        '2015-10-23 jiayun add 最后再校验有没有自动料号
        Dim customerWoPTTemp As String
        Dim npiPTTemp As String
        
        customerWoPTTemp = ""
        npiPTTemp = ""
        
        If Text3.Text = "" Then
        
                '查询客户OI上的机种
                'customerWoPTTemp = GetOICustomerPTNum(UCase(Trim(TxtSourceBatchId.Text)))
                
                 customerWoPTTemp = GetOICustomerPTNum((Trim(TxtSourceBatchId.Text))) 'GD02客户ID里面有小写字母20161116CCSADD
                
                
                If customerWoPTTemp = "" Then
                   MsgBox "客户订单上查不到客户机种，请联系市场部！ ", vbInformation, "友情提示"
                   
                Else
                   MsgBox "客户订单上客户机种为： " & customerWoPTTemp, vbInformation, "友情提示"
                End If
                
                
                '再查NPI 表
                npiPTTemp = GetNpiCustomerPTNum(customerWoPTTemp)
                
                If npiPTTemp = "" Then
                   MsgBox "NPI产品对照表上查不到客户机种" & customerWoPTTemp & " ，请联系NPI！ ", vbInformation, "友情提示"
                   
                Else
                   MsgBox "NPI产品对照表上客户机种为： " & npiPTTemp & " 两者如果一致，但还不能自动带料号，请联系IT！", vbInformation, "友情提示"
                End If
                
        
        
        
        
        End If
        
        
        
        
    
    Else
        MsgBox "查询不到数据，请确认 SourceBatchId ", vbInformation, "友情提示"
        Exit Sub
 
    
    End If
    

    
    
End If
End Sub

Private Sub getAutoWo(lotidTemp2 As String)

Dim lotIDTemp As String
lotIDTemp = lotidTemp2
Dim pfType As String
Dim trayType As String
Dim testno As String

Dim ptFirst As String

pfType = GetString(lotIDTemp)
'LblPF.Caption = pfType

trayType = GetTrayString(lotIDTemp)
'LblTrayType.Caption = trayType

testno = GetTestNoString(lotIDTemp)
'LblTestNo.Caption = testno

'成品料号
'根据OI，查出成品料号的前9位

ptFirst = GetFirstPtString(lotIDTemp)

Dim test1 As String
test1 = GetAllPtString(ptFirst, pfType, trayType, testno)

Text3.Text = GetAllPtString(ptFirst, pfType, trayType, testno)


Dim deptId As String

TxtWoDept.Text = GetWoDept(Text3.Text)

'根据部门查代码

deptId = GetGWoDeptID(TxtWoDept.Text)

TxtWoDept.Text = TxtWoDept.Text & "_" & deptId



End Sub
Private Function getStr(strTemp As Variant)
getStr = Trim("" & strTemp)
End Function

'2013-05-13 jiayun add
Private Sub getOtherCustomerPt(lotidTemp2 As String)
Dim deptId As String


Text3.Text = GetCustomerPtNum(lotidTemp2)


TxtWoDept.Text = GetWoDept(Text3.Text)

'根据部门查代码

deptId = GetGWoDeptID(TxtWoDept.Text)

TxtWoDept.Text = TxtWoDept.Text & "_" & deptId



End Sub

'2016-01-28 jiayun add GC SX 自动带料号
Private Sub getGCSXCustomerPt(lotidTemp2 As String)
Dim deptId As String


Text3.Text = GetCustomerPtNum(lotidTemp2)


TxtWoDept.Text = GetWoDept(Text3.Text)

'根据部门查代码

deptId = GetGWoDeptID(TxtWoDept.Text)

TxtWoDept.Text = TxtWoDept.Text & "_" & deptId



End Sub




Private Sub Command2_Click()
Dim strTmp As String
Dim strTemp As String
Dim lotIDTemp As String
strTemp = ""
'2016-01-14 jiayun Add  工单号有没有输

If Trim(Text2.Text) = "" Then
  MsgBox "请先输入工单号，再选Wafer明细!"
  Exit Sub

End If


With Lst
        '开始查找赋值
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                strTmp = GetLot003(.List(i)) & "','"
                lotIDTemp = GetLot003(.List(i))
                strTemp = strTemp & strTmp

            End If
        Next
 End With
 
 If strTemp = "" Then
 
 MsgBox "请先选择LotId !"
 Exit Sub
 
 Else
 
 strTemp = Mid(strTemp, 1, Len(strTemp) - 3)
 
Call GetFpsData(strTemp, UCase(Trim(CmbCustomer.Text)), UCase(Trim(Text2.Text)), UCase(Trim(TxtCustomerPT.Text)))

'2015-05-20 add  AA New 根据lotid,查询CS
If CmbCustomer.Text = "AA(ON)" Then

  Text11.Text = GetONCS(lotIDTemp)
  
  Txt260.Text = GetONBCPlace(lotIDTemp)
  
  
End If

ChkAll.Value = 1
ChkAll_Click

End If


End Sub
'2013-04-23 jiayun add
Private Function GetLot003(lotIDTemp As String)

GetLot003 = Replace(lotIDTemp, "00A", "003")

End Function

Private Sub GetFpsData(strwhereTemp As String, customerTemp As String, woTemp As String, ptTemp As String)
'明细数据
Dim i As Integer
Dim waferIdTemp As String
Dim woType As String
'ST
woType = Mid(woTemp, 2, 2)


If (customerTemp = "AA" Or customerTemp = "AA(ON)") And (woType = "ST" Or woType = "ET") Then
    
    Set listRS = GetFpsAARTWo(strwhereTemp, customerTemp, woType)

ElseIf customerTemp = "SX" Or customerTemp = "HJ" Then

    Set listRS = GetFpsForSX(strwhereTemp, customerTemp, ptTemp)

Else

    Set listRS = GetFps(strwhereTemp, customerTemp)

End If


If listRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
    
Else

    '2014-11-12 jiayun add

    fps(0).MaxRows = listRS.RecordCount
    For i = 0 To listRS.RecordCount - 1
         
         waferIdTemp = CStr(listRS.fields(1).Value)
         
         ' 如果是HD客户，则查看Sqlserver里，仓库有没有收晶圆
         '2014-11-13 jiayun cancel
'         If customerTemp = "HD" Then
'
'            If JudgeHDWaferStatus(waferidTemp) = False Then
'
'                GoTo NextRecord
'
'            End If
'
'
'         End If
         
         
         
         With fps(0)
                 .Row = i + 1
                 .Col = E_FPS0.E_ID
                 .Text = i + 1
                 
                .Row = i + 1
                 .Col = E_FPS0.E_WaferID
                .Text = CStr(listRS.fields(1).Value)
                
                
                 .Row = i + 1
                 .Col = E_FPS0.E_CompleteFlag
                .Text = ""
                
                  .Row = i + 1
                 .Col = E_FPS0.E_TotalDie
                ' If CmbCustomer.Text = "95" And Len(txt95QTY.Text) > 0 Then
               '  .Text = CStr(txt95QTY.Text)
              '   Else
                 .Text = CStr(listRS.fields(3).Value)
              '   End If
                 
                  .Row = i + 1
                 .Col = E_FPS0.E_GoodDie
                 .Text = CStr(listRS.fields(4).Value)
                 
                 
                  .Row = i + 1
                 .Col = E_FPS0.E_WaferLot
                 .Text = CStr(listRS.fields(5).Value)
                 
                  .Row = i + 1
                 .Col = E_FPS0.E_MarkingCode
                 .Text = "" & listRS.fields(6).Value
                
                
                 .Row = i + 1
                 .Col = E_FPS0.E_OK
                .Text = CStr("1")
                
                   
        
        End With
    
NextRecord:
       
        listRS.MoveNext

    Next


End If

'With fps(0)
'        .MaxRows = 0
'        If listRS.RecordCount > 0 Then
'            Set .DataSource = listRS
'        End If
'End With


End Sub

Private Sub GetBomData(ptTemp As String)
'明细数据

Set bomRS = GetFpsBom(ptTemp)
If bomRS.RecordCount <= 0 Then
    MsgBox "明细表中没有相关数据，请确认"
    Exit Sub
End If

With fps(1)
        .MaxRows = 0
        If bomRS.RecordCount > 0 Then
            Set .DataSource = bomRS
        End If
End With

End Sub



Private Sub InitListBox(customerTemp As String)
Dim i As Integer
      Set listRS = GetLotDetailData(customerTemp)
       With Lst
            .Clear
            listRS.MoveFirst
            
            For i = 0 To listRS.RecordCount - 1
            
         
                .AddItem "" & listRS!source_batch_id
                
                If "" & listRS!source_batch_id = TxtSourceBatchId.Text Then
                    Lst.Selected(i) = True
                End If
                
                listRS.MoveNext
         
            
            Next
        End With
        
      
        

listRS.Close
Set listRS = Nothing

End Sub

Private Sub InitListBoxForSo(customerTemp As String, soTemp As String)
Dim i As Integer
      Set listRS = GetLotDetailDataForSo(customerTemp, soTemp)
      If listRS.RecordCount > 0 Then
      
       With Lst
            .Clear
            listRS.MoveFirst
            
            For i = 0 To listRS.RecordCount - 1
            
         
                .AddItem "" & listRS!source_batch_id
                
                If "" & listRS!source_batch_id = TxtSourceBatchId.Text Then
                    Lst.Selected(i) = True
                End If
                
                listRS.MoveNext
         
            
            Next
        End With
        
        End If
        
      
        

listRS.Close
Set listRS = Nothing

End Sub

Private Sub InitListBoxForSoNew(customerTemp As String, soTemp As String, customerPTTemp As String)
Dim i As Integer
'    If customerTemp = "37(ICI)" Then
'    customerTemp = "37"
'
'    End If

      Set listRS = GetLotDetailDataForSoNew(customerTemp, soTemp, customerPTTemp)
      If listRS.RecordCount > 0 Then
      
       With Lst
            .Clear
            listRS.MoveFirst
            
            For i = 0 To listRS.RecordCount - 1
            
         
                .AddItem "" & listRS!source_batch_id
                
                If "" & listRS!source_batch_id = TxtSourceBatchId.Text Then
                    Lst.Selected(i) = True
                End If
                
                listRS.MoveNext
         
            
            Next
        End With
        
        End If
        
      
        

listRS.Close
Set listRS = Nothing

End Sub

Private Sub InitListBoxForSoNewForOn(customerPTTemp As String, opnTemp As String)
Dim i As Integer
      Set listRS = GetLotDetailDataForSoNewOn(customerPTTemp, opnTemp)
      If listRS.RecordCount > 0 Then
      
       With Lst
            .Clear
            listRS.MoveFirst
            
            For i = 0 To listRS.RecordCount - 1
            
         
                .AddItem "" & listRS!source_batch_id
                
'                If "" & listRS!source_batch_id = TxtSourceBatchId.Text Then
'                    Lst.Selected(i) = True
'                End If
                
                listRS.MoveNext
         
            
            Next
        End With
        
        End If
        
      
        

listRS.Close
Set listRS = Nothing

End Sub





Private Sub Command3_Click()

  Dim woTemp As String
  woTemp = UCase(Trim(Text2.Text))
  
 Dim sqlTemp As String
 sqlTemp = "select SEQ_IBWO,ORDERNAME,ORDERTYPE,DESCRIPTION,EVENTTYPE,ERPUSER,PRODUCT,PRODUCTREVISION,QTY,PRODUCTBOM,ERPCREATEDATE,PLANSTARTDATE,PLANENDDATE," & _
         " Customer , SalesOrder, PRODUCTFAMILY, ModifyFlag, CUSTOMERPN, FabFacility, ImagerRev, Designid, MLevel235, Mlevel260, NGFlag, Para1, Para2, Para3, Para4, Para5, Para6, PARA7, PARA8, PARA9, PARA10, Protective_Film_Apld, LOT_STATUS, MPN " & _
         " From IB_WOHISTORY where ORDERNAME='" + woTemp + "'order by SEQ_IBWO desc "
  ExporToExcel (sqlTemp)
End Sub

Private Sub Command4_Click()

  Dim woTemp As String
  woTemp = UCase(Trim(Text2.Text))
  
 Dim sqlTemp As String
 sqlTemp = "select ORDERNAME,WAFERID,COMPLETEFLAG,DIEQTY,FGDIEQTY,WAFERLOT,WAFERSEQUENCE,MARKINGCODE from IB_WAFERLIST where ordername ='" + woTemp + "' order by ORDERNAME, WAFERID"
  ExporToExcel (sqlTemp)

End Sub

Private Sub Command5_Click()

ClearData

End Sub

Private Sub ClearData()
'清空上一笔的数据
TxtSourceBatchId.Text = ""
Text2.Text = ""
Text3.Text = ""
TxtNum.Text = ""
TxtPo.Text = ""
TxtCustomerPT.Text = ""
TxtFab.Text = ""
TxtCusRev.Text = ""
TxtDesignId.Text = ""
Text11.Text = ""
Txt260.Text = ""
Text13.Text = ""
TxtMarkingcode.Text = ""
Text15.Text = ""
TxtCounFab.Text = ""
TxtMMaterial.Text = ""
TxtPoItem.Text = ""
TxtLotStatus.Text = ""
TxtMpn.Text = ""
TxtFilmApld.Text = ""
TxtRequestDate.Text = ""
TxtShipSite.Text = ""
CmbCheckCustomer.Text = ""
Lst.Clear

fps(0).MaxRows = 0


End Sub

Private Sub Command6_Click()

On Error Resume Next

Dim lotIDTemp As String
Dim qtyTemp As Long
Dim productTemp As String
Dim erpdate As String
Dim woDeptIDTemp As String



Set billLotTemp = GetBillLot2()
If (billLotTemp.RecordCount > 0) Then

    Do While Not billLotTemp.EOF
        lotIDTemp = billLotTemp.fields("waferlot").Value
        productTemp = billLotTemp.fields("productname").Value
        qtyTemp = CLng(billLotTemp.fields("qty").Value)
        
        erpdate = Format(CDate(billLotTemp.fields("erpcreationdate").Value), "YYYY-MM-DD")
        
        woDeptIDTemp = billLotTemp.fields("deptid").Value
        
   
                
          '-----begin------
        
         Set adoCmd = New ADODB.Command
         Set adoCmd.ActiveConnection = INIadoCon2
             adoCmd.CommandText = "uspPMC_XDInterface"
             adoCmd.Parameters.Refresh
             adoCmd.CommandType = adCmdStoredProc
        
          Set adoprm1 = New ADODB.Parameter   '工单号
          adoprm1.Type = adChar
          adoprm1.Size = 20
          adoprm1.Direction = adParamInput
          adoprm1.Value = lotIDTemp
          adoCmd.Parameters.Append adoprm1
        
          Set adoprm2 = New ADODB.Parameter   '料号
          adoprm2.Type = adChar
          adoprm2.Size = 20
          adoprm2.Direction = adParamInput
          adoprm2.Value = productTemp
          adoCmd.Parameters.Append adoprm2
        
          Set adoprm3 = New ADODB.Parameter   '数量
          adoprm3.Type = adInteger
          adoprm3.Direction = adParamInput
          adoprm3.Value = qtyTemp
          adoCmd.Parameters.Append adoprm3
          
            Set adoprm4 = New ADODB.Parameter   '时间
             
          adoprm4.Type = adChar
         adoprm4.Size = 20
          adoprm4.Direction = adParamInput
          adoprm4.Value = erpdate
          adoCmd.Parameters.Append adoprm4
          
          
            Set adoprm5 = New ADODB.Parameter   '线别
          adoprm5.Type = adInteger
          adoprm5.Direction = adParamInput
          adoprm5.Value = 1
          adoCmd.Parameters.Append adoprm5
          
          Set adoprm6 = New ADODB.Parameter   '部门id
          adoprm6.Type = adChar
          adoprm6.Size = 30
          adoprm6.Direction = adParamInput
          adoprm6.Value = woDeptIDTemp
          adoCmd.Parameters.Append adoprm6
          
          
          
        
          adoCmd.Execute

        
        billLotTemp.MoveNext
   
    Loop
    
End If

MsgBox ("OK")





End Sub

Private Sub ComSave_Click()
'保存工单

Dim headerTemp As BillHeader
Dim detailTemp As BillDetail
Dim typeId As Integer
Dim SumQty As Long
Dim i As Integer
Dim yTemp As String
Dim yTemp1 As String
Dim date0 As Date
Dim date1 As Date
Dim ww As String
Dim mCodetemp As String




ComSave.Enabled = False

SumQty = 0


 
' stRfpc95qtytxt = Trim(txt95QTY.Text)
 
' If CmbCustomer.Text = "95" And TxtSourceBatchId = "95FPC" Then
 '  stRfpc95qtytxt = Trim(txt95QTY.Text)
 '  If stRfpc95qtytxt = "" Then
 '  MsgBox ("95FPC数量必须要录入!")
 '  Exit Sub
  ' End If
' End If

Call addLogTxt(UCase(Trim(Text2.Text)), " 点击保存按钮 ")




'Check介面数据是否填写
If Trim(Text15.Text) = "" Then
     MsgBox "比率不可以为空！"
     ComSave.Enabled = True
     Exit Sub
End If

'附值
 headerTemp.id = GetSeqID()
 headerTemp.OrderName = Replace(UCase(Trim(Text2.Text)), Chr(13) + Chr(10), "")
 
 If UCase(Trim(Text2.Text)) = "" Then
      MsgBox "工单号不可以为空！"
      ComSave.Enabled = True
     Exit Sub
 
 End If
 
  If Len(UCase(Trim(Text2.Text))) <> 12 Then
      MsgBox "工单号长度不对！"
      ComSave.Enabled = True
     Exit Sub
 
 End If
 
 
 
 
  If UCase(Trim(TxtWoDept.Text)) = "" Then
      MsgBox "工单部门不可以为空！"
      ComSave.Enabled = True
     Exit Sub
 
 End If
 
 
 
 '2013-08-30 jiayun add 校验工单号是否已存在
 Set bomRS2 = GetWoData(UCase(Trim(Text2.Text)))
If bomRS2.RecordCount > 0 Then
    MsgBox "Mes系统中已存在此工单号，请确认工单号 ！"
    ComSave.Enabled = True
    Exit Sub
End If

 '2013-08-30 jiayun add 校验工单号是否已存在
 Set bomRS2 = GetWoData2(UCase(Trim(Text2.Text)))
If bomRS2.RecordCount > 0 Then
    MsgBox "中间表中已存在此工单号，请不要重复提交 ！"
    ComSave.Enabled = True
    Exit Sub
End If

 
 Call addLogTxt(UCase(Trim(Text2.Text)), " Insert DB 前，数据校验 ")
 
 '2012-11-30 jiayun add 判断料号的bom是否存在
Set bomRS2 = GetProductBom(Text3.Text)
If bomRS2.RecordCount <= 0 Then
    MsgBox "新系统中这料号的Bom不存在！请联系相关的人，先维护Bom ！"
    ComSave.Enabled = True
    Exit Sub
End If


'ccs add 校验产品对照表是否建立
'Set bomRS = GetProductCheckTT(Text3.Text)
'If bomRS.RecordCount <= 0 Then
 '   MsgBox "产品对照表没有建立 ！"
  '  ComSave.Enabled = True
  '  Exit Sub
'End If
'2014-01-13 jiayun add判断料号金碟是否有成本对象

Set bomRS2 = GetProductJDObject(Text3.Text)
If bomRS2.RecordCount <= 0 Then
    MsgBox "此料号在金碟系统中无成本对象，请找相关人员确认 ！"
    ComSave.Enabled = True
    
    Exit Sub
End If


'GetWLOWo


 '2012-12-19 jiayun add 校验料号是否存在
Set bomRS2 = GetProduct_Check(Text3.Text)
If bomRS2.RecordCount <= 0 Then
    MsgBox "料号不存在！请联系相关的人，先维护料号 ！"
    ComSave.Enabled = True
    Exit Sub
End If



 '2014-01-14 jiayun add 判断新ERP bom 有没有签核过
 
Set bomRS2 = GetProductBomERpSign(Text3.Text)
If bomRS2.RecordCount <= 0 Then
    MsgBox "新系统中这料号的Bom没有被审核通过，请联系工程部！"
    ComSave.Enabled = True
    Exit Sub
End If


'2015-05-21 jiayun add AA 新系统加校验

If CmbCustomer.Text = "AA(ON)" Then
    If TxtCustomerPT.Text = "" Then
    
    End If
    
     If UCase(Trim(TxtCustomerPT.Text)) = "" Then
      MsgBox "客户料号不可以为空！"
      Exit Sub
 
      End If
      
      If UCase(Trim(Text11.Text)) = "" Then
      MsgBox "CS字段不可以为空！"
      Exit Sub
 
      End If
      
      If UCase(Trim(Txt260.Text)) = "" Then
      MsgBox "加工地字段不可以为空！"
      Exit Sub
      End If
 
End If




'2014-01-02  jiayun add Check AA MarkingCode
'-----begin

With Lst
        '开始查找赋值
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                strTmp = .List(i)
                
                
              

            End If
        Next
 End With



'-----end

Select Case Combo2.Text
Case "一般工单"
    typeId = 1
Case "再加工工单"
    typeId = 5
Case "委外工单"
    typeId = 7
    
Case "重工委外工单"
    typeId = 8
    
Case "拆件式工单"
    typeId = 11
    
Case "预测工单"
    typeId = 13
Case "试产工单"
    typeId = 15
    
Case Else
   typeId = 0
End Select

 headerTemp.OrderType = CStr(typeId)
 headerTemp.EventType = "CREATED"
 headerTemp.ERPUser = "Auto"
 headerTemp.product = Text3.Text
                            
 headerTemp.RequestDate = Now
 headerTemp.ERPCreateDate = DateTime.Date
 headerTemp.PlanStartDate = CDate(TxtDate.Text)
 headerTemp.PlanEndDate = DTPicker1.Value
 headerTemp.CUSTOMER = Replace(CmbCustomer.Text, "(ICI)", "")
 headerTemp.SalesOrder = TxtPo.Text
 headerTemp.ModifyFlag = 0
 headerTemp.CustomerERPN = TxtCustomerPT.Text
 headerTemp.FabFacility = TxtFab.Text
headerTemp.ImagerRev = TxtCusRev.Text
headerTemp.DesignId = TxtDesignId.Text
headerTemp.MLevel235 = Text11.Text
headerTemp.Mlevel260 = Txt260.Text

headerTemp.NGFlag = Val(Text13.Text)

'If CmbCustomer.Text = "81" Then
'yTemp = Mid(Year(Date), 3, 1)
'yTemp1 = Mid(Year(Date), 4, 1)
'date1 = Date
'date0 = Year(date1) & "/01/01"
'ww = DateDiff("ww", date0, date1) + 1
'If Len(ww) < 2 Then
'ww = "0" & ww
'End If
'mCodetemp = "HS" & yTemp & "A" & yTemp1 & "S" & ww
'headerTemp.Para1 = mCodetemp
'Else
headerTemp.Para1 = TxtMarkingcode.Text

'End If


headerTemp.Para2 = Text15.Text
headerTemp.Para3 = TxtCounFab.Text

'2016-03-16 jiayun add 贸易类型
'headerTemp.Para4 = TxtMMaterial.Text

headerTemp.Para4 = Trim(TxtTradeType.Text)

headerTemp.Para5 = TxtPoItem.Text
headerTemp.Para6 = TxtShipSite.Text
headerTemp.Para8 = TxtWoDept.Text
headerTemp.Para10 = TxtDateCode.Text

headerTemp.Protective_Film_Apld = TxtFilmApld.Text
headerTemp.Lot_Stauts = TxtLotStatus.Text
headerTemp.MPN = TxtMpn.Text
 
With fps(0)

For i = 1 To .MaxRows
    .Row = i
    .Col = 8
    If .Text = 1 Then
        .Row = i
        .Col = 4
        SumQty = SumQty + CLng(.Text)
    End If

Next i

End With

'If headerTemp.CUSTOMER = "95" And Len(txt95QTY.Text) > 0 And TxtSourceBatchId = "95FPC" Then
'headerTemp.qty = txt95QTY.Text
'Else
headerTemp.qty = SumQty
'End If


Call addLogTxt(UCase(Trim(Text2.Text)), " 工单类 各字段付值成功 ")
 '--附值End
Call AddBillHeaderWo(headerTemp)
  
Call AddWOPRI(headerTemp.OrderName, Trim(Combo1.Text))
  
'--保存Heand End

'--- Begin Detail

'判断这笔工单，对应客户的OI,是否已用完


'MsgBox "工单：" & Text2.Text & "建立成功 !"

ComSave.Enabled = True

End Sub

Private Sub InitCtrl1()
Dim i                   As Integer
Dim strSql              As String
Dim Rs                  As New ADODB.Recordset
    
    '加载单据类型
    strSql = "  select distinct a.pri as PRI from PJ_WO_PRI a "
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    Combo1.Clear
    If Not Rs.EOF Then
        Do While Not Rs.EOF
            Combo1.AddItem Trim$("" & Rs!PRI)
            Rs.MoveNext
        Loop
        Combo1.ListIndex = 0
    End If
    Rs.Close
   
End Sub

Private Sub Form_Activate()
Text15.Text = "25"
End Sub

Private Sub Form_Load()
'CmbCustomer.AddItem ("AA")
'CmbCustomer.AddItem ("BD")
'CmbCustomer.AddItem ("DN")
'CmbCustomer.AddItem ("GC")
'CmbCustomer.AddItem ("HY")
'CmbCustomer.AddItem ("MC")
'CmbCustomer.AddItem ("OT")
'CmbCustomer.AddItem ("PT")
'CmbCustomer.AddItem ("QT")
'CmbCustomer.AddItem ("RD")
'CmbCustomer.AddItem ("SX")
'CmbCustomer.AddItem ("SY")
'CmbCustomer.AddItem ("ZX")
'CmbCustomer.AddItem ("ENG")
'根据金碟系统中的客户代码　来初始化

'ComSave.Enabled = False


IniCustomerName

CmbCheckCustomer.AddItem ("AA")
CmbCheckCustomer.AddItem ("WLC")
CmbCheckCustomer.AddItem ("GC")
CmbCheckCustomer.AddItem ("SX")
CmbCheckCustomer.AddItem ("SY")
CmbCheckCustomer.AddItem ("BD")
CmbCheckCustomer.AddItem ("HY")
CmbCheckCustomer.AddItem ("MG")


IniProduct

TxtDate.Text = Format(Now, "yyyy-mm-dd")
DTPicker1.Value = TxtDate.Text

Combo2.AddItem ("一般工单")
Combo2.AddItem ("再加工工单")
Combo2.AddItem ("委外工单")
Combo2.AddItem ("重工委外工单")
Combo2.AddItem ("拆件式工单")
Combo2.AddItem ("预测工单")
Combo2.AddItem ("试产工单")
Combo2.AddItem ("小批量试产工单")

IniFpsHeader
'IniFpsBom

Call InitCtrl1

End Sub

Private Sub IniProduct()
Set mainItemRS = GetProduct()
Set Text3.RowSource = mainItemRS
Text3.ListField = mainItemRS("productname").Name
Text3.BoundColumn = mainItemRS("PID").Name

End Sub

Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CmbCustomer.RowSource = mainItemRS
CmbCustomer.ListField = mainItemRS("productname").Name
CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub



Private Sub IniProductTwo(customerTemp As String)
If customerTemp = "AA" Then
    Set Text3.RowSource = Nothing
    Set mainItemRS = GetProductAA()
    Set Text3.RowSource = mainItemRS
    Text3.ListField = mainItemRS("productname").Name
    Text3.BoundColumn = mainItemRS("PID").Name
    
 ElseIf customerTemp = "GC" Then
    
    Set Text3.RowSource = Nothing
    Set mainItemRS = GetProductBB()
    Set Text3.RowSource = mainItemRS
    Text3.ListField = mainItemRS("productname").Name
    Text3.BoundColumn = mainItemRS("PID").Name
    
ElseIf customerTemp = "BD" Then
    
    Set Text3.RowSource = Nothing
    Set mainItemRS = GetProductBD()
    Set Text3.RowSource = mainItemRS
    Text3.ListField = mainItemRS("productname").Name
    Text3.BoundColumn = mainItemRS("PID").Name
    
End If

'Set mainItemRS = GetProduct()
'Set Text3.RowSource = mainItemRS
'Text3.ListField = mainItemRS("productname").Name
'Text3.BoundColumn = mainItemRS("PID").Name

End Sub


Private Sub IniFpsHeader()
    With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .Col = E_FPS0.E_OK
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
          
        .SetText E_FPS0.E_ID, 0, "序号"
        .SetText E_FPS0.E_WaferID, 0, "WaferId"
        .SetText E_FPS0.E_CompleteFlag, 0, "完成标志"
        .SetText E_FPS0.E_TotalDie, 0, "TotalDie数量"
        .SetText E_FPS0.E_GoodDie, 0, "GoodDie数量"
        .SetText E_FPS0.E_WaferLot, 0, "WaferLot"
        .SetText E_FPS0.E_MarkingCode, 0, "MarkingCode"
        .SetText E_FPS0.E_OK, 0, "选择"
        
        
        .ColWidth(E_FPS0.E_ID) = 10
        .ColWidth(E_FPS0.E_WaferID) = 15
        .ColWidth(E_FPS0.E_CompleteFlag) = 10
        .ColWidth(E_FPS0.E_TotalDie) = 12
        .ColWidth(E_FPS0.E_GoodDie) = 12
        .ColWidth(E_FPS0.E_WaferLot) = 15
        .ColWidth(E_FPS0.E_MarkingCode) = 10
        .ColWidth(E_FPS0.E_OK) = 10

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .Col = E_FPS0.E_OK
        .Lock = False
        
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub IniFpsBom()
    With fps(1)
        .ReDraw = False
        .MaxCols = E_FPS1.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
      
        
        .SetText E_FPS1.E_ID, 0, "序号"
        .SetText E_FPS1.E_BomID, 0, "材料规范编号"
        .SetText E_FPS1.E_PT, 0, "料号"
        .SetText E_FPS1.E_Mt, 0, "物料编号"
        .SetText E_FPS1.E_Name, 0, "名称"
        .SetText E_FPS1.e_Qty, 0, "每只用量"
        .SetText E_FPS1.E_Unit, 0, "单位"
        
        .SetText E_FPS1.E_Pt2, 0, "备料料号"
        .SetText E_FPS1.E_Mt2, 0, "备料物料编号"
        .SetText E_FPS1.E_Name2, 0, "备料名称"
        .SetText E_FPS1.E_Qty2, 0, "备料每只用量"
        .SetText E_FPS1.E_Unit2, 0, "备料单位"
    
        
        
        .ColWidth(E_FPS1.E_ID) = 6
        .ColWidth(E_FPS1.E_BomID) = 12
        .ColWidth(E_FPS1.E_PT) = 14
        .ColWidth(E_FPS1.E_Mt) = 14
        .ColWidth(E_FPS1.E_Name) = 14
        .ColWidth(E_FPS1.e_Qty) = 10
        .ColWidth(E_FPS1.E_Unit) = 8
        
        .ColWidth(E_FPS1.E_Pt2) = 14
        .ColWidth(E_FPS1.E_Mt2) = 14
        .ColWidth(E_FPS1.E_Name2) = 14
        .ColWidth(E_FPS1.E_Qty2) = 10
        .ColWidth(E_FPS1.E_Unit2) = 8
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
'生成工单号
'年年月月+四位编码
Dim FirstChar As String
Dim SeqChar As String
Dim typenameTemp As String
Dim yMonthTemp As String
Dim seqTemp As Integer
Dim headChar As String
Dim mdChar As String
Dim id As Long





'2012-11-06 因新旧系统　临时取消自动生成

FirstChar = UCase(Trim(Text2.Text))
 If KeyAscii = 13 Then
    If FirstChar = "" Then
        MsgBox "请输入工单前三位!"
        Exit Sub
    End If
    
     If Len(FirstChar) <> 3 Then
        MsgBox "请输入工单前三位!"
        Exit Sub
    End If

headChar = FirstChar

    SeqChar = GetWoIDTemp(FirstChar)
    mdChar = Right(Year(DateTime.Date), 2) & Right("0" & Month(DateTime.Date), 2)
    FirstChar = FirstChar & "-" & mdChar
    
    SeqChar = Right("000" & CStr(CInt(SeqChar)), 4)
    
    id = CLng(SeqChar)
    
    Text2.Text = FirstChar & SeqChar

    If Mid$(Trim(Text2.Text), 2, 1) = "P" Then
        Combo2.Text = "一般工单"
    End If

    If Mid$(Trim(Text2.Text), 2, 1) = "T" Then
        Combo2.Text = "小批量试产工单"
    End If
    
    '把序列号写到表中
    
  cmdStr = "insert into TSV_WO_SEQ_TAB(wotype,ymonth,sequenceid,flag) values ( '" & headChar & "','" & mdChar & "'," & id & ", 'Y' ) "
  AddSql (cmdStr)
    
 End If


End Sub

Private Sub Text3_Change()
'选择产品料号，来显示Bom
'Dim ptTemp As String
''ptTemp = Text3.Text
'
'ptTemp = "18V117FD00CF"
' Call GetBomData(ptTemp)


Dim deptId As String


TxtWoDept.Text = GetWoDept(Text3.Text)

'根据部门查代码

deptId = GetGWoDeptID(TxtWoDept.Text)

TxtWoDept.Text = TxtWoDept.Text & "_" & deptId




End Sub

Private Sub TxtCustomerPT_KeyPress(KeyAscii As Integer)
'Dim ptTemp As String
'
' If KeyAscii = 13 And UCase(Trim(CmbCustomer.Text)) = "AA(ON)" Then
'
' ptTemp = UCase(Trim(TxtCustomerPT))
'
' '串厂内料号
' Text3.Text = GetON_HTKJPT(ptTemp)
'
' If Text3.Text = "" Then
'    MsgBox "厂内料号查询不到，请确认厂内产品料号对照表！", vbInformation, "友情提示"
'    Exit Sub
'
' Else
' '查询lotid
'  Call InitListBoxForSoNewForOn(ptTemp)
'
'  '查询dateCode
'  TxtDateCode.Text = GetONDateCode()
'
'  TxtMpn.Text = GetONOPN(TxtCustomerPT.Text)
'
'  TxtMarkingcode.Text = GetONWoMarkingCode(TxtMpn.Text)
'
'
'
'
' End If
'
'
'
' End If
 
End Sub

Private Sub TxtMpn_KeyPress(KeyAscii As Integer)
Dim ptTemp As String

 If KeyAscii = 13 And UCase(Trim(CmbCustomer.Text)) = "AA(ON)" Then
 
 ptTemp = UCase(Trim(TxtMpn))
 
 '串厂内料号
 Text3.Text = GetON_HTKJPT(ptTemp)
 
 If Text3.Text = "" Then
    MsgBox "厂内料号查询不到，请确认厂内产品料号对照表！", vbInformation, "友情提示"
    Exit Sub

 Else
 '查询lotid
 
 
 TxtCustomerPT.Text = GetONOPN_WSG(ptTemp)
 
  Call InitListBoxForSoNewForOn(TxtCustomerPT.Text, ptTemp)

  '查询dateCode
  TxtDateCode.Text = GetONDateCode()

  'TxtMpn.Text = GetONOPN(TxtCustomerPT.Text)

  TxtMarkingcode.Text = GetONWoMarkingCode(TxtMpn.Text)




 End If
 
 

 End If
 


End Sub

Private Sub TxtQueryLotID_KeyPress(KeyAscii As Integer)
'输入LotID，回车
    If KeyAscii = 13 Then
    
    CmdNewQuery_Click
    
    
    
    
    End If
End Sub
