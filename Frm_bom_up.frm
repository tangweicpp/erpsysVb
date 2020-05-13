VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_bom_up 
   Caption         =   "Form_BOM"
   ClientHeight    =   13350
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   21855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form_BOM"
   MDIChild        =   -1  'True
   ScaleHeight     =   13350
   ScaleWidth      =   21855
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox CoPZD 
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   1920
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTBOM_UP 
      Height          =   13095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   21615
      _ExtentX        =   38126
      _ExtentY        =   23098
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "BOM_UP"
      TabPicture(0)   =   "Frm_bom_up.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtPath"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabel1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabel4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabel5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabel6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CobPN"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblBOM"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fpS(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CommonDialog1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdUP"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdQu"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdExp"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtText4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtText5"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "ChkAll"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtMPN"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TextWLBH"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Frm_bom_up.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox TextWLBH 
         Height          =   375
         Left            =   1560
         TabIndex        =   23
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtMPN 
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox ChkAll 
         Caption         =   "全选/全不选"
         Height          =   495
         Left            =   18120
         TabIndex        =   17
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtText5 
         Height          =   375
         Left            =   5040
         TabIndex        =   16
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtText4 
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Caption         =   "修改"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   14295
         Begin VB.CommandButton CmdBomDel 
            BackColor       =   &H000000FF&
            Caption         =   "删除"
            Height          =   360
            Left            =   13200
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   480
            Width           =   990
         End
         Begin VB.CommandButton CmdBomAddSave 
            BackColor       =   &H000080FF&
            Caption         =   "添加后提交"
            Height          =   360
            Left            =   7440
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton CmdBomAdd 
            Caption         =   "添加一行"
            Height          =   360
            Left            =   6120
            TabIndex        =   10
            Top             =   360
            Width           =   990
         End
         Begin VB.CommandButton CmdBomModify 
            BackColor       =   &H00C0C000&
            Caption         =   "修改用量提交"
            Height          =   360
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblLabel3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "修改用量"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdExp 
         Caption         =   "导出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8880
         TabIndex        =   4
         Top             =   840
         Width           =   990
      End
      Begin VB.CommandButton cmdQu 
         Caption         =   "查询"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7560
         TabIndex        =   3
         Top             =   840
         Width           =   990
      End
      Begin VB.CommandButton cmdUP 
         Caption         =   "整体上传"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10320
         TabIndex        =   2
         Top             =   840
         Width           =   1350
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   12480
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin FPSpreadADO.fpSpread fpS 
         Height          =   9615
         Index           =   0
         Left            =   105
         TabIndex        =   6
         Top             =   3360
         Width           =   21375
         _Version        =   524288
         _ExtentX        =   37703
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
         MaxCols         =   3
         MaxRows         =   0
         SpreadDesigner  =   "Frm_bom_up.frx":0038
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "料号："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   240
         TabIndex        =   24
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label lblBOM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BOM站点："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3840
         TabIndex        =   21
         Top             =   1920
         Width           =   1140
      End
      Begin MSForms.ComboBox CobPN 
         Height          =   375
         Left            =   5040
         TabIndex        =   20
         Top             =   840
         Width           =   2295
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4048;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblLabel6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "机种名："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   920
         Width           =   960
      End
      Begin VB.Label lblLabel5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上传时间："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         TabIndex        =   15
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblLabel4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人员："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblLabel1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "成品料号:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3840
         TabIndex        =   5
         Top             =   920
         Width           =   1035
      End
      Begin MSForms.TextBox txtPath 
         Height          =   315
         Left            =   12000
         TabIndex        =   1
         Top             =   840
         Width           =   5655
         VariousPropertyBits=   746604563
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "9975;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
   End
End
Attribute VB_Name = "Frm_bom_up"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FPSMaxRowBeforeAdd     As Integer

Public bomId         As String '材料规范编号
Public url           As String

Private Enum E_FPS1          'Bom

    E_PRODUCTID = 1              '工序号
    E_PT                     '料号
    E_MATNUM                '物料编号
    E_name                   '名称
    E_GG                     '规格
    E_XH                     '型号
    
    E_QTY                    '每只用量
    E_Rate                   '损耗量
    E_UNIT                   '单位
    
    E_Typeid                 '序号1
    E_TypePT                 '材料类型
    E_TypePT1                '材料类型1
    E_SEL                    '勾选
    E_END
    
End Enum

Private Sub ChkAll_Click()

    Dim i As Integer
    
    If chkall.Value = 1 Then

        For i = 1 To Fps(0).MaxRows

            With Fps(0)
                .Row = i
                .Col = E_FPS1.E_SEL
                .Text = 1

            End With

        Next i
        
    ElseIf chkall.Value = 0 Then

        For i = 1 To Fps(0).MaxRows

            With Fps(0)
                .Row = i
                .Col = E_FPS1.E_SEL
                .Text = 0

            End With

        Next i
        
    End If
End Sub

Private Sub CmdBomAdd_Click() '新增一笔
    Dim i              As Integer

    Dim strproduct     As String

    With Fps(0)
        
        .MaxRows = .MaxRows + 1
        i = .MaxRows
        
        .Row = i - 1
        .Col = 1
        strproduct = .Text

        .Row = i
        .Col = 1
        .Text = strproduct '工序号沿用上一行
        
        .Row = i
        .Col = 2
        .Lock = False    '料号栏可编辑

        
        .Row = i
        .Col = 7
        .Lock = False    '每只用量可编辑
        .Text = "0.0000"
        
        .Row = i
        .Col = 8
        .Lock = False    '损耗量可编辑
        .Text = "0.0000"
        
        .Row = i
        .Col = 10
        .Lock = False     '序号1可编辑
        
        .Row = i
        .Col = 10
        .Lock = False     '材料类型
        
        .Row = i
        .Col = 11
        .Lock = False    '材料类型1

    End With
End Sub

Private Sub CmdBomAddSave_Click() '添加后提交
    Dim i              As Integer

    Dim strproduct     As String

    Dim strmateriel    As String
    
    Dim materiel_num   As String

    Dim strname        As String

    Dim strspec        As String

    Dim strmodel       As String

    Dim usage          As Double

    Dim strloss        As Double

    Dim unit           As String

    Dim strsite        As String

    Dim strtype        As String
    
    Dim bom_group      As String
    
    Dim strSql         As String
    
    Dim User As String
    
    User = gUserName
    
    If Fps(0).MaxRows = FPSMaxRowBeforeAdd Then
        MsgBox "请新增行之后再点提交!", vbInformation, "提示"
        Exit Sub
    End If

    With Fps(0)
    '先检查数据
        For i = FPSMaxRowBeforeAdd + 1 To .MaxRows
   
            .Row = i
            .Col = 2
            strmateriel = Trim(.Text) '料号

            .Row = i
            .Col = 3
            materiel_num = Trim(.Text) '物料编号

            .Row = i
            .Col = 11
            strtype = Trim(.Text) '材料类型
            
            .Row = i
            .Col = 7
            
            If IsNumeric(Trim(.Text)) = False Then
            
                MsgBox "料号" & strmateriel & "的用量填写有误，请修正后再提交", vbInformation, "提示"
                Exit Sub
                
            End If
            If Trim(.Text) <= 0 Then
            
                MsgBox "料号" & strmateriel & "的用量填写有误，请修正后再提交", vbInformation, "提示"
                Exit Sub
            
            End If

   
            .Row = i
            .Col = 8

            If IsNumeric(Trim(.Text)) = False Then
            
                MsgBox "料号" & strmateriel & "的损耗填写有误，请修正后再提交", vbInformation, "提示"
                Exit Sub
                
            End If
            
            .Row = i
            .Col = 10
            strsite = Trim(.Text) '序号1
            
            If strmateriel = "" Or materiel_num = "" Or strtype = "" Or strsite = "" Then
                MsgBox "请将数据补充完整再提交!", vbInformation, "友情提示"
                Exit Sub
            
            End If
            'merry 20191104判断工序在MES中是否存在
            If CheckProc(url & strsite) = "NG" Then
                MsgBox strsite & "站别在MES中不存在，请修改后再上传！", vbInformation, "提示"
                Exit Sub
            End If
            
        
        Next
        For i = FPSMaxRowBeforeAdd + 1 To .MaxRows
   
            .Row = i
            .Col = 1
            strproduct = Trim(.Text) '工序号
   
            .Row = i
            .Col = 2
            strmateriel = Trim(.Text) '料号
            
            
            .Row = i
            .Col = 3
            materiel_num = Trim(.Text) '物料编号
          
            .Row = i
            .Col = 4
            strname = Trim(.Text) '名称
   
            .Row = i
            .Col = 5
            strspec = Trim(.Text) '规格
   
            .Row = i
            .Col = 6
            strmodel = Trim(.Text) '型号
   
            .Row = i
            .Col = 7
            usage = CDbl(Trim(.Text)) '每只用量
   
            .Row = i
            .Col = 8
            strloss = CDbl(Trim(.Text)) '损耗
   
            .Row = i
            .Col = 9
            unit = Trim(.Text) '单位
   
            .Row = i
            .Col = 10
            strsite = Trim(.Text) '序号1
   
            .Row = i
            .Col = 11
            strtype = Trim(.Text) '材料类型
            
            .Row = i
            .Col = 12
            bom_group = Trim(.Text) '材料类型1
                    
            
            strSql = "INSERT INTO erpdata..TSVtblMRuleData(材料规范编号,工序号,料号,物料编号,名称,规格,型号,每只用量,损耗,单位,序号1,材料类型,材料类型1) " & _
            "  values ('" & bomId & "','" & strproduct & "','" & strmateriel & "','" & materiel_num & "','" & strname & "','" & strspec & "','" & strmodel & "' " & _
            "  ,'" & usage & "','" & strloss & "','" & unit & "','" & strsite & "','" & strtype & "','" & bom_group & "') "

            AddSql2 (strSql)
            AddSql2 (" UPDATE erpdata..TSVtblSetMRule SET 审核日期 = GETDATE(),材料类型 = '" & User & "'  WHERE 材料规范编号 = '" & bomId & "' ")
            
            Dim strSql_log As String
            '写log
            strSql_log = "INSERT INTO erpdata..TSVtblBom_Modify_log(修改日期,项目,材料规范编号,工序号,料号,物料编号,名称,规格,型号,每只用量,损耗,单位,序号1,材料类型,材料类型1,修改人员) " & _
            "  values (GETDATE(),'新增','" & bomId & "','" & strproduct & "','" & strmateriel & "','" & materiel_num & "','" & strname & "','" & strspec & "','" & strmodel & "' " & _
            "  ,'" & usage & "','" & strloss & "','" & unit & "','" & strsite & "','" & strtype & "','" & bom_group & "','" & User & "') "

            AddSql2 (strSql_log)
         
            MsgBox "添加成功!", vbInformation, "友情提示"
    
            
     
        Next i

    End With
    cmdQu_Click '查询
End Sub


Private Sub CmdBomModify_Click() '单笔修改
    Dim usage        As String
    
    Dim qtyTemp        As String
    
    Dim strmateriel As String

    Dim dzm As String
    
    Dim i   As Integer
    
    Dim sqlTemp_old_log As String
    
    Dim sqlTemp_new_log As String
    
    Dim sqlTemp As String
    
    Dim User As String
    
    User = gUserName
    
'    If Trim(TxtUsage.Text) = "" Then

        With Fps(0)

            For i = 1 To .MaxRows
                .Row = i
                .Col = E_FPS1.E_SEL

                If .Text <> "" Then
                    If .Text = 1 Then
            
                        .Col = 2
                        strmateriel = Trim(.Text) '料号
                    
                        .Col = E_FPS1.E_QTY
                        usage = Trim$(.Text) '每只用量
                        
 
                        .Col = E_FPS1.E_Typeid
                        dzm = Trim$(.Text) '序号1
                        
                        
                        If IsNumeric(usage) = False Then
                        
                            MsgBox "料号" & strmateriel & "的用量填写有误，请修正后再提交", vbInformation, "提示"
                            Exit Sub
                        End If
                        
                        If usage <= 0 Then
                        
                            MsgBox "料号" & strmateriel & "的用量填写有误，请修正后再提交", vbInformation, "提示"
                            Exit Sub
                        End If
                        
                        If dzm = "" Then
                            sqlTemp_old_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'修改用量前',*,'" & User & "'  from  erpdata..TSVtblMRuleData where 材料规范编号='" & bomId & "' and 料号='" & strmateriel & "' and (序号1 is  null or 序号1='') "
                            sqlTemp = "Update  [erpdata].[dbo].[TSVtblMRuleData]  Set 每只用量 = " & usage & "   where 材料规范编号='" & bomId & "' and 料号='" & strmateriel & "' and (序号1 is  null or 序号1='')  "
                            sqlTemp_new_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'修改用量后',*,'" & User & "'  from  erpdata..TSVtblMRuleData where 材料规范编号='" & bomId & "' and 料号='" & strmateriel & "' and (序号1 is  null or 序号1='') "
 
                        Else
                            sqlTemp_old_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'修改用量前',*,'" & User & "'  from  erpdata..TSVtblMRuleData where 材料规范编号='" & bomId & "' and 料号='" & strmateriel & "' and 序号1 = '" & dzm & "'"
                            sqlTemp = "Update  [erpdata].[dbo].[TSVtblMRuleData]  Set 每只用量 = " & usage & "   where 材料规范编号='" & bomId & "' and 料号='" & strmateriel & "' and 序号1 = '" & dzm & "'"
                            sqlTemp_new_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'修改用量后',*,'" & User & "'  from  erpdata..TSVtblMRuleData where 材料规范编号='" & bomId & "' and 料号='" & strmateriel & "' and 序号1 = '" & dzm & "'"

                        End If
                        AddSql2 (sqlTemp_old_log) ''修改前数据备份
                        AddSql2 (sqlTemp) '修改
                        AddSql2 (" UPDATE erpdata..TSVtblSetMRule SET 审核日期 = GETDATE(),材料类型 = '" & User & "'  WHERE 材料规范编号 = '" & bomId & "' ")
                        AddSql2 (sqlTemp_new_log) '修改后数据备份
                        

                    End If
            
                End If

            Next i

        End With

        '
        ''    MsgBox "用量不可以为空！", vbInformation, "友情提示"
        '    Exit Sub
 
    ' Else
        ' qtyTemp = Val(Trim(TxtUsage.Text))
    
        ' With fps(0)

            ' For i = 1 To .MaxRows
                ' .Row = i
                ' .Col = E_FPS1.E_SEL

                ' If .Text <> "" Then
                    ' If .Text = 1 Then
                        ' .Col = 1
                        ' bomIDtTemp = Trim(.Text) '材料规范编号
            
                        ' .Col = 2
                        ' strmateriel = Trim(.Text) '料号
            
                        ' sqlTemp = "Update  [erpdata].[dbo].[TSVtblMRuleData]  Set 每只用量 = " & qtyTemp & "   where 材料规范编号='" & bomID & "' and 料号='" & strmateriel & "'"
                        ' AddSql2 (sqlTemp)

                    ' End If

                ' End If

            ' Next i

        ' End With

    ' End If

    cmdQu_Click
End Sub



Private Sub cmdExp_Click()
Dim strSql As String
Dim product As String

product = Replace(Trim(CobPN.Text), Chr(13) + Chr(10), "")

If Len(Replace(Trim(CobPN.Text), Chr(13) + Chr(10), "")) = 0 Then
      strSql = " SELECT a.工序号,a.料号,a.物料编号,a.名称,a.规格,a.型号, CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,8),  a.每只用量)) as 每只用量" & _
             " ,CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,4),  a.损耗)) as 损耗,a.单位,a.序号1 as BOM站点,a.材料类型,a.材料类型1 " & _
             " FROM   erpdata..TSVtblMRuleData a,erpdata..TSVtblSetMRule b WHERE 1 = 1  AND b.材料规范编号 = a.材料规范编号 "
    ' 物料编号
    If Trim(TextWLBH.Text) <> "" Then
 
        strSql = strSql & " and a.料号  = '" & Trim(TextWLBH.Text) & "' "
 
    End If
    
    '站点
    If Trim(CoPZD.Text) <> "" Then
        strSql = strSql & " and a.序号1 = '" & Trim(CoPZD.Text) & "' "

    End If
    
    '审核人员
'    If Trim(txtText4.Text) <> "" Then
'        strSql = strSql & "and left(b.审核,5) = '" & Trim(txtText4.Text) & "'"
'    End If
'
'    If Trim(txtText5.Text) <> "" Then
'        strSql = strSql & "and b.建立日期 = '" & Trim(txtText5.Text) & "'"
'    End If
    
    If Trim(TextWLBH.Text) = "" And Trim(CobPN.Text) = "" And Trim(CoPZD.Text) = "" Then
        MsgBox "成品料号，料号和站点不可全空"
        Exit Sub
    End If
    
    If Trim(TextWLBH.Text) = "" And Trim(CobPN.Text) = "" And Trim(CoPZD.Text) <> "" Then
        If MsgBox("提示：请在输入料号或成品料号，不然可能会卡" & "是否继续？", vbOKCancel, "提示") <> vbOK Then
        Exit Sub
        End If
    End If
    strSql = strSql + "order by b.物料编号,a.料号,a.序号1"
    SqlServerExporToExcel (strSql)
  
   Exit Sub
End If

  strSql = " SELECT a.工序号,a.料号,a.物料编号,a.名称,a.规格,a.型号, CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,8),  a.每只用量)) as 每只用量" & _
                 " ,CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,4),  a.损耗)) as 损耗,a.单位,a.序号1 as BOM站点,a.材料类型,a.材料类型1 " & _
                "  FROM   erpdata..TSVtblMRuleData a,erpdata..TSVtblSetMRule b WHERE a.工序号 IN ('" & product & "')  AND b.材料规范编号 = a.材料规范编号 "
       
    If Trim(TextWLBH.Text) <> "" Then
 
        strSql = strSql & " and a.料号  = '" & Trim(TextWLBH.Text) & "' "
 
    End If
    
    '站点
    If Trim(CoPZD.Text) <> "" Then
        strSql = strSql & " and a.序号1 = '" & Trim(CoPZD.Text) & "' "

    End If
    SqlServerExporToExcel (strSql)
    
End Sub



Private Sub cmdQu_Click()

If Len(Replace(Trim(CobPN.Text), Chr(13) + Chr(10), "")) = 0 Then
CmdBomModify.Visible = False
CmdBomAdd.Visible = False
CmdBomAddSave.Visible = False
CmdBomDel.Visible = False
Query1
Exit Sub
End If


Query (Replace(Trim(CobPN.Text), Chr(13) + Chr(10), ""))

End Sub

Private Sub cmdup_Click()

    CommonDialog1.Filter = "所有文件(*.*)|*.*|Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.filename = "" Then
        Exit Sub

    End If

    txtPath.Text = CommonDialog1.filename

    CommonDialog1.filename = ""
    
    If txtPath.Text = "" Then
        MsgBox "请选择要上传的文件", vbInformation, "提示"
        Exit Sub

    End If
    

    Call Upload_0


End Sub


Private Sub Upload_0()
    On Error GoTo ErrHandle

    Dim VBExcel     As Excel.Application

    Dim xlBook      As Excel.Workbook

    Dim xlSheet     As Excel.Worksheet
    
    Dim strproduct  As String

    Dim strmateriel As String
    
    Dim materiel_num As String
    
    Dim strmateriel_old As String

    Dim strname  As String
    
    Dim strspec  As String

    Dim strmodel As String

    Dim usage As String
    
    Dim unit As String
    
    Dim strloss  As String

    Dim strsite As String

    Dim strtype  As String
    
    Dim bom_group As String
    
    Dim User As String
    
    Dim iRes  As Integer
     
    Dim rs   As New ADODB.Recordset

    Dim strSql   As String
    
    Dim old     As Integer
    
    Dim pro_bom As Integer
    
    Dim recordNo As String
    
    Dim i As Integer
    
    Dim J As Integer
    
    Dim up_flag As Integer
    
    Dim strsite_list  As String
    
    Dim strsite_temp  As String
 
    Dim strsite_match  As Boolean
    User = gUserName
    Fps(0).MaxRows = 0
    strmateriel_old = ""
    strproduct = ""
        
    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(txtPath.Text)
    Set xlSheet = xlBook.Worksheets(1)
 
    If xlSheet.Range("A1").CurrentRegion.Columns.count <> 10 Then
        MsgBox "Excel中的列数和设定的模版列数不一致，请确认Excel是否正确！", vbInformation, "提示"
        GoTo EXITPRO
        Exit Sub

    End If
    'Merry 20191104判断流程strsite是否在MES中存在
    strsite = ""
    strsite_list = ""
    strsite_temp = ""
    
    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        strsite = Replace(Trim(xlSheet.Range("I" & i)), Chr(13) + Chr(10), "") '序号1
        If strsite = "" Then
            MsgBox "第" & i & "行站别未填写，请填写后再上传！", vbInformation, "提示"
        End If

        strsite_match = False
        For J = 0 To UBound(Split(strsite_list, ","))
            strsite_temp = Split(strsite_list, ",")(J)
            If strsite = strsite_temp Then
                strsite_match = True
                Exit For
            End If
        Next
        If strsite_match = False Then
            If strsite_list = "" Then
                strsite_list = strsite
            Else
                strsite_list = strsite_list & "," & strsite
            End If
        End If
    Next

    For J = 0 To UBound(Split(strsite_list, ","))
        strsite_temp = Split(strsite_list, ",")(J)
        If CheckProc(url & strsite_temp) = "NG" Then
            MsgBox strsite_temp & " 站别在MES中不存在，请修改后再上传！", vbInformation, "提示"
            GoTo EXITPRO
            Exit Sub
        End If
    Next
    
    strsite = ""
    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
       
     If strproduct <> Replace(Trim(xlSheet.Range("A" & i)), Chr(13) + Chr(10), "") Then
        
        up_flag = 0
        
        strproduct = Replace(Trim(xlSheet.Range("A" & i)), Chr(13) + Chr(10), "")
        
'
        If Not JudgeBomProduct(strproduct) Then

            MsgBox "成品料号不对：" & strproduct & "，请确认!", vbInformation, "友情提示"
            GoTo EXITPRO
            Exit Sub

        End If
     
        
        
         Dim adoprm1      As ADODB.Parameter
         
            Dim adoPrmReturn As ADODB.Parameter

            Set adoCmd = New ADODB.Command
            Set adoCmd.ActiveConnection = INIadoCon2
            adoCmd.CommandText = "TSVuspgy_setmIndex "
            adoCmd.Parameters.Refresh
            adoCmd.CommandType = adCmdStoredProc

            Set adoPrmReturn = New ADODB.Parameter
            adoPrmReturn.Type = adChar
            adoPrmReturn.Size = 12
            adoPrmReturn.Direction = adParamOutput
            adoPrmReturn.Value = adParamReturnValue
            adoCmd.Parameters.Append adoPrmReturn
            adoCmd.Execute
            recordNo = adoPrmReturn.Value
            
            
            
            
      pro_bom = Get_SqlserverCnt(" SELECT * FROM erpdata..TSVtblSetMRule a WHERE a.物料编号 = '" & strproduct & "'  ")
        
        If pro_bom > 0 Then

        iRes = MsgBox("成品料号已存在BOM，请确认是否更新组件!", vbYesNoCancel, "提示:")
        If iRes <> vbYes Then
         GoTo EXITPRO
         Exit Sub

        Else
        
        AddSql2 (" DELETE FROM erpdata..TSVtblMRuleData WHERE 工序号 = '" & strproduct & "' ")
        AddSql2 (" UPDATE erpdata..TSVtblSetMRule SET 材料规范编号 = '" & recordNo & "' , 审核日期 = GETDATE(),材料类型 = '" & User & "'  WHERE 物料编号 = '" & strproduct & "' ")
        
        
        End If
      Else
        
      AddSql2 ("INSERT INTO erpdata..TSVtblSetMRule(材料规范编号,工艺,建立日期,状态标记,是否共用标记,物料编号,产线标记) " & _
                " values ('" & recordNo & "','" & User & "',GETDATE(),0,0,'" & strproduct & "',1)")
        
        
      End If
     End If
         strmateriel = Replace(Trim(xlSheet.Range("B" & i)), Chr(13) + Chr(10), "")
         
    
        If Not JudgeBomProduct(strmateriel) Then

            MsgBox "半成品料号不对：" & strmateriel & "，请确认!", vbInformation, "友情提示"
            GoTo EXITPRO
            Exit Sub

        End If
        
        old = Get_SqlserverCnt(" SELECT * FROM erptemp..bom_substitutes a  WHERE a.materiel_1 = '" & strmateriel & "' ")
        If old > 0 Then

           MsgBox "组件" & strmateriel & "存在新料号,请使用新料号导入!", vbInformation, "提示"
           GoTo EXITPRO
           Exit Sub

         End If
         
           
             
           strSql = "  SELECT a.FNumber,isnull(b.materiel_1,''),isnull(b.sub_code,'') ,isnull(c.计量单位名称,' '), isnull(a.FModel,' '),isnull(a.F_103,' ')   FROM AIS20141114094336.dbo.t_ICItem a  LEFT JOIN erptemp..bom_substitutes b   ON b.materiel_2 = a.F_101  " & _
           "   LEFT JOIN ERPBASE..tblUnitData c ON c.结构编码 = a.FProductUnitID WHERE a.F_101 ='" & strmateriel & "'"

            If rs.State = adStateOpen Then rs.Close
            rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not rs.EOF Then
            
            materiel_num = rs.Fields(0).Value '物料编号
            strmateriel_old = rs.Fields(1).Value
            bom_group = rs.Fields(2).Value '材料类型1
            unit = rs.Fields(3).Value '单位
            strspec = rs.Fields(4).Value '规格
            strmodel = rs.Fields(5).Value '型号
            End If
            
        strname = Replace(Trim(xlSheet.Range("C" & i)), Chr(13) + Chr(10), "") '名称
        'strspec = Replace(Trim(xlSheet.Range("D" & i)), Chr(13) + Chr(10), "")
        'strmodel = Replace(Trim(xlSheet.Range("E" & i)), Chr(13) + Chr(10), "")
        usage = Replace(Trim(xlSheet.Range("F" & i)), Chr(13) + Chr(10), "") '每只用量
        strloss = Replace(Trim(xlSheet.Range("H" & i)), Chr(13) + Chr(10), "") '损耗
        strsite = Replace(Trim(xlSheet.Range("I" & i)), Chr(13) + Chr(10), "") '序号1
        strtype = Replace(Trim(xlSheet.Range("J" & i)), Chr(13) + Chr(10), "") '材料类型
                
        If usage <> "0" Then
       
            If strmateriel_old <> "" Then

                AddSql2 ("INSERT INTO erpdata..TSVtblMRuleData(材料规范编号,工序号,料号,物料编号,名称,规格,型号,每只用量,损耗,单位,序号1,材料类型,材料类型1) " & _
                "  SELECT '" & recordNo & "','" & strproduct & "','" & strmateriel_old & "',c.FNumber, c.FName,c.FModel,c.F_103,'" & usage & "','" & strloss & "' " & _
                "  ,'" & unit & "','" & strsite & "','" & strtype & "','" & bom_group & "' FROM AIS20141114094336..t_ICItem c WHERE c.F_101 IN ('" & strmateriel_old & "') ")

            End If
'
            AddSql2 ("INSERT INTO erpdata..TSVtblMRuleData(材料规范编号,工序号,料号,物料编号,名称,规格,型号,每只用量,损耗,单位,序号1,材料类型,材料类型1) " & _
            "  values ('" & recordNo & "','" & strproduct & "','" & strmateriel & "','" & materiel_num & "','" & strname & "','" & strspec & "','" & strmodel & "' " & _
            "  ,'" & usage & "','" & strloss & "','" & unit & "','" & strsite & "','" & strtype & "','" & bom_group & "') ")
       
        End If
    Next
    
    MsgBox "上传完成", vbInformation, "提示"
    Query (strproduct)
   
EXITUPLOAD:

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
   
    Exit Sub
EXITPRO:

    On Error Resume Next

    MousePointer = 0
    
    MsgBox "成品料号" & strproduct & "组件" & strmateriel & "数据异常上传失败", vbInformation, "提示"

    If Not VBExcel Is Nothing Then

        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing

    End If

    Exit Sub
ErrHandle:
    GoTo EXITPRO

End Sub

Private Sub Query(product As String)
       
    Dim rs         As New ADODB.Recordset

    Dim strSql     As String
     
    Dim SMR        As New ADODB.Recordset
    
    CmdBomModify.Visible = True
    CmdBomAdd.Visible = True
    CmdBomAddSave.Visible = True
    CmdBomDel.Visible = True
    
    
    'merry 20191009查询时显示上传人员及审核日期
    strSql = "SELECT DISTINCT 材料规范编号,isnull(材料类型,'') as 材料类型 ,isnull(审核日期,'') as 审核日期  FROM erpdata..TSVtblSetMRule where 物料编号 IN ('" & product & "')"
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount = 1 Then
        SMR.MoveFirst
        txtText4.Text = SMR("材料类型")
        txtText5.Text = SMR("审核日期")
        bomId = Replace(Trim(SMR("材料规范编号")), Chr(13) + Chr(10), "")
    End If
    SMR.Close
    Set SMR = Nothing
    
    strSql = " SELECT a.工序号,a.料号,a.物料编号,a.名称,a.规格,a.型号, CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,8),  a.每只用量)) as 每只用量" & _
                 " ,CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,4),  a.损耗)) as 损耗,a.单位,a.序号1 as BOM站点,a.材料类型,a.材料类型1 " & _
                "  FROM   erpdata..TSVtblMRuleData a,erpdata..TSVtblSetMRule b WHERE a.工序号 IN ('" & product & "')  AND b.材料规范编号 = a.材料规范编号 "
       
    If Trim(TextWLBH.Text) <> "" Then
 
        strSql = strSql & " and a.料号  = '" & Trim(TextWLBH.Text) & "' "
 
    End If
    
    '站点
    If Trim(CoPZD.Text) <> "" Then
        strSql = strSql & " and a.序号1 = '" & Trim(CoPZD.Text) & "' "

    End If
    
    '审核人员
'    If Trim(txtText4.Text) <> "" Then
'        strSql = strSql & "and left(b.审核,5) = '" & Trim(txtText4.Text) & "'"
'    End If
    
'    If Trim(txtText5.Text) <> "" Then
'        strSql = strSql & "and b.建立日期 = '" & Trim(txtText5.Text) & "'"
'    End If
'
'    strSql = strSql + "order by b.物料编号,a.料号,a.序号1"
'
    Fps(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs.EOF Then  '表示有数据了
        Call ListDataType(rs)
    Else
        
        MsgBox "无成品料号" & product & "的BOM信息", vbInformation, "提示"
        Exit Sub

    End If

End Sub

Private Sub Query1()
       
    Dim rs         As New ADODB.Recordset

    Dim strSql     As String
     
    Dim SMR        As New ADODB.Recordset
    
    'merry 20191009查询时显示上传人员及审核日期
    strSql = " SELECT a.工序号,a.料号,a.物料编号,a.名称,a.规格,a.型号, CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,8),  a.每只用量)) as 每只用量" & _
             " ,CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,4),  a.损耗)) as 损耗,a.单位,a.序号1 as BOM站点,a.材料类型,a.材料类型1 " & _
             " FROM   erpdata..TSVtblMRuleData a,erpdata..TSVtblSetMRule b WHERE 1 = 1  AND b.材料规范编号 = a.材料规范编号 "
    ' 物料编号
    If Trim(TextWLBH.Text) <> "" Then
 
        strSql = strSql & " and a.料号  = '" & Trim(TextWLBH.Text) & "' "
 
    End If
    
    '站点
    If Trim(CoPZD.Text) <> "" Then
        strSql = strSql & " and a.序号1 = '" & Trim(CoPZD.Text) & "' "

    End If
    
    '审核人员
'    If Trim(txtText4.Text) <> "" Then
'        strSql = strSql & "and left(b.审核,5) = '" & Trim(txtText4.Text) & "'"
'    End If
'
'    If Trim(txtText5.Text) <> "" Then
'        strSql = strSql & "and b.建立日期 = '" & Trim(txtText5.Text) & "'"
'    End If
    
    If Trim(TextWLBH.Text) = "" And Trim(CobPN.Text) = "" And Trim(CoPZD.Text) = "" Then
        MsgBox "成品料号，料号和站点不可全空"
        Exit Sub
    End If
    
    If Trim(TextWLBH.Text) = "" And Trim(CobPN.Text) = "" And Trim(CoPZD.Text) <> "" Then
        If MsgBox("提示：请在输入料号或成品料号，不然可能会很卡" & "是否继续？", vbOKCancel, "提示") <> vbOK Then
        Exit Sub
        End If
    End If
    strSql = strSql + "order by b.物料编号,a.料号,a.序号1"
     
    Fps(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs.EOF Then  '表示有数据了
        Call ListDataType(rs)
    Else
        
        MsgBox "无BOM信息", vbInformation, "提示"
        Exit Sub

    End If
    rs.Close

End Sub

Private Sub ListDataType(rs As ADODB.Recordset)

    Dim i As Long
    Dim K As Double
   
 
   
    With Fps(0)
        
        .MaxRows = 0
        

        Set .DataSource = rs

    End With
    
    With Fps(0)
        .MaxCols = .MaxCols + 1
       
        For i = 1 To .MaxRows
'            .Row = i
'            .Col = 6
'
'            .Text = Format(.Text, "#0.0000")
'            .Col = 7
'            .Text = Format(Trim$(.Text), "0.0000")
'            .ColWidth(1) = 2
'            .CellType = CellTypeCheckBox
'            .Text = 1


            .Row = i
            .Col = E_FPS1.E_QTY
            .Lock = False
            
            ' .Row = i
            ' .Col = E_FPS1.E_RATE
            ' .Lock = False
'------------------
            .Row = i
            .Col = E_FPS1.E_SEL
            .SetText E_FPS1.E_SEL, 0, "选择"
            .CellType = CellTypeCheckBox
            .Text = 0
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            .Lock = False
            
            .ReDraw = True
            
            

        Next
        FPSMaxRowBeforeAdd = .MaxRows '记录新增之前的最大行数，新增提交时只上传后面的行
    End With

End Sub

Private Sub Form_Load()

    With Fps(0)
    
        .Col = -1
        .Row = -1
        .Lock = True

    End With
    url = "http://10.160.2.30:9090/psb.web/api/v1/operations?operation="
End Sub

Private Sub fps_LeaveCell(Index As Integer, _
                          ByVal Col As Long, _
                          ByVal Row As Long, _
                          ByVal NewCol As Long, _
                          ByVal NewRow As Long, _
                          Cancel As Boolean)
                          
    On Error GoTo ErrHandle
    Dim oiRS         As New ADODB.Recordset

    Dim strSql     As String

    Dim strmateriel As String
    
    Dim old     As Integer
    If Row <= FPSMaxRowBeforeAdd Then Exit Sub

    If (Col = E_FPS1.E_PT And Row > FPSMaxRowBeforeAdd) Then

        With Fps(0)
            .Row = Row
            .Col = Col

            strmateriel = .Text
          '  bomProduct = bomProductTemp
        
            '根据料号，查询相关信息

'----------------------------------------------
        If Not JudgeBomProduct(strmateriel) Then

            MsgBox "半成品料号不对：" & strmateriel & "，请确认!", vbInformation, "友情提示"
            Exit Sub

        End If
        
        old = Get_SqlserverCnt(" SELECT * FROM erptemp..bom_substitutes a  WHERE a.materiel_1 = '" & strmateriel & "' ")
        If old > 0 Then

           MsgBox "组件" & strmateriel & "存在新料号,请使用新料号导入!", vbInformation, "提示"
           Exit Sub

         End If
         
           
             
           strSql = "  SELECT a.FNumber,a.FName,isnull(b.materiel_1,''),isnull(b.sub_code,'') ,isnull(c.计量单位名称,' '), isnull(a.FModel,' '),isnull(a.F_103,' ')   FROM AIS20141114094336.dbo.t_ICItem a  LEFT JOIN erptemp..bom_substitutes b   ON b.materiel_2 = a.F_101  " & _
           "   LEFT JOIN ERPBASE..tblUnitData c ON c.结构编码 = a.FProductUnitID WHERE a.F_101 ='" & strmateriel & "'"
       
            If oiRS.State = adStateOpen Then oiRS.Close
            oiRS.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not oiRS.EOF Then
            
                .Row = Row
                .Col = Col + 1
                .Text = Trim(oiRS.Fields(0).Value) '物料编码
                
                .Row = Row
                .Col = Col + 2
                .Text = Trim(oiRS.Fields(1).Value) '名称
                

                .Row = Row
                .Col = Col + 3
                .Text = Trim(oiRS.Fields(5).Value) '规格
            
                .Row = Row
                .Col = Col + 4
                .Text = Trim(oiRS.Fields(6).Value) '型号
                
                .Row = Row
                .Col = Col + 7
                .Text = Trim(oiRS.Fields(4).Value) '单位
            
                .Row = Row
                .Col = Col + 8
                .Text = Trim(oiRS.Fields(2).Value) '材料类型1
        
            End If
            oiRS.Close
            Set oiRS = Nothing

        End With

    End If
    
EXITPRO:

    On Error Resume Next


    If Not oiRS Is Nothing Then

        Set oiRS = Nothing

    End If

    Exit Sub
ErrHandle:
    GoTo EXITPRO

End Sub

Private Sub CmdBomDel_Click() '删除

    Dim strmateriel As String

    Dim dzm         As String
    
    Dim i           As Integer
    
    Dim sqlTemp     As String
    
    Dim sqlTemp_log As String
    
    Dim User        As String
    
    Dim DelCnt     As Integer
    
    Dim DelMaterial As String
    
    User = gUserName
    DelCnt = 0
    DelMaterial = ""
    With Fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = E_FPS1.E_SEL

            If .Text <> "" Then
                If .Text = 1 Then
            
                    .Col = E_FPS1.E_PT      '料号
                    If DelMaterial = "" Then
                        DelMaterial = Trim(.Text)
                    Else
                        DelMaterial = DelMaterial & "," & Trim(.Text)
                    End If
                    DelCnt = DelCnt + 1
                End If

            End If
        Next i
        If MsgBox("你确认要删除" & DelMaterial & ",共" & DelCnt & "笔物料吗?", vbOKCancel, "提示") = vbCancel Then
            Exit Sub

        End If
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = E_FPS1.E_SEL

            If .Text <> "" Then
                If .Text = 1 Then
            
                    .Col = E_FPS1.E_PT      '料号
                    strmateriel = Trim(.Text)
                
                    .Col = E_FPS1.E_Typeid    '序号1
                    dzm = Trim$(.Text)
                    
                    If dzm = "" Then
                        sqlTemp = "delete from  [erpdata].[dbo].[TSVtblMRuleData]  where 材料规范编号='" & bomId & "' and 料号='" & strmateriel & "' and (序号1 is  null or 序号1='') "
                        sqlTemp_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'删除',*,'" & User & "'  from  erpdata..TSVtblMRuleData where 材料规范编号='" & bomId & "' and 料号='" & strmateriel & "' and (序号1 is  null or 序号1='') "
                    Else
                        sqlTemp = "delete from  [erpdata].[dbo].[TSVtblMRuleData]  where 材料规范编号='" & bomId & "' and 料号='" & strmateriel & "' and 序号1 = '" & dzm & "'"
                        sqlTemp_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'删除',*,'" & User & "'  from  erpdata..TSVtblMRuleData where 材料规范编号='" & bomId & "' and 料号='" & strmateriel & "' and 序号1 = '" & dzm & "'"
                    End If
                    AddSql2 (sqlTemp_log)
                    AddSql2 (sqlTemp)
                    AddSql2 (" UPDATE erpdata..TSVtblSetMRule SET 审核日期 = GETDATE(),材料类型 = '" & User & "'  WHERE 材料规范编号 = '" & bomId & "' ")
        

                End If

            End If

        Next i

    End With
    cmdQu_Click '查询
End Sub


Private Sub txtMPN_DblClick()
    
    Dim SMR        As New ADODB.Recordset
    
    Dim strSql     As String
    
    Dim i          As Integer
    
    If txtmpn.Text = "" Then Exit Sub
    CobPN.Text = ""
    strSql = "SELECT  DISTINCT QTECHPTNO2 FROM erptemp .. tbltsvnpiproduct where QTECHPTNO='" & Trim$(txtmpn.Text) & "'"
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        If SMR.RecordCount = 1 Then CobPN.Text = Trim(SMR("QTECHPTNO2"))
        For i = 1 To SMR.RecordCount
            CobPN.AddItem (Trim(SMR("QTECHPTNO2")))
            SMR.MoveNext
        Next
    End If
    SMR.Close
    Set SMR = Nothing
    MousePointer = 0
End Sub

Private Sub txtMPN_Change()
    
    CobPN.Clear
End Sub


Private Function CheckProc(url As String)

Dim xmlHttp As Object
Dim XMLDoc As Object
Dim NGresult As String
Dim Result As String
Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
CheckProc = "OK"
xmlHttp.Open "GET", url, True
xmlHttp.Send (Null)
While xmlHttp.readyState <> 4
DoEvents
Wend
Result = xmlHttp.responseText
' 若MES不存在该站点,则可返回如下结果
' {
    ' "header": {
        ' "code": 0,
        ' "message": ""
    ' },
    ' "value": []
' }
NGresult = Chr(34) & "value" & Chr(34) & ":[]"
If InStr(Result, NGresult) Then
    CheckProc = "NG"
End If

End Function








