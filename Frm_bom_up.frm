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
         Caption         =   "ȫѡ/ȫ��ѡ"
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
         Caption         =   "�޸�"
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
            Caption         =   "ɾ��"
            Height          =   360
            Left            =   13200
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   480
            Width           =   990
         End
         Begin VB.CommandButton CmdBomAddSave 
            BackColor       =   &H000080FF&
            Caption         =   "��Ӻ��ύ"
            Height          =   360
            Left            =   7440
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton CmdBomAdd 
            Caption         =   "���һ��"
            Height          =   360
            Left            =   6120
            TabIndex        =   10
            Top             =   360
            Width           =   990
         End
         Begin VB.CommandButton CmdBomModify 
            BackColor       =   &H00C0C000&
            Caption         =   "�޸������ύ"
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
            Caption         =   "�޸�����"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdExp 
         Caption         =   "����"
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
         Caption         =   "��ѯ"
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
         Caption         =   "�����ϴ�"
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
         Caption         =   "�Ϻţ�"
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
         Caption         =   "BOMվ�㣺"
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
         Caption         =   "��������"
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
         Caption         =   "�ϴ�ʱ�䣺"
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
         Caption         =   "�����Ա��"
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
         Caption         =   "��Ʒ�Ϻ�:"
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

Public bomId         As String '���Ϲ淶���
Public url           As String

Private Enum E_FPS1          'Bom�֭�

    E_PRODUCTID = 1              '�����
    E_PT                     '�Ϻ�
    E_MATNUM                '���ϱ��
    E_name                   '���ƪ�
    E_GG                     '���
    E_XH                     '�ͺ�
    
    E_QTY                    'ÿֻ����
    E_Rate                   '�����
    E_UNIT                   '��λ
    
    E_Typeid                 '���1
    E_TypePT                 '��������
    E_TypePT1                '��������1
    E_SEL                    '��ѡ
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

Private Sub CmdBomAdd_Click() '����һ��
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
        .Text = strproduct '�����������һ��
        
        .Row = i
        .Col = 2
        .Lock = False    '�Ϻ����ɱ༭

        
        .Row = i
        .Col = 7
        .Lock = False    'ÿֻ�����ɱ༭
        .Text = "0.0000"
        
        .Row = i
        .Col = 8
        .Lock = False    '������ɱ༭
        .Text = "0.0000"
        
        .Row = i
        .Col = 10
        .Lock = False     '���1�ɱ༭
        
        .Row = i
        .Col = 10
        .Lock = False     '��������
        
        .Row = i
        .Col = 11
        .Lock = False    '��������1

    End With
End Sub

Private Sub CmdBomAddSave_Click() '��Ӻ��ύ
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
        MsgBox "��������֮���ٵ��ύ!", vbInformation, "��ʾ"
        Exit Sub
    End If

    With Fps(0)
    '�ȼ������
        For i = FPSMaxRowBeforeAdd + 1 To .MaxRows
   
            .Row = i
            .Col = 2
            strmateriel = Trim(.Text) '�Ϻ�

            .Row = i
            .Col = 3
            materiel_num = Trim(.Text) '���ϱ��

            .Row = i
            .Col = 11
            strtype = Trim(.Text) '��������
            
            .Row = i
            .Col = 7
            
            If IsNumeric(Trim(.Text)) = False Then
            
                MsgBox "�Ϻ�" & strmateriel & "��������д���������������ύ", vbInformation, "��ʾ"
                Exit Sub
                
            End If
            If Trim(.Text) <= 0 Then
            
                MsgBox "�Ϻ�" & strmateriel & "��������д���������������ύ", vbInformation, "��ʾ"
                Exit Sub
            
            End If

   
            .Row = i
            .Col = 8

            If IsNumeric(Trim(.Text)) = False Then
            
                MsgBox "�Ϻ�" & strmateriel & "�������д���������������ύ", vbInformation, "��ʾ"
                Exit Sub
                
            End If
            
            .Row = i
            .Col = 10
            strsite = Trim(.Text) '���1
            
            If strmateriel = "" Or materiel_num = "" Or strtype = "" Or strsite = "" Then
                MsgBox "�뽫���ݲ����������ύ!", vbInformation, "������ʾ"
                Exit Sub
            
            End If
            'merry 20191104�жϹ�����MES���Ƿ����
            If CheckProc(url & strsite) = "NG" Then
                MsgBox strsite & "վ����MES�в����ڣ����޸ĺ����ϴ���", vbInformation, "��ʾ"
                Exit Sub
            End If
            
        
        Next
        For i = FPSMaxRowBeforeAdd + 1 To .MaxRows
   
            .Row = i
            .Col = 1
            strproduct = Trim(.Text) '�����
   
            .Row = i
            .Col = 2
            strmateriel = Trim(.Text) '�Ϻ�
            
            
            .Row = i
            .Col = 3
            materiel_num = Trim(.Text) '���ϱ��
          
            .Row = i
            .Col = 4
            strname = Trim(.Text) '����
   
            .Row = i
            .Col = 5
            strspec = Trim(.Text) '���
   
            .Row = i
            .Col = 6
            strmodel = Trim(.Text) '�ͺ�
   
            .Row = i
            .Col = 7
            usage = CDbl(Trim(.Text)) 'ÿֻ����
   
            .Row = i
            .Col = 8
            strloss = CDbl(Trim(.Text)) '���
   
            .Row = i
            .Col = 9
            unit = Trim(.Text) '��λ
   
            .Row = i
            .Col = 10
            strsite = Trim(.Text) '���1
   
            .Row = i
            .Col = 11
            strtype = Trim(.Text) '��������
            
            .Row = i
            .Col = 12
            bom_group = Trim(.Text) '��������1
                    
            
            strSql = "INSERT INTO erpdata..TSVtblMRuleData(���Ϲ淶���,�����,�Ϻ�,���ϱ��,����,���,�ͺ�,ÿֻ����,���,��λ,���1,��������,��������1) " & _
            "  values ('" & bomId & "','" & strproduct & "','" & strmateriel & "','" & materiel_num & "','" & strname & "','" & strspec & "','" & strmodel & "' " & _
            "  ,'" & usage & "','" & strloss & "','" & unit & "','" & strsite & "','" & strtype & "','" & bom_group & "') "

            AddSql2 (strSql)
            AddSql2 (" UPDATE erpdata..TSVtblSetMRule SET ������� = GETDATE(),�������� = '" & User & "'  WHERE ���Ϲ淶��� = '" & bomId & "' ")
            
            Dim strSql_log As String
            'дlog
            strSql_log = "INSERT INTO erpdata..TSVtblBom_Modify_log(�޸�����,��Ŀ,���Ϲ淶���,�����,�Ϻ�,���ϱ��,����,���,�ͺ�,ÿֻ����,���,��λ,���1,��������,��������1,�޸���Ա) " & _
            "  values (GETDATE(),'����','" & bomId & "','" & strproduct & "','" & strmateriel & "','" & materiel_num & "','" & strname & "','" & strspec & "','" & strmodel & "' " & _
            "  ,'" & usage & "','" & strloss & "','" & unit & "','" & strsite & "','" & strtype & "','" & bom_group & "','" & User & "') "

            AddSql2 (strSql_log)
         
            MsgBox "��ӳɹ�!", vbInformation, "������ʾ"
    
            
     
        Next i

    End With
    cmdQu_Click '��ѯ
End Sub


Private Sub CmdBomModify_Click() '�����޸�
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
                        strmateriel = Trim(.Text) '�Ϻ�
                    
                        .Col = E_FPS1.E_QTY
                        usage = Trim$(.Text) 'ÿֻ����
                        
 
                        .Col = E_FPS1.E_Typeid
                        dzm = Trim$(.Text) '���1
                        
                        
                        If IsNumeric(usage) = False Then
                        
                            MsgBox "�Ϻ�" & strmateriel & "��������д���������������ύ", vbInformation, "��ʾ"
                            Exit Sub
                        End If
                        
                        If usage <= 0 Then
                        
                            MsgBox "�Ϻ�" & strmateriel & "��������д���������������ύ", vbInformation, "��ʾ"
                            Exit Sub
                        End If
                        
                        If dzm = "" Then
                            sqlTemp_old_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'�޸�����ǰ',*,'" & User & "'  from  erpdata..TSVtblMRuleData where ���Ϲ淶���='" & bomId & "' and �Ϻ�='" & strmateriel & "' and (���1 is  null or ���1='') "
                            sqlTemp = "Update  [erpdata].[dbo].[TSVtblMRuleData]  Set ÿֻ���� = " & usage & "   where ���Ϲ淶���='" & bomId & "' and �Ϻ�='" & strmateriel & "' and (���1 is  null or ���1='')  "
                            sqlTemp_new_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'�޸�������',*,'" & User & "'  from  erpdata..TSVtblMRuleData where ���Ϲ淶���='" & bomId & "' and �Ϻ�='" & strmateriel & "' and (���1 is  null or ���1='') "
 
                        Else
                            sqlTemp_old_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'�޸�����ǰ',*,'" & User & "'  from  erpdata..TSVtblMRuleData where ���Ϲ淶���='" & bomId & "' and �Ϻ�='" & strmateriel & "' and ���1 = '" & dzm & "'"
                            sqlTemp = "Update  [erpdata].[dbo].[TSVtblMRuleData]  Set ÿֻ���� = " & usage & "   where ���Ϲ淶���='" & bomId & "' and �Ϻ�='" & strmateriel & "' and ���1 = '" & dzm & "'"
                            sqlTemp_new_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'�޸�������',*,'" & User & "'  from  erpdata..TSVtblMRuleData where ���Ϲ淶���='" & bomId & "' and �Ϻ�='" & strmateriel & "' and ���1 = '" & dzm & "'"

                        End If
                        AddSql2 (sqlTemp_old_log) ''�޸�ǰ���ݱ���
                        AddSql2 (sqlTemp) '�޸�
                        AddSql2 (" UPDATE erpdata..TSVtblSetMRule SET ������� = GETDATE(),�������� = '" & User & "'  WHERE ���Ϲ淶��� = '" & bomId & "' ")
                        AddSql2 (sqlTemp_new_log) '�޸ĺ����ݱ���
                        

                    End If
            
                End If

            Next i

        End With

        '
        ''    MsgBox "����������Ϊ�գ�", vbInformation, "������ʾ"
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
                        ' bomIDtTemp = Trim(.Text) '���Ϲ淶���
            
                        ' .Col = 2
                        ' strmateriel = Trim(.Text) '�Ϻ�
            
                        ' sqlTemp = "Update  [erpdata].[dbo].[TSVtblMRuleData]  Set ÿֻ���� = " & qtyTemp & "   where ���Ϲ淶���='" & bomID & "' and �Ϻ�='" & strmateriel & "'"
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
      strSql = " SELECT a.�����,a.�Ϻ�,a.���ϱ��,a.����,a.���,a.�ͺ�, CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,8),  a.ÿֻ����)) as ÿֻ����" & _
             " ,CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,4),  a.���)) as ���,a.��λ,a.���1 as BOMվ��,a.��������,a.��������1 " & _
             " FROM   erpdata..TSVtblMRuleData a,erpdata..TSVtblSetMRule b WHERE 1 = 1  AND b.���Ϲ淶��� = a.���Ϲ淶��� "
    ' ���ϱ��
    If Trim(TextWLBH.Text) <> "" Then
 
        strSql = strSql & " and a.�Ϻ�  = '" & Trim(TextWLBH.Text) & "' "
 
    End If
    
    'վ��
    If Trim(CoPZD.Text) <> "" Then
        strSql = strSql & " and a.���1 = '" & Trim(CoPZD.Text) & "' "

    End If
    
    '�����Ա
'    If Trim(txtText4.Text) <> "" Then
'        strSql = strSql & "and left(b.���,5) = '" & Trim(txtText4.Text) & "'"
'    End If
'
'    If Trim(txtText5.Text) <> "" Then
'        strSql = strSql & "and b.�������� = '" & Trim(txtText5.Text) & "'"
'    End If
    
    If Trim(TextWLBH.Text) = "" And Trim(CobPN.Text) = "" And Trim(CoPZD.Text) = "" Then
        MsgBox "��Ʒ�Ϻţ��Ϻź�վ�㲻��ȫ��"
        Exit Sub
    End If
    
    If Trim(TextWLBH.Text) = "" And Trim(CobPN.Text) = "" And Trim(CoPZD.Text) <> "" Then
        If MsgBox("��ʾ�����������ϺŻ��Ʒ�Ϻţ���Ȼ���ܻῨ" & "�Ƿ������", vbOKCancel, "��ʾ") <> vbOK Then
        Exit Sub
        End If
    End If
    strSql = strSql + "order by b.���ϱ��,a.�Ϻ�,a.���1"
    SqlServerExporToExcel (strSql)
  
   Exit Sub
End If

  strSql = " SELECT a.�����,a.�Ϻ�,a.���ϱ��,a.����,a.���,a.�ͺ�, CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,8),  a.ÿֻ����)) as ÿֻ����" & _
                 " ,CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,4),  a.���)) as ���,a.��λ,a.���1 as BOMվ��,a.��������,a.��������1 " & _
                "  FROM   erpdata..TSVtblMRuleData a,erpdata..TSVtblSetMRule b WHERE a.����� IN ('" & product & "')  AND b.���Ϲ淶��� = a.���Ϲ淶��� "
       
    If Trim(TextWLBH.Text) <> "" Then
 
        strSql = strSql & " and a.�Ϻ�  = '" & Trim(TextWLBH.Text) & "' "
 
    End If
    
    'վ��
    If Trim(CoPZD.Text) <> "" Then
        strSql = strSql & " and a.���1 = '" & Trim(CoPZD.Text) & "' "

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

    CommonDialog1.Filter = "�����ļ�(*.*)|*.*|Excel�ļ�(*.xls;*.xlsx)|*.xls;*.xlsx"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.filename = "" Then
        Exit Sub

    End If

    txtPath.Text = CommonDialog1.filename

    CommonDialog1.filename = ""
    
    If txtPath.Text = "" Then
        MsgBox "��ѡ��Ҫ�ϴ����ļ�", vbInformation, "��ʾ"
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
        MsgBox "Excel�е��������趨��ģ��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        GoTo EXITPRO
        Exit Sub

    End If
    'Merry 20191104�ж�����strsite�Ƿ���MES�д���
    strsite = ""
    strsite_list = ""
    strsite_temp = ""
    
    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.count
        strsite = Replace(Trim(xlSheet.Range("I" & i)), Chr(13) + Chr(10), "") '���1
        If strsite = "" Then
            MsgBox "��" & i & "��վ��δ��д������д�����ϴ���", vbInformation, "��ʾ"
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
            MsgBox strsite_temp & " վ����MES�в����ڣ����޸ĺ����ϴ���", vbInformation, "��ʾ"
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

            MsgBox "��Ʒ�ϺŲ��ԣ�" & strproduct & "����ȷ��!", vbInformation, "������ʾ"
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
            
            
            
            
      pro_bom = Get_SqlserverCnt(" SELECT * FROM erpdata..TSVtblSetMRule a WHERE a.���ϱ�� = '" & strproduct & "'  ")
        
        If pro_bom > 0 Then

        iRes = MsgBox("��Ʒ�Ϻ��Ѵ���BOM����ȷ���Ƿ�������!", vbYesNoCancel, "��ʾ:")
        If iRes <> vbYes Then
         GoTo EXITPRO
         Exit Sub

        Else
        
        AddSql2 (" DELETE FROM erpdata..TSVtblMRuleData WHERE ����� = '" & strproduct & "' ")
        AddSql2 (" UPDATE erpdata..TSVtblSetMRule SET ���Ϲ淶��� = '" & recordNo & "' , ������� = GETDATE(),�������� = '" & User & "'  WHERE ���ϱ�� = '" & strproduct & "' ")
        
        
        End If
      Else
        
      AddSql2 ("INSERT INTO erpdata..TSVtblSetMRule(���Ϲ淶���,����,��������,״̬���,�Ƿ��ñ��,���ϱ��,���߱��) " & _
                " values ('" & recordNo & "','" & User & "',GETDATE(),0,0,'" & strproduct & "',1)")
        
        
      End If
     End If
         strmateriel = Replace(Trim(xlSheet.Range("B" & i)), Chr(13) + Chr(10), "")
         
    
        If Not JudgeBomProduct(strmateriel) Then

            MsgBox "���Ʒ�ϺŲ��ԣ�" & strmateriel & "����ȷ��!", vbInformation, "������ʾ"
            GoTo EXITPRO
            Exit Sub

        End If
        
        old = Get_SqlserverCnt(" SELECT * FROM erptemp..bom_substitutes a  WHERE a.materiel_1 = '" & strmateriel & "' ")
        If old > 0 Then

           MsgBox "���" & strmateriel & "�������Ϻ�,��ʹ�����Ϻŵ���!", vbInformation, "��ʾ"
           GoTo EXITPRO
           Exit Sub

         End If
         
           
             
           strSql = "  SELECT a.FNumber,isnull(b.materiel_1,''),isnull(b.sub_code,'') ,isnull(c.������λ����,' '), isnull(a.FModel,' '),isnull(a.F_103,' ')   FROM AIS20141114094336.dbo.t_ICItem a  LEFT JOIN erptemp..bom_substitutes b   ON b.materiel_2 = a.F_101  " & _
           "   LEFT JOIN ERPBASE..tblUnitData c ON c.�ṹ���� = a.FProductUnitID WHERE a.F_101 ='" & strmateriel & "'"

            If rs.State = adStateOpen Then rs.Close
            rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not rs.EOF Then
            
            materiel_num = rs.Fields(0).Value '���ϱ��
            strmateriel_old = rs.Fields(1).Value
            bom_group = rs.Fields(2).Value '��������1
            unit = rs.Fields(3).Value '��λ
            strspec = rs.Fields(4).Value '���
            strmodel = rs.Fields(5).Value '�ͺ�
            End If
            
        strname = Replace(Trim(xlSheet.Range("C" & i)), Chr(13) + Chr(10), "") '����
        'strspec = Replace(Trim(xlSheet.Range("D" & i)), Chr(13) + Chr(10), "")
        'strmodel = Replace(Trim(xlSheet.Range("E" & i)), Chr(13) + Chr(10), "")
        usage = Replace(Trim(xlSheet.Range("F" & i)), Chr(13) + Chr(10), "") 'ÿֻ����
        strloss = Replace(Trim(xlSheet.Range("H" & i)), Chr(13) + Chr(10), "") '���
        strsite = Replace(Trim(xlSheet.Range("I" & i)), Chr(13) + Chr(10), "") '���1
        strtype = Replace(Trim(xlSheet.Range("J" & i)), Chr(13) + Chr(10), "") '��������
                
        If usage <> "0" Then
       
            If strmateriel_old <> "" Then

                AddSql2 ("INSERT INTO erpdata..TSVtblMRuleData(���Ϲ淶���,�����,�Ϻ�,���ϱ��,����,���,�ͺ�,ÿֻ����,���,��λ,���1,��������,��������1) " & _
                "  SELECT '" & recordNo & "','" & strproduct & "','" & strmateriel_old & "',c.FNumber, c.FName,c.FModel,c.F_103,'" & usage & "','" & strloss & "' " & _
                "  ,'" & unit & "','" & strsite & "','" & strtype & "','" & bom_group & "' FROM AIS20141114094336..t_ICItem c WHERE c.F_101 IN ('" & strmateriel_old & "') ")

            End If
'
            AddSql2 ("INSERT INTO erpdata..TSVtblMRuleData(���Ϲ淶���,�����,�Ϻ�,���ϱ��,����,���,�ͺ�,ÿֻ����,���,��λ,���1,��������,��������1) " & _
            "  values ('" & recordNo & "','" & strproduct & "','" & strmateriel & "','" & materiel_num & "','" & strname & "','" & strspec & "','" & strmodel & "' " & _
            "  ,'" & usage & "','" & strloss & "','" & unit & "','" & strsite & "','" & strtype & "','" & bom_group & "') ")
       
        End If
    Next
    
    MsgBox "�ϴ����", vbInformation, "��ʾ"
    Query (strproduct)
   
EXITUPLOAD:

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
   
    Exit Sub
EXITPRO:

    On Error Resume Next

    MousePointer = 0
    
    MsgBox "��Ʒ�Ϻ�" & strproduct & "���" & strmateriel & "�����쳣�ϴ�ʧ��", vbInformation, "��ʾ"

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
    
    
    'merry 20191009��ѯʱ��ʾ�ϴ���Ա���������
    strSql = "SELECT DISTINCT ���Ϲ淶���,isnull(��������,'') as �������� ,isnull(�������,'') as �������  FROM erpdata..TSVtblSetMRule where ���ϱ�� IN ('" & product & "')"
    If SMR.State = adStateOpen Then SMR.Close
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount = 1 Then
        SMR.MoveFirst
        txtText4.Text = SMR("��������")
        txtText5.Text = SMR("�������")
        bomId = Replace(Trim(SMR("���Ϲ淶���")), Chr(13) + Chr(10), "")
    End If
    SMR.Close
    Set SMR = Nothing
    
    strSql = " SELECT a.�����,a.�Ϻ�,a.���ϱ��,a.����,a.���,a.�ͺ�, CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,8),  a.ÿֻ����)) as ÿֻ����" & _
                 " ,CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,4),  a.���)) as ���,a.��λ,a.���1 as BOMվ��,a.��������,a.��������1 " & _
                "  FROM   erpdata..TSVtblMRuleData a,erpdata..TSVtblSetMRule b WHERE a.����� IN ('" & product & "')  AND b.���Ϲ淶��� = a.���Ϲ淶��� "
       
    If Trim(TextWLBH.Text) <> "" Then
 
        strSql = strSql & " and a.�Ϻ�  = '" & Trim(TextWLBH.Text) & "' "
 
    End If
    
    'վ��
    If Trim(CoPZD.Text) <> "" Then
        strSql = strSql & " and a.���1 = '" & Trim(CoPZD.Text) & "' "

    End If
    
    '�����Ա
'    If Trim(txtText4.Text) <> "" Then
'        strSql = strSql & "and left(b.���,5) = '" & Trim(txtText4.Text) & "'"
'    End If
    
'    If Trim(txtText5.Text) <> "" Then
'        strSql = strSql & "and b.�������� = '" & Trim(txtText5.Text) & "'"
'    End If
'
'    strSql = strSql + "order by b.���ϱ��,a.�Ϻ�,a.���1"
'
    Fps(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType(rs)
    Else
        
        MsgBox "�޳�Ʒ�Ϻ�" & product & "��BOM��Ϣ", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

Private Sub Query1()
       
    Dim rs         As New ADODB.Recordset

    Dim strSql     As String
     
    Dim SMR        As New ADODB.Recordset
    
    'merry 20191009��ѯʱ��ʾ�ϴ���Ա���������
    strSql = " SELECT a.�����,a.�Ϻ�,a.���ϱ��,a.����,a.���,a.�ͺ�, CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,8),  a.ÿֻ����)) as ÿֻ����" & _
             " ,CONVERT(VARCHAR(100), CONVERT(DECIMAL(18,4),  a.���)) as ���,a.��λ,a.���1 as BOMվ��,a.��������,a.��������1 " & _
             " FROM   erpdata..TSVtblMRuleData a,erpdata..TSVtblSetMRule b WHERE 1 = 1  AND b.���Ϲ淶��� = a.���Ϲ淶��� "
    ' ���ϱ��
    If Trim(TextWLBH.Text) <> "" Then
 
        strSql = strSql & " and a.�Ϻ�  = '" & Trim(TextWLBH.Text) & "' "
 
    End If
    
    'վ��
    If Trim(CoPZD.Text) <> "" Then
        strSql = strSql & " and a.���1 = '" & Trim(CoPZD.Text) & "' "

    End If
    
    '�����Ա
'    If Trim(txtText4.Text) <> "" Then
'        strSql = strSql & "and left(b.���,5) = '" & Trim(txtText4.Text) & "'"
'    End If
'
'    If Trim(txtText5.Text) <> "" Then
'        strSql = strSql & "and b.�������� = '" & Trim(txtText5.Text) & "'"
'    End If
    
    If Trim(TextWLBH.Text) = "" And Trim(CobPN.Text) = "" And Trim(CoPZD.Text) = "" Then
        MsgBox "��Ʒ�Ϻţ��Ϻź�վ�㲻��ȫ��"
        Exit Sub
    End If
    
    If Trim(TextWLBH.Text) = "" And Trim(CobPN.Text) = "" And Trim(CoPZD.Text) <> "" Then
        If MsgBox("��ʾ�����������ϺŻ��Ʒ�Ϻţ���Ȼ���ܻ�ܿ�" & "�Ƿ������", vbOKCancel, "��ʾ") <> vbOK Then
        Exit Sub
        End If
    End If
    strSql = strSql + "order by b.���ϱ��,a.�Ϻ�,a.���1"
     
    Fps(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType(rs)
    Else
        
        MsgBox "��BOM��Ϣ", vbInformation, "��ʾ"
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
            .SetText E_FPS1.E_SEL, 0, "ѡ��"
            .CellType = CellTypeCheckBox
            .Text = 0
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            .Lock = False
            
            .ReDraw = True
            
            

        Next
        FPSMaxRowBeforeAdd = .MaxRows '��¼����֮ǰ����������������ύʱֻ�ϴ��������
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
        
            '�����Ϻţ���ѯ�����Ϣ

'----------------------------------------------
        If Not JudgeBomProduct(strmateriel) Then

            MsgBox "���Ʒ�ϺŲ��ԣ�" & strmateriel & "����ȷ��!", vbInformation, "������ʾ"
            Exit Sub

        End If
        
        old = Get_SqlserverCnt(" SELECT * FROM erptemp..bom_substitutes a  WHERE a.materiel_1 = '" & strmateriel & "' ")
        If old > 0 Then

           MsgBox "���" & strmateriel & "�������Ϻ�,��ʹ�����Ϻŵ���!", vbInformation, "��ʾ"
           Exit Sub

         End If
         
           
             
           strSql = "  SELECT a.FNumber,a.FName,isnull(b.materiel_1,''),isnull(b.sub_code,'') ,isnull(c.������λ����,' '), isnull(a.FModel,' '),isnull(a.F_103,' ')   FROM AIS20141114094336.dbo.t_ICItem a  LEFT JOIN erptemp..bom_substitutes b   ON b.materiel_2 = a.F_101  " & _
           "   LEFT JOIN ERPBASE..tblUnitData c ON c.�ṹ���� = a.FProductUnitID WHERE a.F_101 ='" & strmateriel & "'"
       
            If oiRS.State = adStateOpen Then oiRS.Close
            oiRS.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not oiRS.EOF Then
            
                .Row = Row
                .Col = Col + 1
                .Text = Trim(oiRS.Fields(0).Value) '���ϱ���
                
                .Row = Row
                .Col = Col + 2
                .Text = Trim(oiRS.Fields(1).Value) '����
                

                .Row = Row
                .Col = Col + 3
                .Text = Trim(oiRS.Fields(5).Value) '���
            
                .Row = Row
                .Col = Col + 4
                .Text = Trim(oiRS.Fields(6).Value) '�ͺ�
                
                .Row = Row
                .Col = Col + 7
                .Text = Trim(oiRS.Fields(4).Value) '��λ
            
                .Row = Row
                .Col = Col + 8
                .Text = Trim(oiRS.Fields(2).Value) '��������1
        
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

Private Sub CmdBomDel_Click() 'ɾ��

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
            
                    .Col = E_FPS1.E_PT      '�Ϻ�
                    If DelMaterial = "" Then
                        DelMaterial = Trim(.Text)
                    Else
                        DelMaterial = DelMaterial & "," & Trim(.Text)
                    End If
                    DelCnt = DelCnt + 1
                End If

            End If
        Next i
        If MsgBox("��ȷ��Ҫɾ��" & DelMaterial & ",��" & DelCnt & "��������?", vbOKCancel, "��ʾ") = vbCancel Then
            Exit Sub

        End If
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = E_FPS1.E_SEL

            If .Text <> "" Then
                If .Text = 1 Then
            
                    .Col = E_FPS1.E_PT      '�Ϻ�
                    strmateriel = Trim(.Text)
                
                    .Col = E_FPS1.E_Typeid    '���1
                    dzm = Trim$(.Text)
                    
                    If dzm = "" Then
                        sqlTemp = "delete from  [erpdata].[dbo].[TSVtblMRuleData]  where ���Ϲ淶���='" & bomId & "' and �Ϻ�='" & strmateriel & "' and (���1 is  null or ���1='') "
                        sqlTemp_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'ɾ��',*,'" & User & "'  from  erpdata..TSVtblMRuleData where ���Ϲ淶���='" & bomId & "' and �Ϻ�='" & strmateriel & "' and (���1 is  null or ���1='') "
                    Else
                        sqlTemp = "delete from  [erpdata].[dbo].[TSVtblMRuleData]  where ���Ϲ淶���='" & bomId & "' and �Ϻ�='" & strmateriel & "' and ���1 = '" & dzm & "'"
                        sqlTemp_log = "INSERT INTO erpdata..TSVtblBom_Modify_log SELECT getdate(),'ɾ��',*,'" & User & "'  from  erpdata..TSVtblMRuleData where ���Ϲ淶���='" & bomId & "' and �Ϻ�='" & strmateriel & "' and ���1 = '" & dzm & "'"
                    End If
                    AddSql2 (sqlTemp_log)
                    AddSql2 (sqlTemp)
                    AddSql2 (" UPDATE erpdata..TSVtblSetMRule SET ������� = GETDATE(),�������� = '" & User & "'  WHERE ���Ϲ淶��� = '" & bomId & "' ")
        

                End If

            End If

        Next i

    End With
    cmdQu_Click '��ѯ
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
' ��MES�����ڸ�վ��,��ɷ������½��
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








