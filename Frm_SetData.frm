VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_SetData 
   Caption         =   "��Ϣά��"
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
            Caption         =   "��ѯ"
            Key             =   "QUE"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "ADD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�޸�"
            Key             =   "MOD"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ɾ��"
            Key             =   "DEL"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�"
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
         Caption         =   "����"
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
         Caption         =   "���"
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
         Caption         =   "��ѯ����2"
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
         Caption         =   "��ѯ����"
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
         Caption         =   "ά������"
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

        Case "�ϸ�Ӧ�̺ϸ�����"
            lbl(1) = "��ѯ�Ϻ�"
            lbl(2).Visible = False
            txtKey2.Visible = False
            'lbl(3).Visible = False
            'DTPicker1.Visible = False
            
        Case "�ͻ�������������"
            lbl(1) = "��ѯ�ͻ�����"
            
            lbl(2).Visible = False
            txtKey2.Visible = False
            'lbl(3).Visible = False
            'DTPicker1.Visible = False
            
        Case "ERP������Ч�ڸ���"
            lbl(1) = "��ѯ�Ϻ�"
            lbl(2) = "����"
            lbl(2).Visible = True
            txtKey2.Visible = True
        

            
        Case "����ֳ���"
            lbl(1) = "��ѯ�ͻ�����"
            lbl(2) = "����"
            lbl(2).Visible = True
            txtKey2.Visible = True
            cmdCommand1.Visible = True
            
        Case "������ϸ��"
            lbl(1) = "��������"
'           lbl(2) = "�Ϻ�"
'           lbl(2).Visible = True
'           txtKey2.Visible = True
            lbl(3).Visible = False
            DTPicker1.Visible = False
            lab2 = "���"
            Combo1.Clear
            Combo1.AddItem ("��˰ԭ����")
            Combo1.AddItem ("�Ǳ�˰ԭ����")
            Combo1.AddItem ("�㲿��")
            Combo1.AddItem ("��˰�豸")
            Combo1.AddItem ("�Ǳ�˰�豸")
            Combo1.AddItem ("��ʱ������")
            Combo1.AddItem ("��Ʒ����")
        
        Case "������ϸ��"
            lbl(1) = "�ɹ�����"
'            lbl(2) = "�Ϻ�"
'            lbl(2).Visible = True
'            txtKey2.Visible = True
            lbl(3).Visible = False
            DTPicker1.Visible = False
            lab2 = "���"
'            Combo1.Visible = True
            Combo1.Clear
            Combo1.AddItem ("��˰��Ʒ")
            Combo1.AddItem ("�Ǳ�˰��Ʒ")
            Combo1.AddItem ("�ϼ�����")
            Combo1.AddItem ("�㲿������")
            Combo1.AddItem ("�豸����")
            Combo1.AddItem ("��ʱ����������")

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

                Case "�ϸ�Ӧ�̺ϸ�����"
                    ForMod1
        
                Case "�ͻ�������������"
                    ForMod2
                
                Case "ERP������Ч�ڸ���"
                    'MsgBox "��ʱ����"
                    ForMod3

                Case "����ֳ���"
                    ForMod3
                    
                Case "������ϸ��"
                    ForMod5
                
                Case "������ϸ��"
                    ForMod6
                   
            End Select
        
        Case "DEL"
            
            Select Case cmbCombo1.Text

                Case "�ϸ�Ӧ�̺ϸ�����"
                    ForDel1
        
                Case "�ͻ�������������"
                    ForDel2
                    
                Case "ERP������Ч�ڸ���"
                    MsgBox "��ά�����Ͳ�֧��ɾ������", vbInformation, "��ʾ"
                    Exit Sub
                    
                Case "������ϸ��"
                    ForDel5
                
                Case "������ϸ��"
                    ForDel6

            End Select

        Case "EXIT"
            Unload Me

    End Select

End Sub

Private Sub ForQuery()

    If cmbCombo1.Text = "" Then
        MsgBox "��ѡ��ά������", vbInformation, "��ʾ"
        Exit Sub

    End If

    Select Case cmbCombo1.Text

        Case "�ϸ�Ӧ�̺ϸ�����"
            QueType1
        
        Case "�ͻ�������������"
            QueType2
        
        Case "ERP������Ч�ڸ���"
            QueType3
            
        Case "����ֳ���"
            QueType4
        
        Case "������ϸ��"
            QueType5
                
        Case "������ϸ��"
            QueType6
        

    End Select

End Sub

Private Sub QueType1()

    Dim rs     As New ADODB.Recordset

    Dim strMat As String

    Dim strSql As String
    
    strMat = Trim$(txtKey.Text)

    If txtKey.Text = "" Then
        strSql = "select ���,��Ӧ�̱��,��Ӧ������,�Ϻ�, ���ϱ��,��Ч��,����ʱ��,'' as '��' from ERPBASE..tblCG_PassSupplier where ��Ӧ������ <> ''"
    Else
        strSql = "select ���,��Ӧ�̱��,��Ӧ������,�Ϻ�, ���ϱ��,��Ч��,����ʱ��,'' as '��' from ERPBASE..tblCG_PassSupplier where �Ϻ� = '" & strMat & "' and ��Ӧ������ <> ''"

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType1(rs)
    Else
        
        MsgBox "��ѯ�������Ϻ�", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

Private Sub QueType2()

    Dim rs         As New ADODB.Recordset

    Dim strCusCode As String

    Dim strSql     As String

    If txtKey.Text = "" Then
        strSql = "select CUSTOMER as �ͻ�����,WAREHOUSE as ��������, FLAG as �Ƿ���Ч,'' as '��' from erptemp..tbltransfer"
    Else
        strCusCode = Trim$(txtKey.Text)
        strSql = "select CUSTOMER as �ͻ�����,WAREHOUSE as ��������, FLAG as �Ƿ���Ч,'' as '��' from erptemp..tbltransfer where CUSTOMER = '" & strCusCode & "' "

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType2(rs)
    Else
        
        MsgBox "��ѯ�����ÿͻ�����", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

Private Sub QueType3()

    Dim rs         As New ADODB.Recordset

    Dim strCusCode As String

    Dim strSql     As String
    
    If txtKey.Text = "" Then
        MsgBox "���������ϵ��Ϻ�", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    If txtKey2.Text = "" Then
        strSql = "select '' as '��',AA.id,AA.�ֿ���,BB.F_101 as �Ϻ�,BB.FName as ��������,AA.���ϱ��,AA.����, AA.��Ч����,AA.��������, AA.��λ, AA.��ǰ���� from erpbase.dbo.tblStockNum AA INNER JOIN  AIS20141114094336.dbo.t_ICItem BB ON AA.���ϱ�� = BB.FNumber AND   BB.F_101 = '" & UCase(Trim(txtKey.Text)) & "' and AA.��ǰ���� > 0 "
    Else
        strSql = "select '' as '��',AA.id,AA.�ֿ���,BB.F_101 as �Ϻ�,BB.FName as ��������,AA.���ϱ��,AA.����, AA.��Ч����,AA.��������, AA.��λ, AA.��ǰ���� from erpbase.dbo.tblStockNum AA INNER JOIN  AIS20141114094336.dbo.t_ICItem BB ON AA.���ϱ�� = BB.FNumber AND   BB.F_101 = '" & UCase(Trim(txtKey.Text)) & "' and AA.��ǰ���� > 0 and AA.���� = '" & txtKey2.Text & "' "
        
    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType3(rs)
    Else
        
        MsgBox "��ѯ������������Ϣ", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

Private Sub QueType4()

    Dim rs         As New ADODB.Recordset

    Dim strCusCode As String

    Dim strSql     As String
    
    If txtKey.Text = "" Then
        MsgBox "������ͻ�����", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    If txtKey2.Text = "" Then
        strSql = "      select '' as '��',cc.�ͻ�����,bb.��Ӧ�̱��,cc.�ͻ�����,AA.�ֿ���,aa.����,aa.��ǰ���� ,'' AS �������� FROM erpbase..tblStockNum AA  INNER JOIN  tblSupplierData  bb   ON  bb.��Ӧ�̱�� = aa.��Ӧ�̱��  " & "    LEFT JOIN erpdata..tblXCustomer cc ON cc.�ͻ����� = bb.��Ӧ������  WHERE aa.�ֿ��� = '54'  AND  cc.�ͻ����� like '%" & UCase(Trim(txtKey.Text)) & "%'  AND aa.��ǰ���� > 0   "
    Else

        strSql = "      select '' as '��',cc.�ͻ�����,bb.��Ӧ�̱��,cc.�ͻ�����,AA.�ֿ���,aa.����,aa.��ǰ���� ,'' AS �������� FROM erpbase..tblStockNum AA  INNER JOIN  tblSupplierData  bb   ON  bb.��Ӧ�̱�� = aa.��Ӧ�̱��  " & "    LEFT JOIN erpdata..tblXCustomer cc ON cc.�ͻ����� = bb.��Ӧ������  WHERE aa.�ֿ��� = '54'  AND  cc.�ͻ����� like '%" & UCase(Trim(txtKey.Text)) & "%'  AND aa.��ǰ���� > 0  and  AA.���� = '" & UCase(Trim(txtKey2.Text)) & "' "

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType4(rs)
    Else
        
        MsgBox "��ѯ������Ϣ", vbInformation, "��ʾ"
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
        MsgBox "�������������", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If txtKey2.Text = "" Then
        strSql = "select ��������,�Ϻ�,��Ʊ��,��������,����,���,���ص���,Ʒ��,�ֲ����,��λ,�ܼ�,�ֲ��,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,'' as '��' from erptemp.dbo.ksexport where �������� = '" & strInv & "' and flag = '0' "
    Else

        strSql = "select ��������,�Ϻ�,��Ʊ��,��������,����,���,���ص���,Ʒ��,�ֲ����,��λ,�ܼ�,�ֲ��,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,'' as '��' from erptemp.dbo.ksexport where �������� = '" & strInv & "' and �Ϻ� = '" & strInv1 & "' and flag = '0'"

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType5(rs)
    Else
        
        MsgBox "��ѯ�����ó��ڵ�����Ϣ", vbInformation, "��ʾ"
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
        MsgBox "������ɹ�����", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If txtKey2.Text = "" Then
        strSql = "select �ɹ�����,�Ϻ�,���,���񵽻�����,��׼die,�볡����,��Ʊ��,Ʒ��,���,����,�ֲ��,��˰,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,'' as '��' from erptemp.dbo.ksimport where �ɹ����� = '" & strInv & "' and flag = '0' "
    Else

        strSql = "select �ɹ�����,�Ϻ�,���,���񵽻�����,��׼die,�볡����,��Ʊ��,Ʒ��,���,����,�ֲ��,��˰,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,'' as '��' from erptemp.dbo.ksimport where �ɹ����� = '" & strInv & "' and �Ϻ� = '" & strInv1 & "' and flag = '0'"

    End If
    
    fpS(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType6(rs)
    Else
        
        MsgBox "��ѯ�����ó��ڵ�����Ϣ", vbInformation, "��ʾ"
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

    If Toolbar1.Buttons(3).Caption = "�ύ" Then
        
        Select Case cmbCombo1.Text

            Case "�ϸ�Ӧ�̺ϸ�����"
                ForCommit1
        
            Case "�ͻ�������������"
                ForCommit2
             
            Case "������ϸ��"
                ForCommit5
                
            Case "������ϸ��"
                ForCommit6

        End Select
        
        Exit Sub

    End If

    If cmbCombo1.Text = "" Then
        MsgBox "��ѡ��ά������", vbInformation, "��ʾ"
        Exit Sub

    End If

    Select Case cmbCombo1.Text

        Case "�ϸ�Ӧ�̺ϸ�����"
            AddType1
        
        Case "�ͻ�������������"
            AddType2
            
        Case "ERP������Ч�ڸ���"
            MsgBox "��ά�����Ͳ�֧����������", vbInformation, "��ʾ"
            Exit Sub

        Case "������ϸ��"
            AddType5

        Case "������ϸ��"
            AddType6

    End Select

End Sub

Private Sub AddType1()

    Dim rs     As New ADODB.Recordset

    Dim strMat As String, strMatNo As String

    Dim strSql As String

    If txtKey.Text = "" Then
        MsgBox "����дҪά�����Ϻ�", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    fpS(0).MaxRows = 0

    strMat = Trim$(txtKey.Text)
    strMatNo = Get_SqlStr("select ���ϱ�� from dbo.tblSmainM2 where �Ϻ� = '" & strMat & "'")
    
    If strMatNo = "" Then
        MsgBox "��ѯ�������Ϻ���Ϣ,�Ƿ�����", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strSql = "select '' as ��Ӧ�̱��,'' as ��Ӧ������,'" & strMat & "' as �Ϻ�, '" & strMatNo & "' as ���ϱ��"

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    Call ListDataType1(rs)
    
    Toolbar1.Buttons(3).Caption = "�ύ"
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
        MsgBox "����дҪά���Ŀͻ�����", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    fpS(0).MaxRows = 0

    strCusCode = Trim$(txtKey.Text)
 
    strSql = "select '" & strCusCode & "' as �ͻ�����,'' as ��������"

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    Call ListDataType2(rs)
    
    Toolbar1.Buttons(3).Caption = "�ύ"
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
        MsgBox "����дҪά���ĳ�������", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    
    fpS(0).MaxRows = 0

    strInv = Trim$(txtKey.Text)
    

    If Get_SqlserverCnt("SELECT * FROM erpdata..tblStockMove A WHERE A.���ݱ�� = '" & strInv & "'") = 0 Then
        MsgBox "û�д˳�������,����������", vbInformation, "��ʾ"
        Exit Sub

    End If

    strSql = "select b.���ݱ�� as ���ݺ���,c.�Ϻ�,(select distinct ���۷�Ʊ from erptemp.dbo.tblBB_CPFH_Invoice  a where  a.�������� = b.���ݱ��) as ��Ʊ��,CONVERT(varchar(100), b.��������, 23) as ��������,SUM(b.ʵ����Ʒ��+b.ʵ��������+b.ʵ���Ƴ̲�����) as ����,'' as ���,'' as ���ص���,'' as Ʒ��,'' as �ֲ����,'' as ��λ,'' as �ܼ�,'' as �ֲ��,'' as AWB#,'' as Ŀ�ĵ�,'' as ����,'' as �˵�����,'' as ��ע,'' as '��' from   erpdata..tblStockMove b,erpdata..tblSmainM2 c  where    b.���ݱ�� = '" & strInv & "' and  c.���ϱ�� = b.���ϱ�� and c.�Ϻ� not in (select distinct �Ϻ� from erptemp.dbo.ksexport where �������� = '" & strInv & "' and flag = '0') group by b.���ݱ��,c.�Ϻ�,CONVERT(varchar(100), b.��������, 23) "
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    Call ListDataType5(rs)
    
    Toolbar1.Buttons(3).Caption = "�ύ"
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
        MsgBox "����дҪά���Ĳɹ�����", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    ID = 1
    
    fpS(0).MaxRows = 0

    strInv = Trim$(txtKey.Text)

    If Get_SqlserverCnt("SELECT * FROM erpbase..tblCPurDataSub WHERE �ɹ������ = '" & strInv & "'") = 0 Then
        MsgBox "û�д˲ɹ�����,����������", vbInformation, "��ʾ"
        Exit Sub

    End If

    strSql = "SELECT a.�ɹ������,b.�Ϻ�,'' AS ���,ceiling(sum(a.��׼�ɹ�����) - isnull(SUM(c.���񵽻�����),0)) as ���񵽻�����,'' as ��׼die,'' as �볡����,'' as ��Ʊ��,'' as Ʒ��,'' as ���,'' as ����,'' as �ֲ��,'' as ��˰,'' as ��ֵ˰,'' as ���ص���,'' as AWB#,'' as ����,'' as �˵�����,'' as ��ע,'' as id ,'' as '��' FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b  left join erptemp.dbo.ksimport c on c.�Ϻ� = b.�Ϻ� and flag = '0' and c.�ɹ����� = '" & strInv & "' WHERE a.�ɹ������ = '" & strInv & "' and a.���ϱ�� = b.���ϱ�� GROUP by a.�ɹ������,b.�Ϻ� "
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    Call ListDataType6(rs)
    
    Toolbar1.Buttons(3).Caption = "�ύ"
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
            MsgBox "�����빩Ӧ�̱��", vbInformation, "��ʾ"
            Exit Sub

        End If
    
        strGYSNo = Trim$(.Text)
    
        .Col = 2

        If .Text = "" Then
            MsgBox "�����빩Ӧ������", vbInformation, "��ʾ"
            Exit Sub

        End If
    
        strGYSName = Trim$(.Text)
    
        .Col = 3
        strMat = Trim$(.Text)
    
        .Col = 4
        strMatNo = Trim$(.Text)

    End With

    AddSql2 ("insert into ERPBASE..tblCG_PassSupplier( ��Ӧ�̱��,��Ӧ������,�Ϻ�, ���ϱ��,��Ч��,����ʱ��) values('" & strGYSNo & "','" & strGYSName & "','" & strMat & "','" & strMatNo & "','1',GetDate())")

    MsgBox "�����ɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(3).Caption = "����"
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
            MsgBox "�������������", vbInformation, "��ʾ"
            Exit Sub

        End If
    
        strHouse = Trim$(.Text)

        If Get_SqlserverCnt("select * from erpdata..tblStockmovesub where ���ݱ�� = '" & strHouse & "'") = 0 Then
            MsgBox "�鲻���ó�������", vbInformation, "��ʾ"
            Exit Sub

        End If
        
    End With
    
    Dim rs As New ADODB.Recordset
    
    Set rs = Get_SqlserveRs("SELECT b.�ͻ�����,A.���ݱ��,a.�Ϻ�,SUM(a.����) as DIE����,COUNT(DISTINCT a.���̿����)  as Ƭ���� FROM erpdata..tblStockmovesub A,erpdata..tblStockmove B " & " WHERE A.���ݱ�� = '" & strHouse & "' AND b.���ݱ�� = a.���ݱ�� AND b.��� = a.������� GROUP BY  A.���ݱ��, b.�ͻ�����,a.�Ϻ�")

    With fpS(0)
        
        .MaxRows = 0

        If rs.RecordCount > 0 Then
            Set .DataSource = rs

        End If

    End With
    
    If MsgBox("ȷ����Ϣ�Ƿ�����", vbYesNoCancel, "��ʾ") = vbNo Then
        Exit Sub

    End If
    
    AddSql2 ("insert into erptemp..tbltransfer(CUSTOMER,WAREHOUSE,CREATE_DATE,CREATE_BY,LAST_UPDATE_DATE,LAST_UPDATE_BY, FLAG) values('" & strCusCode & "','" & strHouse & "',CONVERT(varchar(100), GETDATE(), 23),'" & gUserName & "','','',1)")

    MsgBox "�����ɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(3).Caption = "����"
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
            MsgBox "û������", vbInformation, "��ʾ"
        
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
                    MsgBox "�������������", vbInformation, "��ʾ"
                    Exit Sub

                End If
    
                strInv1 = Trim$(.Text)
    
                .Col = 2

                If .Text = "" Then
                    MsgBox "�������Ϻ�", vbInformation, "��ʾ"
                    Exit Sub

                End If
    
                strInv2 = Trim$(.Text)
                
                If Get_SqlserverCnt("select * from erptemp.dbo.ksexport where �������� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and flag = '0'") > 0 Then
                    MsgBox "�ñ������Ѿ�������", vbInformation, "��ʾ"
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
                    MsgBox "���������", vbInformation, "��ʾ"
                    Exit Sub

                End If
                
                .Text = Combo1.Text
                
                If .Text = "" Then
                    MsgBox "���������", vbInformation, "��ʾ"
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

                AddSql2 ("insert into erptemp.dbo.ksexport( ��������,�Ϻ�,��Ʊ��,��������,����,���,���ص���,Ʒ��,�ֲ����,��λ,�ܼ�,�ֲ��,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,����ʱ��,�޸�״̬,�޸�ʱ��,ɾ��ʱ��,flag) values('" & strInv1 & "','" & strInv2 & "','" & strInv3 & "','" & strInv4 & "','" & strInv5 & "','" & strInv6 & "','" & strInv7 & "','" & strInv8 & "','" & strInv9 & "','" & strInv10 & "','" & strInv11 & "','" & strInv12 & "','" & strInv13 & "','" & strInv14 & "','" & strInv15 & "','" & strInv16 & "','" & strInv17 & "',GetDate(),NULL,NULL,NULL,'0')")

            End If

            '
            '            If bFlag = False And j = 0 Then
            '                MsgBox "��ѡ��Ҫ��������", vbInformation, "��ʾ"
            '                Exit Sub
            '
            '            End If

        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "��ѡ��Ҫ��������", vbInformation, "��ʾ"
            Exit Sub
            
        End If

    End With
    
    MsgBox "�����ɹ�", vbInformation, "��ʾ"
    Toolbar1.Buttons(3).Caption = "����"
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
            MsgBox "û������", vbInformation, "��ʾ"
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
                    MsgBox "������ɹ�����", vbInformation, "��ʾ"
                    Exit Sub

                End If
    
                strInv1 = Trim$(.Text)
    
                .Col = 2

                If .Text = "" Then
                    MsgBox "�������Ϻ�", vbInformation, "��ʾ"
                    Exit Sub

                End If
    
                strInv2 = Trim$(.Text)
    
                .Col = 3
                
                lab2.Visible = True
                Combo1.Visible = True

                If Combo1.Text = "" Then
                    MsgBox "���������", vbInformation, "��ʾ"
                    Exit Sub

                End If
                
                .Text = Combo1.Text
                
                If .Text = "" Then
                    MsgBox "���������", vbInformation, "��ʾ"
                    Exit Sub

                End If
                
                strInv3 = Trim$(.Text)
                              
                .Col = 4
                
                strInv4 = Trim$(.Text)
                
                strNo1 = Get_SqlStr("SELECT ceiling(isnull(SUM(a.��׼�ɹ�����),0)) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.�ɹ������ = '" & strInv1 & "' and a.���ϱ�� = b.���ϱ�� and b.�Ϻ� = '" & strInv2 & "' ")
                
                strNo2 = Get_SqlStr("SELECT ceiling(isnull(SUM(���񵽻�����),0)) FROM erptemp.dbo.ksimport where �ɹ����� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and flag = '0'")
                
                strNo3 = strNo1 - strNo2
                
                If strInv4 > strNo3 Then
                    MsgBox "�ñ��Ϻ�" & strInv2 & "��׼�ɹ�����: " & strNo1 & ",�Ѿ�ά������������" & strNo2 & ",�������ֻ��ά����" & strNo3 & "", vbInformation, "��ʾ"
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
                
                If Get_SqlserverCnt("select * from erptemp.dbo.ksimport where �ɹ����� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and flag = '0'") > 0 Then
                
                    strid1 = Get_SqlStr(" select MAX(id) from erptemp.dbo.ksimport where �ɹ����� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and flag = '0'")
                    
                    strID = strid1 + 1
              
                End If
        
                .Text = strID
                
                strInv19 = Trim$(.Text)

                AddSql2 ("insert into erptemp.dbo.ksimport( �ɹ�����,�Ϻ�,���,���񵽻�����,��׼die,�볡����,��Ʊ��,Ʒ��,���,����,�ֲ��,��˰,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,����ʱ��,�޸�״̬,�޸�ʱ��,ɾ��ʱ��,flag) values('" & strInv1 & "','" & strInv2 & "','" & strInv3 & "','" & strInv4 & "','" & strInv5 & "','" & strInv6 & "','" & strInv7 & "','" & strInv8 & "','" & strInv9 & "','" & strInv10 & "','" & strInv11 & "','" & strInv12 & "','" & strInv13 & "','" & strInv14 & "','" & strInv15 & "','" & strInv16 & "','" & strInv17 & "','" & strInv18 & "','" & strInv19 & "',GetDate(),NULL,NULL,NULL,'0')")

            End If

            '           strid = strid + 1
        Next
        
        'j = 0 ��ȡ�����û���Ҫ���������
        If bFlag = False And j = 0 Then
            MsgBox "��ѡ��Ҫ��������", vbInformation, "��ʾ"
            Exit Sub
            
        End If
        
        lab2.Visible = Flase
        Combo1.Visible = Flase
        '        Combol.Text = ""

    End With
    
    MsgBox "�����ɹ�", vbInformation, "��ʾ"
    Toolbar1.Buttons(3).Caption = "����"
    Toolbar1.Buttons(3).Image = 2
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery

End Sub


Private Sub ForMod1()

    Dim i As Integer

    If Toolbar1.Buttons(5).Caption <> "�ύ" Then

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
    
        Toolbar1.Buttons(5).Caption = "�ύ"
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
            MsgBox "û������", vbInformation, "��ʾ"
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
        MsgBox "��ѡ��Ҫ�޸ĵ���", vbInformation, "��ʾ"
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
                
                AddSql2 ("update ERPBASE..tblCG_PassSupplier set ��Ӧ�̱�� = '" & strGYSNo & "', ��Ӧ������ = '" & strGYSName & "' where ��� = '" & strno & "'     ")
            
            End If
            
        Next
    
    End With
    
    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(5).Caption = "�޸�"
    Toolbar1.Buttons(5).Image = 3
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery
    
End Sub

Private Sub ForMod2()

    Dim i As Integer

    If Toolbar1.Buttons(5).Caption <> "�ύ" Then

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
    
        Toolbar1.Buttons(5).Caption = "�ύ"
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
            MsgBox "û������", vbInformation, "��ʾ"
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
        MsgBox "��ѡ��Ҫ�޸ĵ���", vbInformation, "��ʾ"
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
    
    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(5).Caption = "�޸�"
    Toolbar1.Buttons(5).Image = 3
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery
    
End Sub

Private Sub ForMod3()

    Dim i As Integer

    If Toolbar1.Buttons(5).Caption <> "�ύ" Then

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
    
        Toolbar1.Buttons(5).Caption = "�ύ"
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
            MsgBox "û������", vbInformation, "��ʾ"
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
        MsgBox "��ѡ��Ҫ�޸ĵ���", vbInformation, "��ʾ"
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
                
                AddSql2 ("update erpbase.dbo.tblStockNum  set ��Ч���� = '" & strNewDate & "', ��λ = '" & strKW & "'  where id = '" & strID & "' ")
            
            End If
            
        Next
    
    End With
    
    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(5).Caption = "�޸�"
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

    If Toolbar1.Buttons(5).Caption <> "�ύ" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
    
                For m = 7 To 18
            
                    .Col = m
                    .Lock = False
      
                Next
    
            Next
        
        End With
    
        Toolbar1.Buttons(5).Caption = "�ύ"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "û������", vbInformation, "��ʾ"
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
                
                
                AddSql2 ("insert into erptemp.dbo.ksexport (��������,�Ϻ�,��Ʊ��,��������,����,���,���ص���,Ʒ��,�ֲ����,��λ,�ܼ�,�ֲ��,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,����ʱ��,�޸�״̬,�޸�ʱ��,ɾ��ʱ��,flag) SELECT ��������,�Ϻ�,��Ʊ��,��������,����,���,���ص���,Ʒ��,�ֲ����,��λ,�ܼ�,�ֲ��,AWB#,Ŀ�ĵ�,����,�˵�����,��ע,����ʱ��,'�޸�ǰ',�޸�ʱ��,ɾ��ʱ��,'2' FROM erptemp.dbo.ksexport WHERE �������� = '" & strInv1 & "'  AND �Ϻ� =  '" & strInv2 & "' AND  flag = '0'")
                AddSql2 ("update erptemp.dbo.ksexport set ���ص��� =  '" & strInv7 & "',Ʒ�� =  '" & strInv8 & "',�ֲ���� =  '" & strInv9 & "',��λ =  '" & strInv10 & "',�ܼ� =  '" & strInv11 & "',�ֲ�� =  '" & strInv12 & "',AWB# =  '" & strInv13 & "',Ŀ�ĵ� =  '" & strInv14 & "',���� =  '" & strInv15 & "',�˵����� =  '" & strInv16 & "',��ע =  '" & strInv17 & "',�޸�״̬ = '�޸ĺ�',�޸�ʱ�� = '" & strtime & "' where �������� = '" & strInv1 & "' and flag = '0' and �Ϻ�  = '" & strInv2 & "' ")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "��ѡ��Ҫ�޸ĵ���", vbInformation, "��ʾ"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(5).Caption = "�޸�"
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
'    If Toolbar1.Buttons(5).Caption <> "�ύ" Then
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
'        Toolbar1.Buttons(5).Caption = "�ύ"
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
'            MsgBox "û������", vbInformation, "��ʾ"
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
'                strNo1 = Get_SqlStr("SELECT ceiling(isnull(SUM(a.��׼�ɹ�����),0)) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.�ɹ������ = '" & strInv1 & "' and a.���ϱ�� = b.���ϱ�� and b.�Ϻ� = '" & strInv2 & "' ")
'
'                strNo2 = Get_SqlStr("SELECT ceiling(isnull(SUM(���񵽻�����),0)) FROM erptemp.dbo.ksimport where �ɹ����� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and flag = '0'")
'
'                strNo3 = strNo1 - strNo2
'
'                If strInv4 > strNo3 Then
'                    MsgBox "�ñ��Ϻ�" & strInv2 & "��׼�ɹ�����: " & strNo1 & ",�Ѿ�ά������������" & strNo2 & ",�������ֻ��ά����" & strNo3 & "", vbInformation, "��ʾ"
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
'                AddSql2 ("insert into erptemp.dbo.ksimport(�ɹ�����,�Ϻ�,���,���񵽻�����,��׼die,�볡����,��Ʊ��,Ʒ��,���,����,�ֲ��,��˰,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,����ʱ��,�޸�״̬,�޸�ʱ��,ɾ��ʱ��,flag) SELECT �ɹ�����,�Ϻ�,���,���񵽻�����,��׼die,�볡����,��Ʊ��,Ʒ��,���,����,�ֲ��,��˰,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,����ʱ��,'�޸�ǰ',�޸�ʱ��,ɾ��ʱ��,'2' FROM erptemp.dbo.ksimport WHERE �ɹ����� = '" & strInv1 & "'  AND �Ϻ� =  '" & strInv2 & "' AND id =  '" & strInv19 & "'  AND  flag = '0'")
'
'                AddSql2 ("update erptemp.dbo.ksimport set ���񵽻����� = '" & strInv4 & "',��׼die =  '" & strInv5 & "',�볡���� =  '" & strInv6 & "',��Ʊ�� =  '" & strInv7 & "',Ʒ�� =  '" & strInv8 & "',��� =  '" & strInv9 & "',���� =  '" & strInv10 & "',�ֲ�� =  '" & strInv11 & "',��˰ =  '" & strInv12 & "',��ֵ˰ =  '" & strInv13 & "',���ص��� =  '" & strInv14 & "',AWB#  =  '" & strInv15 & "',���� =  '" & strInv16 & "', �˵����� =  '" & strInv17 & "',��ע =  '" & strInv18 & "',�޸�״̬ = '�޸ĺ�',�޸�ʱ�� = '" & strTime & "' where �ɹ����� = '" & strInv1 & "' and flag = '0' and �Ϻ�  = '" & strInv2 & "' and id =  '" & strInv19 & "' ")
'
'            End If
'
'        Next
'
'        If bFlag = False And j = 0 Then
'            MsgBox "��ѡ��Ҫ�޸ĵ���", vbInformation, "��ʾ"
'            Exit Sub
'
'        End If
'
'    End With
'
'    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"
'
'    Toolbar1.Buttons(5).Caption = "�޸�"
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

    If Toolbar1.Buttons(5).Caption <> "�ύ" Then

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
    
        Toolbar1.Buttons(5).Caption = "�ύ"
        Toolbar1.Buttons(5).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Exit Sub

    End If

    bFlag = False
    
    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "û������", vbInformation, "��ʾ"
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
                
                    
                strNo1 = Get_SqlStr("SELECT ceiling(isnull(SUM(a.��׼�ɹ�����),0)) FROM erpbase..tblCPurDataSub a,erpdata..tblSmainM2 b WHERE a.�ɹ������ = '" & strInv1 & "' and a.���ϱ�� = b.���ϱ�� and b.�Ϻ� = '" & strInv2 & "' ")
                
                strNo2 = Get_SqlStr("SELECT ceiling(isnull(SUM(���񵽻�����),0)) FROM erptemp.dbo.ksimport where �ɹ����� = '" & strInv1 & "' and �Ϻ� = '" & strInv2 & "' and id <> '" & strInv19 & "' and flag = '0'")
                
                strNo3 = strNo1 - strNo2
                
                If strInv4 > strNo3 Then
                    MsgBox "�ñ��Ϻ�" & strInv2 & "��׼�ɹ�����: " & strNo1 & ",�Ѿ�ά������������" & strNo2 & ",�������ֻ��ά����" & strNo3 & "", vbInformation, "��ʾ"
                    Exit Sub

                End If
                
                strtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
                
                AddSql2 ("insert into erptemp.dbo.ksimport(�ɹ�����,�Ϻ�,���,���񵽻�����,��׼die,�볡����,��Ʊ��,Ʒ��,���,����,�ֲ��,��˰,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,����ʱ��,�޸�״̬,�޸�ʱ��,ɾ��ʱ��,flag) SELECT �ɹ�����,�Ϻ�,���,���񵽻�����,��׼die,�볡����,��Ʊ��,Ʒ��,���,����,�ֲ��,��˰,��ֵ˰,���ص���,AWB#,����,�˵�����,��ע,id,����ʱ��,'�޸�ǰ',�޸�ʱ��,ɾ��ʱ��,'2' FROM erptemp.dbo.ksimport WHERE �ɹ����� = '" & strInv1 & "'  AND �Ϻ� =  '" & strInv2 & "' AND id =  '" & strInv19 & "'  AND  flag = '0'")
                
                AddSql2 ("update erptemp.dbo.ksimport set ���񵽻����� = '" & strInv4 & "',��׼die =  '" & strInv5 & "',�볡���� =  '" & strInv6 & "',��Ʊ�� =  '" & strInv7 & "',Ʒ�� =  '" & strInv8 & "',��� =  '" & strInv9 & "',���� =  '" & strInv10 & "',�ֲ�� =  '" & strInv11 & "',��˰ =  '" & strInv12 & "',��ֵ˰ =  '" & strInv13 & "',���ص��� =  '" & strInv14 & "',AWB#  =  '" & strInv15 & "',���� =  '" & strInv16 & "', �˵����� =  '" & strInv17 & "',��ע =  '" & strInv18 & "',�޸�״̬ = '�޸ĺ�',�޸�ʱ�� = '" & strtime & "' where �ɹ����� = '" & strInv1 & "' and flag = '0' and �Ϻ�  = '" & strInv2 & "' and id =  '" & strInv19 & "' ")
            
            End If
            
        Next
        
        If bFlag = False And j = 0 Then
            MsgBox "��ѡ��Ҫ�޸ĵ���", vbInformation, "��ʾ"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(5).Caption = "�޸�"
    Toolbar1.Buttons(5).Image = 3
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True

    ForQuery
    
End Sub

Private Sub ForDel1()

    Dim i As Integer

    If Toolbar1.Buttons(7).Caption <> "�ύ" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
            
                .Col = 8
                .Lock = False
            Next
        
        End With
    
        Toolbar1.Buttons(7).Caption = "�ύ"
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
            MsgBox "û������", vbInformation, "��ʾ"
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
        MsgBox "��ѡ��Ҫɾ������", vbInformation, "��ʾ"
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
                
                AddSql2 ("delete from ERPBASE..tblCG_PassSupplier where ��� = '" & strno & "'  ")
            
            End If
            
        Next
    
    End With
    
    MsgBox "ɾ���ɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(7).Caption = "ɾ��"
    Toolbar1.Buttons(7).Image = 5
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(5).Enabled = True

    ForQuery
    
End Sub

Private Sub ForDel2()

    Dim i As Integer

    If Toolbar1.Buttons(7).Caption <> "�ύ" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
            
                .Col = 4
                .Lock = False
            Next
        
        End With
    
        Toolbar1.Buttons(7).Caption = "�ύ"
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
            MsgBox "û������", vbInformation, "��ʾ"
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
        MsgBox "��ѡ��Ҫɾ������", vbInformation, "��ʾ"
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
    
    MsgBox "ɾ���ɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(7).Caption = "ɾ��"
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
            MsgBox "û������", vbInformation, "��ʾ"
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
                    MsgBox "����д��������", vbInformation, "��ʾ"
                    Exit Sub
                    
                End If
           
            End If

        Next

    End With

    If bFlag = False Then
        MsgBox "��ѡ��Ҫ���˵���", vbInformation, "��ʾ"
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
                    
                    AddSql2 ("UPDATE erpbase..tblStockNum SET ��ǰ���� =  " & intQtyN & " -  " & intQtyU & "  WHERE �ֿ��� = '54'  AND ��Ӧ�̱�� = '" & strSup & " '  AND  ���� = '" & strPt & "'")
                Else
                    MsgBox "�����������ڿ������", vbInformation, "��ʾ"
                    Exit Sub

                End If

            End If
            
        Next
    
    End With
    
    MsgBox "���˳ɹ�", vbInformation, "��ʾ"

    ForQuery
 
End Sub

Private Sub ForDel5()

    Dim i       As Integer

    Dim j       As Integer
    
    Dim bFlag   As Boolean
    
    Dim strInv1 As String

    Dim strInv2 As String

    Dim strtime As String

    If Toolbar1.Buttons(7).Caption <> "�ύ" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                
                .Col = 18
                .Lock = False
              
            Next
        
        End With
    
        Toolbar1.Buttons(7).Caption = "�ύ"
        Toolbar1.Buttons(7).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Exit Sub

    End If

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "û������", vbInformation, "��ʾ"
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
                AddSql2 ("update erptemp.dbo.ksexport set flag = '1',ɾ��ʱ��  = '" & strtime & "' where �������� = '" & strInv1 & "' and  �Ϻ� = '" & strInv2 & "' and flag = '0'")

            End If

        Next

    End With

    If bFlag = False And j = 0 Then
        MsgBox "��ѡ��Ҫɾ������", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    MsgBox "ɾ���ɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(7).Caption = "ɾ��"
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

    If Toolbar1.Buttons(7).Caption <> "�ύ" Then

        With fpS(0)

            For i = 1 To .MaxRows
                .Row = i
                
                .Col = 20
                .Lock = False
              
            Next
        
        End With
    
        Toolbar1.Buttons(7).Caption = "�ύ"
        Toolbar1.Buttons(7).Image = 6
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Exit Sub

    End If

    bFlag = False

    With fpS(0)

        If .MaxRows = 0 Then
            MsgBox "û������", vbInformation, "��ʾ"
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
                AddSql2 ("update erptemp.dbo.ksimport set flag = '1',ɾ��ʱ��  = '" & strtime & "' where �ɹ����� = '" & strInv1 & "' and  �Ϻ� = '" & strInv2 & "' and id = '" & strInv19 & "' and flag = '0'")

            End If

        Next

    End With

    If bFlag = False And j = 0 Then
        MsgBox "��ѡ��Ҫɾ������", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    MsgBox "ɾ���ɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(7).Caption = "ɾ��"
    Toolbar1.Buttons(7).Image = 5
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(5).Enabled = True

    ForQuery
    
End Sub
















