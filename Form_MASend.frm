VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form_MASend 
   Caption         =   "MA_Send"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form_MASend"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra 
      Caption         =   "��������"
      Height          =   12855
      Left            =   -720
      TabIndex        =   0
      Top             =   120
      Width           =   22695
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF00FF&
         Caption         =   "�޸�����"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkall 
         Caption         =   "ȫѡ/ȫ��ѡ"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdmap 
         Caption         =   "ƥ��"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6000
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtShop_Order 
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtCust 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H0000FFFF&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdquery 
         Caption         =   "��ѯ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4320
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   5000
         Index           =   0
         Left            =   2400
         TabIndex        =   8
         Top             =   2200
         Width           =   9375
         _Version        =   524288
         _ExtentX        =   16536
         _ExtentY        =   8819
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
         MaxCols         =   6
         MaxRows         =   0
         SpreadDesigner  =   "Form_MASend.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   5000
         Index           =   1
         Left            =   11760
         TabIndex        =   9
         Top             =   2200
         Width           =   8000
         _Version        =   524288
         _ExtentX        =   14111
         _ExtentY        =   8819
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
         MaxCols         =   9
         MaxRows         =   0
         SpreadDesigner  =   "Form_MASend.frx":04E2
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   5055
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   7200
         Width           =   11655
         _Version        =   524288
         _ExtentX        =   20558
         _ExtentY        =   8916
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
         SpreadDesigner  =   "Form_MASend.frx":09C4
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   5000
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   2200
         Width           =   2295
         _Version        =   524288
         _ExtentX        =   4048
         _ExtentY        =   8819
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
         MaxCols         =   2
         MaxRows         =   0
         SpreadDesigner  =   "Form_MASend.frx":0EA6
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   5175
         Index           =   4
         Left            =   11760
         TabIndex        =   15
         Top             =   7200
         Width           =   8000
         _Version        =   524288
         _ExtentX        =   14111
         _ExtentY        =   9128
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
         SpreadDesigner  =   "Form_MASend.frx":136A
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin MSForms.ComboBox cbWarehouse 
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   2655
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4683;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lbl03 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label lbl02 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   915
      End
      Begin VB.Label lbl01 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ֿ���:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   915
      End
   End
End
Attribute VB_Name = "Form_MASend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strRealName As String
Dim order_success  As String
Dim order_fail  As String


Private Enum FpsH 'fps(0)

    e_ID
    E_CHOOSE
    e_order
  '  e_num
    e_matpn
    e_matname
    E_QTY
    e_wh
    e_shop_order
    e_ma
    e_MCol1

End Enum



Private Enum FpsD 'fps(1)

    e_ID
    e_order
    'e_index
    e_matpn
    e_matname
    e_mastock
    E_LOT
    E_stock
    e_stock_k
    e_stock_s
    e_date
    e_duedate
    e_manum
    E_maid
    e_MCol2

End Enum


Private Enum FpsO 'fps(3)


    e_ID
    e_choice
    e_ordernum
    e_MCol3

End Enum


Private Enum FpsM 'fps(4)

    e_ID
    E_CHOOSE
    e_order
    e_Item
    e_matpn
    e_matname
    E_QTY_O
    E_QTY_N
    e_shop_order
    e_ma
    e_MCol1

End Enum




Private Sub ChkAll_Click()
    Dim i As Integer
    
    If chkall.Value = 1 Then

        For i = 1 To Fps(3).MaxRows

            With Fps(3)
                .Row = i
                .Col = FpsO.e_choice
            '    .Text = 1
                Call fps_Click(3, FpsO.e_choice, i)

            End With

        Next i
        
    ElseIf chkall.Value = 0 Then

        For i = 1 To Fps(3).MaxRows

            With Fps(3)
                .Row = i
                .Col = FpsO.e_choice
               ' .Text = 0
                Call fps_Click(3, FpsO.e_choice, i)

            End With

        Next i
        
    End If
End Sub

Private Sub cmdmap_Click()
Dim strorder As String
Dim strmatno As String
Dim strwh As String
Dim strqty As String
Dim i As Integer

AddSql2 ("delete from erpbase..Posting_temp ")
With Fps(1)
   .MaxRows = 0
End With

With Fps(0)
    For i = 1 To .MaxRows
        .Row = i
        .Col = FpsH.e_order
         strorder = Trim$(.text)
        .Col = FpsH.e_ma
         strmatno = Trim$(.text)
        .Col = FpsH.e_wh
        strwh = Trim$(.text)
        .Col = FpsH.E_QTY
        strqty = Trim$(.text)
        AddSql2 ("delete from erpbase..Posting_temp where �ֿ���='" & strwh & "' and ���ϱ��='" & strmatno & "' and ��ǰ����>0 ")
        AddSql2 ("insert into  erpbase..Posting_temp select �ֿ���,���ϱ��,����,��ǰ����,��������,ID,��Ч���� ,row_number() OVER (order BY ��Ч����,ID) from erpbase..tblstocknum where �ֿ���='" & strwh & "' and ���ϱ��='" & strmatno & "' and ��ǰ����>0 ")
        AddSql2 ("delete from erpbase..Posting_temp_bak where �ֿ���='" & strwh & "' and ���ϱ��='" & strmatno & "' and ��ǰ����>0 ")
        AddSql2 ("insert into  erpbase..Posting_temp_bak select �ֿ���,���ϱ��,����,��ǰ����,��������,ID,��Ч����  ,row_number() OVER (order BY ��Ч����,ID)  from erpbase..tblstocknum where �ֿ���='" & strwh & "' and ���ϱ��='" & strmatno & "' and ��ǰ����>0 ")
       
    Next
    
    For i = 1 To .MaxRows
        .Row = i
        .Col = FpsH.e_order
         strorder = Trim$(.text)
        .Col = FpsH.e_ma
         strmatno = Trim$(.text)
        .Col = FpsH.e_wh
        strwh = Trim$(.text)
        .Col = FpsH.E_QTY
        strqty = Trim$(.text)
         Call Searchmat_Bymatno(strorder, strmatno, strwh, strqty, 1)
    Next
End With

End Sub


Private Sub CmdSend_Click()
Dim MsgRly As String
Dim i As Integer
Dim strorder As String
order_success = "" '���˳ɹ��ĵ��ݱ��
MsgRly = MsgBox("��ע��˶Թؼ����ݣ�ȷ��Ҫ���е�ǰ���������", vbYesNo + vbInformation, "��ʾ")
If MsgRly = vbNo Then
    Exit Sub
End If


With Fps(1)
If .MaxRows = 0 Then
    MsgBox "û��Ҫ���˵����ϣ���ȷ��", vbInformation, "��ʾ"
    Exit Sub
End If
End With

With Fps(3)
If .MaxRows = 0 Then
    MsgBox "��ѡ��Ҫ���˵ĵ��ݱ�ţ���ȷ��", vbInformation, "��ʾ"
    Exit Sub
End If

cmdmap_Click

For i = 1 To .MaxRows
    .Row = i
    .Col = FpsH.E_CHOOSE
    If Trim(.text) = "1" Then
        .Col = FpsH.e_order
        strorder = Trim(.text)
        If Posting(strorder) Then
        
        Else
          '  MsgBox strorder & "δ�ɹ����ˣ�������", vbInformation, "��ʾ"
            order_fail = order_fail & "," & strorder
        End If
    End If
Next
End With
If order_fail <> "" Then
    MsgBox "���µ��ݹ���ʧ��" & order_fail & "", vbInformation, "��ʾ"
Else
    MsgBox "������ɣ�", vbInformation, "��ʾ"
    
End If
cmdquery_Click


End Sub

Private Function Posting(order As String)
    Dim adoprm1 As ADODB.Parameter
    Dim adoprm2 As ADODB.Parameter
    Dim adoPrm3 As ADODB.Parameter
    Dim adoPrm4 As ADODB.Parameter
    Dim adoPrm5 As ADODB.Parameter
    Dim adoPrm6 As ADODB.Parameter
    Dim adoPrm7 As ADODB.Parameter
    Dim adoPrm8 As ADODB.Parameter
    Dim adoPrm9 As ADODB.Parameter
    Dim adoprm10 As ADODB.Parameter
    Dim adoPrm11 As ADODB.Parameter
    Dim adoPrm12 As ADODB.Parameter
    Dim adoPrmReturn As ADODB.Parameter
    
    Dim strorder As String
    Dim strIndex As String
    Dim strPickingman As String
    Dim strdepartment As String
    Dim strAuditor As String
    Dim strNote As String
    Dim strPurpose As String
    Dim strwh As String
    Dim strgx As String
    Dim strqty_request As String
    Dim strid As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim intCount As Integer
    Dim strid_list As String
    Dim strqty_list As String

    

    strorder = ""
    strPickingman = ""
    strdepartment = ""
    strAuditor = ""
    strNote = ""
    strPurpose = ""
    strwh = ""
    strgx = ""
    
    Posting = False
    Set rs = Get_SqlserveRs("select * from ERPBASE..tblStockSQ2 where ���ݱ�� ='" & order & "' and ���=1 ")
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        strdepartment = Trim(rs("���ϲ���"))
        strPickingman = Trim(rs("����Ա"))
        strAuditor = Trim(rs("����"))
        strNote = Trim(rs("��ע"))
        strPurpose = Trim(rs("��;"))
        strwh = Trim(rs("�ֿ���"))
        strgx = Trim(rs("����"))
    Else
        MsgBox order & "���ݱ������", vbInformation, "��ʾ"
        Exit Function
    End If
    
    intCount = 0
    With Fps(1)
        For i = 1 To .MaxRows
            strorder = ""
            strqty_request = ""
            strid = ""

            .Row = i
            .Col = FpsD.e_order
            strorder = Trim(.text) '���ݱ��
            
           ' .Col = FpsD.e_index
          '  strIndex = Trim(.Text) '���
            
            .Col = FpsD.e_stock_k '�ۼ����
            strqty_request = Trim(.text)
            
            .Col = FpsD.E_maid 'ID
            strid = Trim(.text)
            
            If strorder = order And Val(strqty_request) > 0 Then
                 strid_list = strid & "��" & strid_list
                 strqty_list = strqty_request & "��" & strqty_list
                 intCount = intCount + 1
            End If
            
        Next
        
    
    End With
   
   Set adoCmd = New ADODB.Command
   Set adoCmd.ActiveConnection = INIadoCon
   adoCmd.CommandText = "uspST_LL2 "
   adoCmd.Parameters.Refresh
   adoCmd.CommandType = adCmdStoredProc
  
   Set adoprm1 = New ADODB.Parameter               'CODE
   adoprm1.type = adChar
   adoprm1.Size = 50
   adoprm1.Direction = adParamInput
   adoprm1.Value = order
   adoCmd.Parameters.Append adoprm1
  
   Set adoprm2 = New ADODB.Parameter              'STOCK OS
   adoprm2.type = adChar
   adoprm2.Size = 50
   adoprm2.Direction = adParamInput
   adoprm2.Value = strRealName
   adoCmd.Parameters.Append adoprm2
  

   Set adoPrm3 = New ADODB.Parameter                 'ʵ������
   adoPrm3.type = adChar
   adoPrm3.Size = 2000
   adoPrm3.Direction = adParamInput
   adoPrm3.Value = strqty_list
   adoCmd.Parameters.Append adoPrm3


   Set adoPrm4 = New ADODB.Parameter                 'ID
   adoPrm4.type = adChar
   adoPrm4.Size = 2000
   adoPrm4.Direction = adParamInput
   adoPrm4.Value = strid_list
   adoCmd.Parameters.Append adoPrm4
   
   Set adoPrm5 = New ADODB.Parameter            '��ע   '
   adoPrm5.type = adChar
   adoPrm5.Size = 255
   adoPrm5.Direction = adParamInput
   adoPrm5.Value = strNote
   adoCmd.Parameters.Append adoPrm5
  
   Set adoPrm6 = New ADODB.Parameter               '���
   adoPrm6.type = adInteger
   adoPrm6.Direction = adParamInput
   adoPrm6.Value = intCount
   adoCmd.Parameters.Append adoPrm6
   
   Set adoPrm7 = New ADODB.Parameter               '��������
   adoPrm7.type = adDBTimeStamp
   adoPrm7.Size = 10
   adoPrm7.Direction = adParamInput
   adoPrm7.Value = Format(Now, "yy-mm-dd")
   adoCmd.Parameters.Append adoPrm7
   
    Set adoPrm8 = New ADODB.Parameter               '����
   adoPrm8.type = adChar
   adoPrm8.Size = 50
   adoPrm8.Direction = adParamInput
   adoPrm8.Value = strgx
   adoCmd.Parameters.Append adoPrm8
   
   Set adoPrm9 = New ADODB.Parameter               '��;
   adoPrm9.type = adChar
   adoPrm9.Size = 255
   adoPrm9.Direction = adParamInput
   adoPrm9.Value = strPurpose
   adoCmd.Parameters.Append adoPrm9
   
   Set adoprm10 = New ADODB.Parameter               '���ϲ���
   adoprm10.type = adChar
   adoprm10.Size = 50
   adoprm10.Direction = adParamInput
   adoprm10.Value = strdepartment
   adoCmd.Parameters.Append adoprm10
   
   Set adoPrm12 = New ADODB.Parameter               '�����
   adoPrm12.type = adChar
   adoPrm12.Size = 50
   adoPrm12.Direction = adParamInput
   adoPrm12.Value = strAuditor
   adoCmd.Parameters.Append adoPrm12

   Set adoPrm11 = New ADODB.Parameter               '����Ա
   adoPrm11.type = adChar
   adoPrm11.Size = 50
   adoPrm11.Direction = adParamInput
   adoPrm11.Value = strPickingman
   adoCmd.Parameters.Append adoPrm11
   
   Set adoPrmReturn = New ADODB.Parameter
   adoPrmReturn.type = adChar
   adoPrmReturn.Size = 50
   adoPrmReturn.Direction = adParamOutput
   adoPrmReturn.Value = adParamReturnValue
   adoCmd.Parameters.Append adoPrmReturn
   adoCmd.Execute

   If Len(Trim(adoPrmReturn.Value)) < 4 Then
    ' MsgBox "���ݲ���ִ�гɹ�����", vbExclamation, Me.Caption
     Posting = True
     order_success = order_success & "," & order
   Else
     'MsgBox "���ݲ���ִ�в��ɹ�:" & adoPrmReturn.Value, vbInformation, Me.Caption
     Exit Function
   End If
   
PROC_EXIT:
 Exit Function
PROC_ERR:

 MsgBox Err.number & vbCrLf & Err.DESCRIPTION, vbInformation, Me.Caption
 On Error GoTo PROC_EXIT

End Function

Private Sub Command1_Click()
Dim i As Integer
Dim strorder As String
Dim stritem As String
Dim strshop_order As String
Dim strqty_O As String
Dim strqty_N As String
Dim Strma As String



With Fps(4)

    For i = 1 To .MaxRows
        .Row = i
        .Col = FpsM.E_CHOOSE
        If .text = "1" Then
            .Col = FpsM.E_QTY_N
            strqty_N = Trim(.text)
            If strqty_N = "" Then
                MsgBox "�޸ĺ���������Ϊ��", vbInformation, "��ʾ"
                Exit Sub
            End If
            If IsNumeric(strqty_N) = False Then
                MsgBox "�޸ĺ���������Ϊ����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    Next
    
    For i = 1 To .MaxRows
        .Row = i
        .Col = FpsM.E_CHOOSE
        If .text = "1" Then
            .Col = FpsM.e_order
            strorder = Trim(.text)
            .Col = FpsM.e_Item
            stritem = Trim(.text)
            .Col = FpsM.E_QTY_O
            strqty_O = Trim(.text)
            .Col = FpsM.E_QTY_N
            strqty_N = Trim(.text)
            .Col = FpsM.e_shop_order
            strshop_order = Trim(.text)
            .Col = FpsM.e_ma
            Strma = Trim(.text)
            AddSql2 ("update ERPBASE..tblStockSQ2 set ���� =" & strqty_N & ",����� =" & strqty_N & " where ���ݱ��='" & strorder & "' and  ���='" & stritem & "'")
            AddSql2 ("insert into  ERPBASE..tblStockSQ2_mod(���ݱ��,��� ,���ϱ��,�޸�ǰ����,�޸ĺ�����,�޸���Ա,�޸�ʱ��) values('" & strorder & "','" & stritem & "','" & Strma & "','" & strqty_O & "','" & strqty_N & "','" & gUserName & "',getdate())")
            
        End If
    Next
    
End With
With Fps(0)
    .MaxRows = 0
End With
With Fps(4)
    .MaxRows = 0
End With
With Fps(3)

    For i = 1 To .MaxRows
        .Row = i
        .Col = FpsM.E_CHOOSE
        If .text = "1" Then
           .Value = 0
           Call fps_Click(3, 1, i)
        End If
        
    Next
    
End With




End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim strSql As String
Dim gUserName1 As String

If Left(gUserName, 1) = "0" Then
   gUserName1 = Right(gUserName, Len(gUserName) - 1)
Else
   gUserName1 = gUserName
End If
strRealName = gUserName & " " & Get_SqlStr2("select rtrim(EmpName) as strRealName from XTW..employee where empno='" & gUserName1 & "'")
If gUserName = "07885" Then
    strRealName = "2622 ���ɾ�"
End If
Set rs = Get_SqlserveRs("select a.�ֿ��� from tblStockManage a  inner join tblstock b on LEFT(a.�ֿ���,CHARINDEX(' ',a.�ֿ���)-1)=b.�ⷿ����  where b.�ֿ�����='��ͨ��' and  a.Ա�����='" & Trim(strRealName) & "'  ")
If rs.RecordCount > 0 Then
    rs.MoveFirst
    For i = 1 To rs.RecordCount
        cbWarehouse.AddItem Trim(rs.Fields(0))
        rs.MoveNext
    Next
End If
    
 
 With Fps(0)
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        .MaxRows = 0
        .ColsFrozen = 5
        .MaxCols = FpsH.e_MCol1 - 1
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .Col = FpsH.E_CHOOSE
         .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        .ZOrder
        .ReDraw = True
    End With

With Fps(0)
 
    .Col = -1
    .Row = -1
    .Lock = True
    .SetText FpsH.E_CHOOSE, 0, "ѡ��"
    .ColWidth(FpsH.E_CHOOSE) = 2
    .SetText FpsH.e_order, 0, "���ݱ��"
    .ColWidth(FpsH.e_order) = 10
   ' .SetText FpsH.e_num, 0, "���"
   '   .ColWidth(FpsH.e_num) = 3
    .SetText FpsH.e_matpn, 0, "�Ϻ�"
      .ColWidth(FpsH.e_matpn) = 15
    .SetText FpsH.e_matname, 0, "��������"
      .ColWidth(FpsH.e_matname) = 15
    .SetText FpsH.E_QTY, 0, "����"
      .ColWidth(FpsH.E_QTY) = 10
    .SetText FpsH.e_wh, 0, "�ֿ�"
      .ColWidth(FpsH.e_wh) = 5
    .SetText FpsH.e_shop_order, 0, "������"
      .ColWidth(FpsH.e_shop_order) = 15
    .SetText FpsH.e_ma, 0, "���ϱ��"
      .ColWidth(FpsH.e_ma) = 15

    
 End With


With Fps(1)
 
    .MaxCols = FpsD.e_MCol2 - 1
    .Col = -1
    .Row = -1
    .Lock = True
    .SetText FpsD.e_order, 0, "���ݱ��"
    .ColWidth(FpsD.e_order) = 10

    .SetText FpsD.e_matpn, 0, "�Ϻ�"
    .ColWidth(FpsD.e_matpn) = 12
    .SetText FpsD.e_matname, 0, "��������"
    .ColWidth(FpsD.e_matname) = 12
    .SetText FpsD.e_mastock, 0, "�ֿ�"
      .ColWidth(FpsD.e_mastock) = 5
    .SetText FpsD.E_LOT, 0, "����"
      .ColWidth(FpsD.E_LOT) = 8
    .SetText FpsD.E_stock, 0, "���"
      .ColWidth(FpsD.e_stock_k) = 8
    .SetText FpsD.e_stock_k, 0, "ƥ������"
      .ColWidth(FpsD.E_stock) = 8
    .SetText FpsD.e_stock_s, 0, "ƥ�������"
      .ColWidth(FpsD.e_stock_s) = 8
    .SetText FpsD.e_date, 0, "�������"
      .ColWidth(FpsD.e_date) = 15
    .SetText FpsD.e_duedate, 0, "��Ч����"
      .ColWidth(FpsD.e_duedate) = 15
    .SetText FpsD.e_manum, 0, "���ϱ��"
    .ColWidth(FpsD.e_manum) = 12
    .SetText FpsD.E_maid, 0, "ID"
      .ColWidth(FpsD.E_maid) = 8
 End With
 
 
 With Fps(3)
    .Col = -1
    .Row = -1
    .Lock = True
    .SetText FpsO.e_choice, 0, "ѡ��"
    .ColWidth(FpsO.e_choice) = 2
     .SetText FpsO.e_ordernum, 0, "���ݱ��"
    .ColWidth(FpsO.e_ordernum) = 10
    .Row = -1
    .Col = FpsO.e_choice
    .CellType = CellTypeCheckBox

 End With
 
  With Fps(4)
    .MaxCols = FpsM.e_MCol1 - 1
    .Col = -1
    .Row = -1
    .Lock = True
    .SetText FpsM.E_CHOOSE, 0, "ѡ��"
    .ColWidth(FpsM.E_CHOOSE) = 2
     .SetText FpsM.e_order, 0, "���ݱ��"
    .ColWidth(FpsM.e_order) = 10
    
     .SetText FpsM.e_Item, 0, "���"
    .ColWidth(FpsM.e_Item) = 2
     .SetText FpsM.e_matname, 0, "��������"
    .ColWidth(FpsM.e_matname) = 15
     .SetText FpsM.e_matpn, 0, "�Ϻ�"
    .ColWidth(FpsM.e_matpn) = 15
    
     .SetText FpsM.e_matname, 0, "��������"
    .ColWidth(FpsM.e_matname) = 15
    
     .SetText FpsM.E_QTY_O, 0, "�޸�ǰ����"
    .ColWidth(FpsM.E_QTY_O) = 5
     .SetText FpsM.E_QTY_N, 0, "�޸ĺ�����"
    .ColWidth(FpsM.E_QTY_N) = 5
     .SetText FpsM.e_shop_order, 0, "������"
    .ColWidth(FpsM.e_shop_order) = 10
     .SetText FpsM.e_ma, 0, "���ϱ��"
    .ColWidth(FpsM.e_ma) = 15
   
    .Row = -1
    .Col = FpsM.E_CHOOSE
    .CellType = CellTypeCheckBox
    .Col = FpsM.E_QTY_N
    .Lock = False

 End With


End Sub

'���沼��

Private Sub Form_Resize()
    control_resize

End Sub


Private Sub control_resize()

    On Error Resume Next
    
    Fra.Move Fra.Left, Fra.Top, Me.Width - Fra.Left - 550, Me.Height - Fra.Top - 500
   ' Fps(3).Move Fps(3).Left, Fps(3).Top, Fps(3).Width, Me.Height / 2 - Fps(3).Top
  '  Fps(0).Move Fps(0).Left, Fps(0).Top, Fps(0).Width, Me.Height / 2 - Fps(0).Top
    Fps(1).Move Fps(1).Left, Fps(1).Top, Fra.Width - Fps(1).Left - 200, Fps(1).Height
    Fps(2).Move Fps(2).Left, Fps(2).Top, Fps(2).Width, Fra.Height - Fps(2).Top - 200
    Fps(4).Move Fps(4).Left, Fps(4).Top, Fra.Width - Fps(4).Left - 200, Fra.Height - Fps(4).Top - 200
 
 End Sub
 
Private Sub cmdquery_Click()
Dim rs1    As New ADODB.Recordset
Dim Rs2    As New ADODB.Recordset
Dim strsql1 As String
Dim strSql2 As String
Dim strwh As String


strsql1 = "  SELECT  distinct 0 as 'ѡ��',a.���ݱ�� " & _
         "  FROM  ERPBASE..tblStockSQ2 a  WHERE  a.���ձ�� = 0 and ����<>'/' AND a.�ֿ��� IN ('43','14','56','46','47','05','66') " & _
         "  AND convert(VARCHAR(100),a.��������,23) >= '2020-04-01' and a.���ݱ�� like 'L%'"
If Trim(cbWarehouse.text) <> "" Then
    strwh = Left(Trim(cbWarehouse.text), InStr(Trim(cbWarehouse.text), " ") - 1)
    strsql1 = strsql1 & " and �ֿ���='" & strwh & "'"
End If
If Trim(txtShop_Order.text) <> "" Then
    strsql1 = strsql1 & " and ������='" & Trim(txtShop_Order.text) & "'"
End If

If Trim(txtCust.text) <> "" Then
    strsql1 = strsql1 & " and ������ in( select ORDERNAME from erpdata..tblTSVworkorder where CUSTOMER='" & Trim(txtCust.text) & " ')"
End If

strsql1 = strsql1 & " ORDER BY a.���ݱ�� "



'strsql1 = " SELECT   0 as 'ѡ��' ,'L2004270017' AS ���ݱ��"

'strsql2 = "  SELECT  a.���ϱ��,replace(a.�ֿ���,' ','') as �ֿ�,  b.ID, b.����,b.��ǰ����,b.��ǰ���� AS ƥ������� ,CONVERT(VARCHAR(100),b.��������,23) as  ������� FROM ERPBASE..tblStockSQ2 a  " & _
          "  LEFT JOIN ERPBASE..tblstocknum b  ON b.���ϱ�� = a.���ϱ�� AND b.�ֿ��� = a.�ֿ��� AND b.��ǰ����>0 " & _
          "  WHERE a.���ձ�� = 0 AND a.�ֿ��� IN ('43','14','56','46','47','05','66','09') AND convert(VARCHAR(100),a.��������,23) >= '2020-04-01' " & _
          "  GROUP BY  a.�ֿ���,a.���ϱ��,b.ID, b.����,b.��ǰ����,b.��������  ORDER BY a.���ϱ��,b.��������"
          
          
          
    If rs1.State = adStateOpen Then rs1.Close
    rs1.Open strsql1, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    With Fps(3)
        .MaxRows = 0
    End With
    With Fps(2)
        .MaxRows = 0
    End With
    With Fps(1)
        .MaxRows = 0
    End With
    With Fps(0)
        .MaxRows = 0
    End With
    With Fps(4)
        .MaxRows = 0
    End With
    If Not rs1.EOF Then  '��ʾ��������

        Call ListDataType1(rs1)

    Else
        
        MsgBox "û�д��������뵥", vbInformation, "��ʾ"

        
        Exit Sub

    End If
          
End Sub


Private Sub ListDataType1(rs As ADODB.Recordset)
Dim order As String
Dim i As Integer


With Fps(3)
     
     Set .DataSource = rs
    .ColWidth(1) = 2
    
    .Row = -1
    .Col = FpsO.e_choice
    .CellType = CellTypeCheckBox
End With
    

End Sub


Private Sub ListDataType2(rs As ADODB.Recordset)

    With Fps(1)
        
        .MaxRows = 0
        .Lock = True

        Set .DataSource = rs

    End With
    
End Sub

Private Sub fps_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
    Dim strorder As String
    Dim strmatno As String
    Dim strwh As String
    Dim strqty As String
    Dim errflag As String
    Dim i As Integer
    Dim doqtyinstock As Double
   
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub

Select Case Index
Case 3
   Fps(1).MaxRows = 0
    strorder = ""
    With Fps(3)
        
        .Row = Row
        .Col = FpsO.e_choice
     
        .Value = Abs(Val(.Value) - 1)
        
        If .Value = 1 Then
            .Col = -1
            .ForeColor = &HFF8080
            .Col = FpsO.e_ordernum
            strorder = Trim$(.text)
    
           errflag = 0
                 
           If errflag = 1 Then
               .Value = 0
           Else
               Call Searchdetail_Byorder(strorder, 1)
           End If
           
           
        ElseIf .Value = 0 Then
            
            .Col = -1
            .ForeColor = vbBlack
            .Col = FpsO.e_ordernum
            strorder = Trim$(.text)
            Call Searchdetail_Byorder(strorder, 2)
        End If
        compareqtywithstock 'check����Ƿ����
    
    End With

Case 4

    With Fps(4)
       .Row = Row
       .Col = FpsM.E_CHOOSE
       .Value = Abs(Val(.Value) - 1)
        If .Value = 1 Then
            .Col = -1
            .ForeColor = &HFF&
        ElseIf .Value = 0 Then
            .Col = -1
            .ForeColor = vbBlack
        End If
    End With
 
    
End Select

End Sub

Private Sub compareqtywithstock()
Dim i As Integer
Dim j As Integer
Dim strorderlist As String
Dim strSql As String
Dim rs         As New ADODB.Recordset
Dim strwh As String
Dim strmatno As String
Dim strwh_fps0 As String
Dim strmatno_fps0 As String
Dim strmatno_list As String



cmdmap.Enabled = True
cmdSend.Enabled = True
With Fps(3)
For i = 1 To .MaxRows
    .Row = i
    .Col = FpsO.e_choice
    If Trim(.text) = "1" Then
        .Col = FpsO.e_ordernum
        
        If strorderlist = "" Then
            strorderlist = Trim(.text)
        Else
            strorderlist = strorderlist & "','" & Trim(.text)
        End If
    End If
Next
End With
If strorderlist = "" Then
    With Fps(2)
        .MaxRows = 0
       ' Set .DataSource = rs
    End With
    Exit Sub
End If

With Fps(0)
.Row = -1
.Col = -1
.ForeColor = &H0&
End With

strSql = "select t1. ���ϱ��,t3.�Ϻ�,t3.��������,t1.�ֿ���,t1.��������,t2.������� from (SELECT ���ϱ��,�ֿ���,sum(����) AS �������� FROM eRPBASE..tblStockSQ2 WHERE ���ݱ�� IN('" & strorderlist & "' ) GROUP BY ���ϱ��,�ֿ���) AS t1" & _
" LEFT JOIN (SELECT ���ϱ�� ,�ֿ���,sum(��ǰ����)  AS ������� FROM erpbase..tblStockNum GROUP BY �ֿ���,���ϱ��) AS t2 ON t1.���ϱ��=t2.���ϱ�� AND t1.�ֿ���=t2.�ֿ���" & _
" LEFT JOIN erpbase..tblSmainM2 t3 on t1.���ϱ��=t3.���ϱ��" & _
" Where t1.�������� > �������"
'strsql = "select t1. ���ϱ��,t1.�ֿ���,t1.��������,t2.������� from (SELECT ���ϱ��,�ֿ���,sum(����) AS �������� FROM eRPBASE..tblStockSQ2 WHERE ���ݱ�� IN('" & strorderlist & "' ) GROUP BY ���ϱ��,�ֿ���) AS t1" & _
" LEFT JOIN (SELECT ���ϱ�� ,�ֿ���,sum(��ǰ����)  AS ������� FROM erpbase..tblStockNum GROUP BY �ֿ���,���ϱ��) AS t2 ON t1.���ϱ��=t2.���ϱ�� AND t1.�ֿ���=t2.�ֿ���"




Set rs = Get_SqlserveRs(strSql)
With Fps(2)
.MaxRows = 0
Set .DataSource = rs
End With

With Fps(0)
.Row = -1
.Col = -1
.ForeColor = &H0&
End With

If rs.RecordCount > 0 Then

    rs.MoveFirst
    For i = 1 To rs.RecordCount
        strmatno = Trim(rs("���ϱ��"))
        If strmatno_list = "" Then
        
            strmatno_list = strmatno
        Else
            strmatno_list = strmatno_list & "','" & strmatno
        
        End If
        
        
        strwh = Trim(rs("�ֿ���"))
        With Fps(0)
            For j = 1 To .MaxRows
                .Row = j
                .Col = FpsH.e_ma
                strmatno_fps0 = Trim(.text)
                .Col = FpsH.e_wh
                strwh_fps0 = Trim(.text)
                If strmatno = strmatno_fps0 And strwh = strwh_fps0 Then
                    .Col = -1
                    .ForeColor = &HFF&
                    cmdmap.Enabled = False
                    cmdSend.Enabled = False
                End If
            Next
            
        End With
        rs.MoveNext
    Next
     
     With Fps(4)
        .MaxRows = 0
        strSql = "  SELECT a.���ݱ��,a.���,b.�Ϻ�,b.��������,a.����� as ԭ����,a.������,a.���ϱ��  FROM ERPBASE..tblStockSQ2 a,erpbase..tblSmainM2 b  where a.���ϱ��=b.���ϱ�� and  rtrim(a.���ݱ��) in ('" & strorderlist & "') and a.���ϱ�� in ('" & strmatno_list & "') order by  a.���ݱ��,a.��� "
    
        Set rs = Get_SqlserveRs(strSql)
        
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            For j = 1 To rs.RecordCount
                .MaxRows = .MaxRows + 1
                
                .SetText FpsM.E_CHOOSE, .MaxRows, 0
                .SetText FpsM.e_order, .MaxRows, Trim$("" & rs!���ݱ��)
                .SetText FpsM.e_Item, .MaxRows, Trim$("" & rs!���)
                .SetText FpsM.e_matpn, .MaxRows, Trim$("" & rs!�Ϻ�)
                .SetText FpsM.e_matname, .MaxRows, Trim$("" & rs!��������)
                .SetText FpsM.E_QTY_O, .MaxRows, Trim$("" & rs!ԭ����)
                .SetText FpsM.E_QTY_N, .MaxRows, ""
               '.SetText FpsM.e_wh, .MaxRows, Trim$("" & rs!�ֿ�)
                .SetText FpsM.e_shop_order, .MaxRows, Trim$("" & rs!������)
                .SetText FpsM.e_ma, .MaxRows, Trim$("" & rs!���ϱ��)

                rs.MoveNext
            Next

        End If

    End With
    
Else
    With Fps(0)
    .Row = -1
    .Col = -1
    .ForeColor = &H0&
    End With
    
End If


End Sub



Private Sub Searchmat_Bymatno(order As String, matno As String, wh As String, QTY As String, intBJ As Integer)
    Dim i          As Integer
    Dim j          As Integer
    Dim strSql     As String
    Dim rs         As New ADODB.Recordset
    
    With Fps(1)

        strSql = " SELECT t1.ID,t1.���ϱ��,t2.�Ϻ�,t2.��������, t1.�ֿ���,t1.����,t1.��������,t1.��Ч����,t1.��ǰ���� ,t1.�ۼ�,t1.��ǰ����-t1.�ۼ�  AS �ۼ������� FROM  ( " & _
        " select b.ID,b.���ϱ��,b.�ֿ���,b.����,b.��������,b.��Ч����,b.��ǰ����,sum(a.��ǰ����)  AS  ����,sum(a.��ǰ����)-" & QTY & "  AS ����," & _
        " CASE WHEN sum(a.��ǰ����)-" & QTY & " <=0 THEN b.��ǰ���� WHEN sum(a.��ǰ����)-" & QTY & "<b.��ǰ���� THEN b.��ǰ����-(sum(a.��ǰ����)-" & QTY & " ) ELSE 0 END AS �ۼ� " & _
        " from erpbase..Posting_temp  a,erpbase..Posting_temp  b " & _
        " where a.����<=b.���� AND a.���ϱ�� =b.���ϱ�� AND a.�ֿ���=b.�ֿ���  AND a.���ϱ�� ='" & matno & "'  AND b.�ֿ���='" & wh & "' AND a.��ǰ����>0 AND b.��ǰ����>0 " & _
        " group by b.ID,b.���ϱ��,b.��������,b.��Ч����,b.����,b.��ǰ����,b.�ֿ���) t1 " & _
        " inner join erpbase..tblSmainM2 t2 on t1.���ϱ��=t2.���ϱ�� " & _
        " GROUP BY t1.ID,t1.���ϱ��,t1.����,t1.��������,t1.��Ч����,t1.��ǰ����,t1.�ۼ�,t1.�ֿ���,t2.�Ϻ�,t2.�������� ORDER BY t1.��Ч����"
                
                  
     '" Where t1.�ۼ� > 0 " & _

    
        Set rs = Get_SqlserveRs(strSql)
           
            
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                For j = 1 To rs.RecordCount
                    .MaxRows = .MaxRows + 1
                    
                    .SetText FpsD.e_order, .MaxRows, order
                    .SetText FpsD.e_manum, .MaxRows, Trim$("" & rs!���ϱ��)
                    .SetText FpsD.e_matpn, .MaxRows, Trim$("" & rs!�Ϻ�)
                    .SetText FpsD.e_matname, .MaxRows, Trim$("" & rs!��������)
                    .SetText FpsD.e_mastock, .MaxRows, Trim$("" & rs!�ֿ���)
                    .SetText FpsD.E_maid, .MaxRows, Trim$("" & rs!id)
                    .SetText FpsD.E_LOT, .MaxRows, Trim$("" & rs!����)
                    .SetText FpsD.E_stock, .MaxRows, Trim$("" & rs!��ǰ����)
                    .SetText FpsD.e_stock_k, .MaxRows, Trim$("" & rs!�ۼ�)
                    .SetText FpsD.e_stock_s, .MaxRows, Trim$("" & rs!�ۼ�������)
                    .SetText FpsD.e_date, .MaxRows, Trim$("" & rs!��������)
                    .SetText FpsD.e_duedate, .MaxRows, Trim$("" & rs!��Ч����)
                    If Val(rs!�ۼ�) > 0 Then
                        .Row = .MaxRows
                        .Col = -1
                        .BackColor = &HFFFF&
                    End If
                    
                    AddSql2 ("update erpbase..Posting_temp set ��ǰ����=��ǰ����-" & rs!�ۼ� & " where id=" & Trim$("" & rs!id))

                    rs.MoveNext
                Next
    
            End If
        End With



End Sub


Private Sub Searchdetail_Byorder(order As String, intBJ As Integer)
    Dim i          As Integer
    Dim j          As Integer
    Dim strSql     As String
    Dim rs         As New ADODB.Recordset
    Dim order_temp   As String



    If intBJ = 1 Then '��ѡ

        With Fps(0)
           If .MaxRows > 0 Then
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = FpsH.e_order
                     order_temp = Trim$(.text)
                    If order_temp = order Then
                        Exit Sub
                    End If
                Next
            End If
     
                    
            strSql = "  SELECT a.���ݱ��, a.���ϱ��,b.�Ϻ�,b.��������,replace(a.�ֿ���,' ','') as �ֿ�,sum(a.�����) as ����,a.������  FROM ERPBASE..tblStockSQ2 a,erpbase..tblSmainM2 b  where a.���ϱ��=b.���ϱ�� and  rtrim(a.���ݱ��)= '" & order & "' group by  a.���ݱ��, a.���ϱ��,b.�Ϻ�,b.��������,replace(a.�ֿ���,' ',''),a.������ "
         
         
            Set rs = Get_SqlserveRs(strSql)
               
                
                If rs.RecordCount > 0 Then
                    rs.MoveFirst
                    For j = 1 To rs.RecordCount
                        .MaxRows = .MaxRows + 1
                        
                        .SetText FpsH.E_CHOOSE, .MaxRows, 1
                        .SetText FpsH.e_order, .MaxRows, Trim$("" & rs!���ݱ��)
                      '  .SetText FpsH.e_num, .MaxRows, Trim$("" & rs!���)
                        .SetText FpsH.e_ma, .MaxRows, Trim$("" & rs!���ϱ��)
                        .SetText FpsH.e_matpn, .MaxRows, Trim$("" & rs!�Ϻ�)
                        .SetText FpsH.e_matname, .MaxRows, Trim$("" & rs!��������)
                        .SetText FpsH.E_QTY, .MaxRows, Trim$("" & rs!����)
                        .SetText FpsH.e_wh, .MaxRows, Trim$("" & rs!�ֿ�)
                        .SetText FpsH.e_shop_order, .MaxRows, Trim$("" & rs!������)
    
                        rs.MoveNext
                    Next
        
                End If
            End With

   ElseIf intBJ = 2 Then 'ȡ����ѡ

        With Fps(0)

            For i = .MaxRows To 1 Step -1
                    .Row = i
                    .Col = FpsH.e_order
                    order_temp = Trim$(.text)

                If order_temp = order Then
                    .DeleteRows i, 1
                    .MaxRows = .MaxRows - 1

                End If

            Next

        End With
        With Fps(4)

            For i = .MaxRows To 1 Step -1
                    .Row = i
                    .Col = FpsM.e_order
                    order_temp = Trim$(.text)

                If order_temp = order Then
                    .DeleteRows i, 1
                    .MaxRows = .MaxRows - 1

                End If

            Next

        End With
    End If

End Sub






Private Sub Fra_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
