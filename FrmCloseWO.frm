VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmCloseWO 
   Caption         =   "�����ر�"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
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
   ScaleHeight     =   9675
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "���ɹر�"
      Height          =   9375
      Left            =   10200
      TabIndex        =   7
      Top             =   240
      Width           =   10215
      Begin VB.CommandButton Command4 
         Caption         =   "������ϸ"
         Height          =   465
         Left            =   4320
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "��ѯ����"
         Height          =   465
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   8415
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   9975
         _Version        =   524288
         _ExtentX        =   17595
         _ExtentY        =   14843
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
         SpreadDesigner  =   "FrmCloseWO.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�ɹر�"
      Height          =   9375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   9975
      Begin VB.CommandButton Command1 
         Caption         =   "��ѯ����"
         Height          =   465
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "������ϸ"
         Height          =   465
         Left            =   4320
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   8415
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   9735
         _Version        =   524288
         _ExtentX        =   17171
         _ExtentY        =   14843
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
         SpreadDesigner  =   "FrmCloseWO.frx":0470
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ʹر�"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   9600
      Width           =   20295
      Begin VB.CheckBox chkPatchClose 
         BackColor       =   &H0000C000&
         Caption         =   "�����ر�"
         Height          =   495
         Left            =   11760
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CmdQuit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�˳�"
         Height          =   480
         Left            =   16800
         TabIndex        =   14
         Top             =   360
         Width           =   990
      End
      Begin VB.CommandButton CmdClose 
         BackColor       =   &H000000FF&
         Caption         =   "�رչ���"
         Height          =   480
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   990
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
      Begin MSDataListLib.DataCombo DtComb 
         Height          =   315
         Left            =   6240
         TabIndex        =   2
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label LblWo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   315
         Left            =   5640
         TabIndex        =   4
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmCloseWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mainItemRS As New ADODB.Recordset

Dim bomRS2     As New ADODB.Recordset

Dim bomRS3     As New ADODB.Recordset

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Enum E_FPS0          'Detail�֭�

    E_id = 0                 'id��
    E_WOId                   'wo
    E_PRODUCT                '�Ϻ�
    E_CreatedQty             '������
    E_InvQty                 '�����
    E_WipQty                 '������
    E_FinishRate             '�깤����
    E_BomFlag                'Bom����flag
    E_CreateDate             '����ʱ��
    E_DateCnt                '��������
    E_CloseFlag              '�Ƿ�ر�
    
    E_End
    
End Enum

Private Enum E_FPS1          'Detail�֭�

    E_id = 0                 'id��
    E_WOId                   'wo
    E_PRODUCT                   '�Ϻ�
    E_CreatedQty             '������
    E_InvQty                 '�����
    E_WipQty                 '������
    E_FinishRate             '�깤����
    E_BomFlag                'Bom����flag
    E_CreateDate             '����ʱ��
    E_DateCnt                '��������
    E_CloseFlag              '�Ƿ�ر�
    E_End
    
End Enum

Private Sub CmdClose_Click()

    Dim cmd_sql As String

    Dim rs      As New ADODB.Recordset

    cmd_sql = "select orderName from [erpdata].[dbo].[Vw_TSV_CloseWo] where wipQty<1 and BomStatus='��'"
    Set rs = getSqlServerStr2(cmd_sql)

    If rs.RecordCount <= 0 Then
        MsgBox "û�пɹرչ���, ��ȷ��"
        Exit Sub

    End If

    If chkPatchClose.Value = 0 Then
        Single_Close
    Else
    
        Do
            Call Patch_Close(rs(0))
    
            rs.MoveNext
            Sleep (500)
        Loop Until rs.BOF = True
    
        MsgBox "����ɾ���ɹ�"

    End If

End Sub

Private Function Patch_Close(ordername As String)

    Dim userid As String

    userid = UCase(gUserName)

    '2015-11-24 jiayun add check �߱�
    If ordername = "" Then
        MsgBox "��ѡ�񹤵���!", vbInformation, "������ʾ"
        Exit Function
    Else

        'У��Wo�Ƿ���ȷ
        If (Not JudgeOracleCloseWo(ordername)) Then
            MsgBox "��ȷ�Ϲ������Ƿ���ȷ!", vbInformation, "������ʾ"
            Exit Function
   
        End If
 
    End If

    '�ж� Oracle�� Wip���Ƿ������ݣ�����У�������رա�
    If Combo1.Text = "TSV" Then

        If (JudgeOracleWipWo(ordername)) Then
            'MsgBox "�ñʹ�����" & woTemp & " ������Mes Wip�ϣ������Թرգ�"
            'Exit Function
    
        End If

    End If

    '2013-05-51 jiayun add

    '�ж�ERP���Ƿ����û������ϣ�������ڣ��������

    If Combo1.Text = "WLO" Then
        Set bomRS2 = GetWLOWoBomLing(ordername)

        If bomRS2.RecordCount > 0 Then
            MsgBox "�ñʹ�������ϵͳ�л�����û���죬�����Թرչ�����"
            Exit Function

        End If

    End If

    Call DoCloseWoNew(ordername, userid)

    If Combo1.Text = "TSV" Then

        Call IniWO(1)

    ElseIf Combo1.Text = "WLO" Then
        Call IniWO(2)

    End If

    '�ӱ���ѡ�񹤵��ر�

    'DoFtpData
    'GetFpsData

    Dim aa As Integer

    aa = 0

End Function

Private Sub Single_Close()

    Dim userid      As String

    Dim queryWoTemp As String

    userid = UCase(gUserName)
    queryWoTemp = ""

    '2015-11-24 jiayun add check �߱�
    If Combo1.Text = "" Then
        MsgBox "��ѡ��������", vbInformation, "������ʾ"
        Exit Sub

    End If

    queryWoTemp = UCase(Trim(DtComb.Text))

    If queryWoTemp = "" Then
        MsgBox "��ѡ�񹤵���!", vbInformation, "������ʾ"
        Exit Sub
     
    Else

        'У��Wo�Ƿ���ȷ
        If (Not JudgeOracleCloseWo(queryWoTemp)) Then
   
            MsgBox "��ȷ�Ϲ������Ƿ���ȷ!", vbInformation, "������ʾ"
            Exit Sub
   
        End If
 
    End If

    'Dim woTemp As String
    '
    'If DtComb.Text = "" Then
    ' MsgBox "��ѡ��Ҫ�رյĹ�����"
    '     Exit Sub
    'End If
    'woTemp = DtComb.Text
    '
    '
    ''�ж� Oracle�� Wip���Ƿ������ݣ�����У�������رա�
    '
    'If Combo1.Text = "TSV" Then
    '
    '    If (JudgeOracleWipWo(Trim(woTemp))) Then
    '       MsgBox "�ñʹ�����" & woTemp & " ������Mes Wip�ϣ������Թرգ�"
    '       Exit Sub
    '
    '    End If
    '
    'End If
    '
    ''2013-05-51 jiayun add
    '
    ''�ж�ERP���Ƿ����û������ϣ�������ڣ��������
    '
    'If Combo1.Text = "WLO" Then
    '    Set bomRS2 = GetWLOWoBomLing(woTemp)
    '    If bomRS2.RecordCount > 0 Then
    '        MsgBox "�ñʹ�������ϵͳ�л�����û���죬�����Թرչ�����"
    '        Exit Sub
    '    End If
    'End If
    '
    '
    'Call DoCloseWo(woTemp)
    '
    'If Combo1.Text = "TSV" Then
    '
    'Call IniWO(1)
    '
    'ElseIf Combo1.Text = "WLO" Then
    'Call IniWO(2)
    '
    'End If

    Dim woTemp As String

    If DtComb.Text <> "" Then

        '���ʹرչ���

        woTemp = UCase(Trim(DtComb.Text))

        '�ж� Oracle�� Wip���Ƿ������ݣ�����У�������رա�

        If Combo1.Text = "TSV" Then

            If (JudgeOracleWipWo(Trim(woTemp))) Then
                ' MsgBox "�ñʹ�����" & woTemp & " ������Mes Wip�ϣ������Թرգ�"
                ' Exit Sub
    
            End If

        End If

        '2013-05-51 jiayun add

        '�ж�ERP���Ƿ����û������ϣ�������ڣ��������

        If Combo1.Text = "WLO" Then
            Set bomRS2 = GetWLOWoBomLing(woTemp)

            If bomRS2.RecordCount > 0 Then
                MsgBox "�ñʹ�������ϵͳ�л�����û���죬�����Թرչ�����"
                Exit Sub

            End If

        End If

        Call DoCloseWoNew(woTemp, userid)

        If Combo1.Text = "TSV" Then

            Call IniWO(1)

        ElseIf Combo1.Text = "WLO" Then
            Call IniWO(2)

        End If

    Else

        '�ӱ���ѡ�񹤵��ر�

        'DoFtpData
        'GetFpsData

        Dim aa As Integer

        aa = 0

    End If

    MsgBox "������" & woTemp & "�رճɹ� !", vbInformation, "��ʾ"

End Sub

Private Sub DoFtpData()

    Dim woTemp As String

    With Fps(0)

        For i = 1 To .MaxRows

            .Row = i
            .Col = 8

            If .Text = "1" Then

                .Row = i
                .Col = 1
                woTemp = Trim(.Text)
    
                Call DoCloseWo(woTemp)
 
            End If

        Next i

    End With

End Sub

Private Sub CmdDelMesWo_Click()

    Dim woTemp         As String

    Dim createDateTemp As Date

    Dim i              As Integer

    If Trim(TxtWO2.Text) = "" Then
        MsgBox "�����빤���ţ�"
        Exit Sub

    End If

    woTemp = UCase(Trim(TxtWO2.Text))

    '��ѯһ�£���ʹ�������������

    Set bomRS2 = GetWOCreateDate(woTemp)

    If bomRS2.RecordCount <= 0 Then
        MsgBox "��ʹ��������ڣ���ȷ�Ϲ����� ��"
        Exit Sub
    
    Else
        createDateTemp = CDate(bomRS2.Fields("createDate").Value)
    
        i = DateDiff("n", createDateTemp, Now)

        If i > 10 Then
            MsgBox "ʱ�����̫�ã�������ɾ�� ��"
            Exit Sub
        
        Else

            Call DelMesWO(woTemp)
    
        End If

    End If

End Sub

Private Sub CmdQuit_Click()
    Unload Me

End Sub

Private Sub Combo1_Change()

    If Combo1.Text = "TSV" Then

        Call IniWO(1)

    ElseIf Combo1.Text = "WLO" Then
        Call IniWO(2)

    End If

End Sub

Private Sub Combo1_Click()

    If Combo1.Text = "TSV" Then

        Call IniWO(1)

    ElseIf Combo1.Text = "WLO" Then
        Call IniWO(2)

    End If

End Sub

Private Sub IniFpsHeader()

    With Fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        '�]�m�榡
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .Col = E_FPS0.E_CloseFlag
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
  
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
          
        .SetText E_FPS0.E_id, 0, "���"
        .SetText E_FPS0.E_WOId, 0, "������"
        .SetText E_FPS0.E_PRODUCT, 0, "��Ʒ�Ϻ�"
        .SetText E_FPS0.E_CreatedQty, 0, "��������"
        .SetText E_FPS0.E_InvQty, 0, "�������"
        .SetText E_FPS0.E_WipQty, 0, "��������"
        .SetText E_FPS0.E_FinishRate, 0, "�깤����"
        .SetText E_FPS0.E_BomFlag, 0, "Bom���Ƿ�������"
        .SetText E_FPS0.E_CreateDate, 0, "��������"
        .SetText E_FPS0.E_DateCnt, 0, "�������"
        .SetText E_FPS0.E_CloseFlag, 0, "�Ƿ�ر�"

        .ColWidth(E_FPS0.E_id) = 5
        .ColWidth(E_FPS0.E_WOId) = 15
        .ColWidth(E_FPS0.E_PRODUCT) = 15
        .ColWidth(E_FPS0.E_CreatedQty) = 15
        .ColWidth(E_FPS0.E_InvQty) = 15
        .ColWidth(E_FPS0.E_WipQty) = 15
        .ColWidth(E_FPS0.E_FinishRate) = 15
        .ColWidth(E_FPS0.E_BomFlag) = 15
        .ColWidth(E_FPS0.E_CloseFlag) = 15

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .Col = E_FPS0.E_CloseFlag
        .Lock = False
        
        .ReDraw = True

    End With
    
    With Fps(1)
        .ReDraw = False
        .MaxCols = E_FPS1.E_End - 1
        .MaxRows = 0
        
        '�]�m�榡
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .Col = E_FPS1.E_CloseFlag
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
  
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
          
        .SetText E_FPS1.E_id, 0, "���"
        .SetText E_FPS1.E_WOId, 0, "������"
        .SetText E_FPS1.E_PRODUCT, 0, "��Ʒ�Ϻ�"
        .SetText E_FPS1.E_CreatedQty, 0, "��������"
        .SetText E_FPS1.E_InvQty, 0, "�������"
        .SetText E_FPS1.E_WipQty, 0, "��������"
        .SetText E_FPS1.E_FinishRate, 0, "�깤����"
        .SetText E_FPS1.E_BomFlag, 0, "Bom���Ƿ�������"
        .SetText E_FPS0.E_CreateDate, 0, "��������"
        .SetText E_FPS0.E_DateCnt, 0, "�������"
        .SetText E_FPS1.E_CloseFlag, 0, "�Ƿ�ر�"

        .ColWidth(E_FPS1.E_id) = 5
        .ColWidth(E_FPS1.E_WOId) = 15
        .ColWidth(E_FPS1.E_PRODUCT) = 15
        .ColWidth(E_FPS1.E_CreatedQty) = 15
        .ColWidth(E_FPS1.E_InvQty) = 15
        .ColWidth(E_FPS1.E_WipQty) = 15
        .ColWidth(E_FPS1.E_FinishRate) = 15
        .ColWidth(E_FPS1.E_BomFlag) = 15
        .ColWidth(E_FPS1.E_CreateDate) = 10
        .ColWidth(E_FPS1.E_DateCnt) = 10
        .ColWidth(E_FPS1.E_CloseFlag) = 15

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .Col = E_FPS1.E_CloseFlag
        .Lock = False
        
        .ReDraw = True

    End With

End Sub

Private Sub Command1_Click()
    '��ѯ����

    GetFpsData

End Sub

Private Sub Command2_Click()

    Dim sqlTemp As String

    ' sqlTemp = " select ������,PRODUCT as ��Ʒ�Ϻ�,Qty as ��������,FGQty as �������,Qty-FGQty as ��������,rtrim(round(cast((FGQty/Qty)* 100 as decimal(10,2)),2))+'%' as �깤����,BomStatus as Bom���Ƿ������� ,'' from ( " & _
    '" select x.������,x.PRODUCT,x.QTY,  erpdata.dbo.Get_TSV_WO_DieQty(x.������) as FGQty,erpdata.dbo.Get_TSV_WO_BomFinish(x.������) as BomStatus  from ( " & _
    '" select distinct e.������,f.PRODUCT ,f.QTY from  [erpbase].[dbo].[tblllplan] e , [erpdata].[dbo].[tblTSVworkorder] f " & _
    '" where    f.ORDERNAME=e.������ and  e.���߱��=1  and e.�깤���=0 ) X)Y "
  
'    sqlTemp = "SELECT  orderName,PRODUCT,woQty,invQty,wipQty,finishRate,BomStatus,flag FROM [erpdata].[dbo].[Vw_TSV_CloseWo] where wipQty<1 and BomStatus='��'"
    
    
    sqlTemp = "SELECT a.orderName,a.PRODUCT, a.woQty, a.invQty, a.wipQty, a.finishRate, a.BomStatus, CONVERT(varchar(100), b.ERPCREATEDATE, 23), DATEDIFF(day,b.ERPCREATEDATE,GETDATE()),a.flag " & _
"FROM [erpdata].[dbo].[Vw_TSV_CloseWo] a left join [erpdata].[dbo].[tblTSVworkorder] b on b.ORDERNAME = a.ORDERNAME where a.wipQty < 1 and a.BomStatus = '��' "
    
    
  
    SqlServerExporToExcel (sqlTemp)

End Sub

Private Sub Command3_Click()
    GetFpsData2

End Sub

Private Sub Command4_Click()

    Dim sqlTemp As String

    ' sqlTemp = " select ������,PRODUCT as ��Ʒ�Ϻ�,Qty as ��������,FGQty as �������,Qty-FGQty as ��������,rtrim(round(cast((FGQty/Qty)* 100 as decimal(10,2)),2))+'%' as �깤����,BomStatus as Bom���Ƿ������� ,'' from ( " & _
    '" select x.������,x.PRODUCT,x.QTY,  erpdata.dbo.Get_TSV_WO_DieQty(x.������) as FGQty,erpdata.dbo.Get_TSV_WO_BomFinish(x.������) as BomStatus  from ( " & _
    '" select distinct e.������,f.PRODUCT ,f.QTY from  [erpbase].[dbo].[tblllplan] e , [erpdata].[dbo].[tblTSVworkorder] f " & _
    '" where    f.ORDERNAME=e.������ and  e.���߱��=1  and e.�깤���=0 ) X)Y "
'
'    sqlTemp = "SELECT  orderName,PRODUCT,woQty,invQty,wipQty,finishRate,BomStatus,flag FROM [erpdata].[dbo].[Vw_TSV_CloseWo] where wipQty>0 or BomStatus='��'"



sqlTemp = "SELECT a.orderName,a.PRODUCT, a.woQty, a.invQty, a.wipQty, a.finishRate, a.BomStatus, CONVERT(varchar(100), b.ERPCREATEDATE, 23), DATEDIFF(day,b.ERPCREATEDATE,GETDATE()),a.flag " & _
"FROM [erpdata].[dbo].[Vw_TSV_CloseWo] a left join [erpdata].[dbo].[tblTSVworkorder] b on b.ORDERNAME = a.ORDERNAME where a.wipQty > 0 "



    SqlServerExporToExcel (sqlTemp)

End Sub

Private Sub DtComb_Click(Area As Integer)
    'ѡ�񹤵��󣬲�ѯ������Bom��

End Sub

Private Sub GetFpsData()
    '��ϸ����

    Set bomRS2 = GetSqlServerFpsCloseWo1()

    If bomRS2.RecordCount <= 0 Then
        MsgBox "��ϸ����û��������ݣ���ȷ��"
        Exit Sub

    End If

    With Fps(0)
        .MaxRows = 0

        If bomRS2.RecordCount > 0 Then
            Set .DataSource = bomRS2

        End If

    End With

End Sub

Private Sub GetFpsData2()
    '��ϸ����

    Set bomRS3 = GetSqlServerFpsCloseWo2()

    If bomRS3.RecordCount <= 0 Then
        MsgBox "��ϸ����û��������ݣ���ȷ��"
        Exit Sub

    End If

    With Fps(1)
        .MaxRows = 0

        If bomRS3.RecordCount > 0 Then
            Set .DataSource = bomRS3

        End If

    End With

End Sub

Private Sub Form_Activate()
    Combo1.SetFocus
    IniFpsHeader

End Sub

Private Sub Form_Load()
    Combo1.AddItem ("TSV")
    Combo1.AddItem ("WLO")
    chkPatchClose.Value = 0
    Combo1.Text = "TSV"

    'IniWO
End Sub

Private Sub IniWO(lineTypeTemp As Integer)
    Set mainItemRS = GetCloseWo(lineTypeTemp)
    Set DtComb.RowSource = mainItemRS
    DtComb.ListField = mainItemRS("productname").Name
    DtComb.BoundColumn = mainItemRS("PID").Name

End Sub
