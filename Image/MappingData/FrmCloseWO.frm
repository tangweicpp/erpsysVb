VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmCloseWO 
   Caption         =   "�����ر�"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19725
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
   ScaleWidth      =   19725
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "���ɹر�"
      Height          =   6375
      Left            =   9960
      TabIndex        =   7
      Top             =   240
      Width           =   9615
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
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   5415
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   960
         Width           =   9615
         _Version        =   196613
         _ExtentX        =   16960
         _ExtentY        =   9551
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
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�ɹر�"
      Height          =   6375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   9615
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
         Height          =   5415
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   960
         Width           =   9615
         _Version        =   196613
         _ExtentX        =   16960
         _ExtentY        =   9551
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
         SpreadDesigner  =   "FrmCloseWO.frx":4474
         TextTip         =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ʹر�"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   19455
      Begin VB.CommandButton CmdQuit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�˳�"
         Height          =   480
         Left            =   13320
         TabIndex        =   14
         Top             =   360
         Width           =   990
      End
      Begin VB.CommandButton CmdClose 
         BackColor       =   &H000000FF&
         Caption         =   "�رչ���"
         Height          =   480
         Left            =   11040
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
Dim bomRS2        As New ADODB.Recordset
Dim bomRS3        As New ADODB.Recordset
Private Enum E_FPS0          'Detail�֭�
    E_ID = 0                 'id��
    E_WOId                   'wo
    E_Product                   '�Ϻ�
    E_CreatedQty             '������
    E_InvQty                 '�����
    E_WipQty                 '������
    E_FinishRate             '�깤����
    E_BomFlag                'Bom����flag
    E_CloseFlag              '�Ƿ�ر�
    E_End
    
End Enum

Private Enum E_FPS1          'Detail�֭�
    E_ID = 0                 'id��
    E_WOId                   'wo
    E_Product                   '�Ϻ�
    E_CreatedQty             '������
    E_InvQty                 '�����
    E_WipQty                 '������
    E_FinishRate             '�깤����
    E_BomFlag                'Bom����flag
    E_CloseFlag              '�Ƿ�ر�
    E_End
    
End Enum



Private Sub CmdClose_Click()

Dim userid As String
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
       MsgBox "�ñʹ�����" & woTemp & " ������Mes Wip�ϣ������Թرգ�"
       Exit Sub
    
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


End Sub

Private Sub DoFtpData()
Dim woTemp As String

With fps(0)

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
Dim woTemp As String
Dim createDateTemp As Date

Dim i As Integer


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
    createDateTemp = CDate(bomRS2.fields("createDate").Value)
    
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
    With fps(0)
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
          
        .SetText E_FPS0.E_ID, 0, "���"
        .SetText E_FPS0.E_WOId, 0, "������"
        .SetText E_FPS0.E_Product, 0, "��Ʒ�Ϻ�"
        .SetText E_FPS0.E_CreatedQty, 0, "��������"
        .SetText E_FPS0.E_InvQty, 0, "�������"
        .SetText E_FPS0.E_WipQty, 0, "��������"
        .SetText E_FPS0.E_FinishRate, 0, "�깤����"
        .SetText E_FPS0.E_BomFlag, 0, "Bom���Ƿ�������"
        .SetText E_FPS0.E_CloseFlag, 0, "�Ƿ�ر�"

        .ColWidth(E_FPS0.E_ID) = 5
        .ColWidth(E_FPS0.E_WOId) = 15
        .ColWidth(E_FPS0.E_Product) = 15
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
    
    
    
    
     With fps(1)
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
          
        .SetText E_FPS1.E_ID, 0, "���"
        .SetText E_FPS1.E_WOId, 0, "������"
        .SetText E_FPS1.E_Product, 0, "��Ʒ�Ϻ�"
        .SetText E_FPS1.E_CreatedQty, 0, "��������"
        .SetText E_FPS1.E_InvQty, 0, "�������"
        .SetText E_FPS1.E_WipQty, 0, "��������"
        .SetText E_FPS1.E_FinishRate, 0, "�깤����"
        .SetText E_FPS1.E_BomFlag, 0, "Bom���Ƿ�������"
        .SetText E_FPS1.E_CloseFlag, 0, "�Ƿ�ر�"

        .ColWidth(E_FPS1.E_ID) = 5
        .ColWidth(E_FPS1.E_WOId) = 15
        .ColWidth(E_FPS1.E_Product) = 15
        .ColWidth(E_FPS1.E_CreatedQty) = 15
        .ColWidth(E_FPS1.E_InvQty) = 15
        .ColWidth(E_FPS1.E_WipQty) = 15
        .ColWidth(E_FPS1.E_FinishRate) = 15
        .ColWidth(E_FPS1.E_BomFlag) = 15
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

Dim sqltemp As String

' sqlTemp = " select ������,PRODUCT as ��Ʒ�Ϻ�,Qty as ��������,FGQty as �������,Qty-FGQty as ��������,rtrim(round(cast((FGQty/Qty)* 100 as decimal(10,2)),2))+'%' as �깤����,BomStatus as Bom���Ƿ������� ,'' from ( " & _
'" select x.������,x.PRODUCT,x.QTY,  erpdata.dbo.Get_TSV_WO_DieQty(x.������) as FGQty,erpdata.dbo.Get_TSV_WO_BomFinish(x.������) as BomStatus  from ( " & _
'" select distinct e.������,f.PRODUCT ,f.QTY from  [erpbase].[dbo].[tblllplan] e , [erpdata].[dbo].[tblTSVworkorder] f " & _
'" where    f.ORDERNAME=e.������ and  e.���߱��=1  and e.�깤���=0 ) X)Y "
  
  sqltemp = "SELECT  orderName,PRODUCT,woQty,invQty,wipQty,finishRate,BomStatus,flag FROM [erpdata].[dbo].[Vw_TSV_CloseWo] where wipQty<1 and BomStatus='��'"
  
  
  SqlServerExporToExcel (sqltemp)





End Sub

Private Sub Command3_Click()
GetFpsData2
End Sub

Private Sub Command4_Click()


Dim sqltemp As String

' sqlTemp = " select ������,PRODUCT as ��Ʒ�Ϻ�,Qty as ��������,FGQty as �������,Qty-FGQty as ��������,rtrim(round(cast((FGQty/Qty)* 100 as decimal(10,2)),2))+'%' as �깤����,BomStatus as Bom���Ƿ������� ,'' from ( " & _
'" select x.������,x.PRODUCT,x.QTY,  erpdata.dbo.Get_TSV_WO_DieQty(x.������) as FGQty,erpdata.dbo.Get_TSV_WO_BomFinish(x.������) as BomStatus  from ( " & _
'" select distinct e.������,f.PRODUCT ,f.QTY from  [erpbase].[dbo].[tblllplan] e , [erpdata].[dbo].[tblTSVworkorder] f " & _
'" where    f.ORDERNAME=e.������ and  e.���߱��=1  and e.�깤���=0 ) X)Y "
  
  
  sqltemp = "SELECT  orderName,PRODUCT,woQty,invQty,wipQty,finishRate,BomStatus,flag FROM [erpdata].[dbo].[Vw_TSV_CloseWo] where wipQty>0 or BomStatus='��'"
  
  
  SqlServerExporToExcel (sqltemp)


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


With fps(0)
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


With fps(1)
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




'IniWO
End Sub

Private Sub IniWO(lineTypeTemp As Integer)
Set mainItemRS = GetCloseWo(lineTypeTemp)
Set DtComb.RowSource = mainItemRS
DtComb.ListField = mainItemRS("productname").Name
DtComb.BoundColumn = mainItemRS("PID").Name

End Sub
