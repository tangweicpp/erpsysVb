VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmTSV_WOBo 
   Caption         =   "TSV���� Bom�޸�"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExport 
      Caption         =   "��������"
      Height          =   480
      Left            =   8400
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdDel 
      BackColor       =   &H000000FF&
      Caption         =   "�޸�"
      Height          =   360
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "��ѯ"
      Height          =   360
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox TxtID 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   8415
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   17895
      _Version        =   524288
      _ExtentX        =   31565
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
      SpreadDesigner  =   "FrmTSV_WOBom.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label LblId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ţ�"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "FrmTSV_WOBo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS1          'Bom�֭�
    e_ID = 1                 'id��
    E_BomID                  '���Ϲ淶���
    E_Createdby              '������
    E_CreatedDate              '��������
    E_PT                     '�Ϻ�
    E_Mt                     '���ϱ��
    E_name                   '���ƪ�
    E_NameNew                   '���ƪ�
    E_GG                     '���
E_CloseFlag              '�Ƿ�ر�
  
    E_END
    
End Enum

Dim bomRS        As New ADODB.Recordset
Public ptTemp As String
Public bomProduct As String




Private Sub CmdAdd_Click()
Dim i As Integer
Dim qtyTemp As String
Dim bomIDtTemp As String
Dim bomProductTemp As String


 With fps(1)
        .MaxRows = .MaxRows + 1
        i = .MaxRows
        
        .Row = i - 1
        .Col = 1
        bomIDtTemp = .Text
        
        
        .Row = i - 1
        .Col = 2
        bomProductTemp = .Text
        
        .Row = i
        .Col = 1
        .Text = bomIDtTemp
        
        .Row = i
        .Col = 2
        .Text = bomProductTemp
        
 End With
 
 
 
 
 
 
End Sub

Private Sub CmdAddSave_Click()
Dim i As Integer
Dim tempProduct As String
Dim bomId As String
Dim product As String
Dim ptid As String
Dim ptname As String
Dim pttype As String
Dim pttypename As String
Dim qty As Double
Dim qtysh As Double
Dim unit As String
Dim notemp As String
Dim noName As String




With fps(1)

     For i = 1 To .MaxRows
     
      .Row = i
      .Col = 5
      tempProduct = Trim(.Text)
      
      If tempProduct = bomProduct Then
      
      '���뵽SqlServer��
      
'      sqlTemp = "Update [erpdata].[dbo].[TSVtblMRuleData]  Set ÿֻ���� = " & qtyTemp & " where ���Ϲ淶���='" & bomIDtTemp & "' and �Ϻ�='" & bomProductTemp & "'"
'        AddSql2 (sqlTemp)
        
      .Row = i
      .Col = 1
       bomId = Trim(.Text)
       
      .Row = i
      .Col = 2
       product = Trim(.Text)
       
'     .Row = i
'      .Col = 5
'       product = Trim(.Text)
       
     .Row = i
      .Col = 6
       ptid = Trim(.Text)
       
     .Row = i
      .Col = 7
       ptname = Trim(.Text)
       
      .Row = i
      .Col = 8
       pttype = Trim(.Text)
       
       
      .Row = i
      .Col = 9
       pttypename = Trim(.Text)
       
      .Row = i
      .Col = 10
       qty = CDbl(Trim(.Text))
       
     .Row = i
      .Col = 11
      qtysh = CDbl(Trim(.Text))
      
       
      .Row = i
      .Col = 12
       unit = Trim(.Text)
       
      .Row = i
      .Col = 13
       notemp = Trim(.Text)
       
      .Row = i
      .Col = 14
       noName = Trim(.Text)
       
       
       sqlTemp = " insert into [erpdata].[dbo].[TSVtblMRuleData](���Ϲ淶���,�����,�Ϻ�,���ϱ��,����,���," & _
      " �ͺ� , ÿֻ����,���,��λ,���,��������)" & _
      "  values ('" & bomId & "','" & product & "','" & tempProduct & "','" & ptid & "','" & ptname & "','" & pttype & "'," & _
      " '" & pttypename & "'," & qty & "," & qtysh & ",'" & unit & "','" & notemp & "','" & noName & "') "
      
        AddSql2 (sqlTemp)
         
        MsgBox "��ӳɹ�!", vbInformation, "������ʾ"
        
        cmdQuery_Click

      
      
      
      End If
      
      
     
     Next i
     
  
    
    

End With


End Sub

Private Sub cmdDel_Click()
Dim qtyTemp As String
Dim bomIDtTemp As String
Dim bomProductTemp As String


Dim woTemp As String
Dim mtID As String
Dim beforQty As Double
Dim afterQty As Double


With fps(1)


For i = 1 To .MaxRows

    .Row = i
    .Col = 10
    
    If .Text = "1" Then

    .Row = i
    .Col = 1
    woTemp = Trim(.Text)
    
     .Row = i
    .Col = 2
    mtID = Trim(.Text)
    
    If Mid(mtID, 1, 5) = "03.06" Or Mid(mtID, 1, 5) = "03.07" Or Mid(mtID, 1, 5) = "03.08" Then
    
    MsgBox "���Ʒ�ϺŲ������޸�����", vbInformation, "��ʾ"
    Exit Sub
    
        
    End If
    
    
    .Row = i
    .Col = 7
    beforQty = CDbl(Trim(.Text))
    
    .Row = i
    .Col = 8
    afterQty = CDbl(Trim(.Text))
    
    
    Call DoModifyWoQty(woTemp, mtID, beforQty, afterQty, gUserName)
 
    End If

Next i


End With



MsgBox "������" & woTemp & "Bom�����޸ĳɹ� !", vbInformation, "��ʾ"

cmdQuery_Click



End Sub

Private Sub cmdModify_Click()
Dim qtyTemp As String
Dim bomIDtTemp As String
Dim bomProductTemp As String


If Trim(TxtModify.Text) = "" Then

    MsgBox "����������Ϊ�գ�", vbInformation, "������ʾ"
    Exit Sub
 
 Else
    qtyTemp = CLng(Trim(TxtModify.Text))
    
End If

With fps(1)
    .Row = .ActiveRow
    .Col = 1
    bomIDtTemp = Trim(.Text)
    
    .Row = .ActiveRow
    .Col = 5
    bomProductTemp = Trim(.Text)
    
    

End With

     
'���µ�SqlServer

sqlTemp = "Update [erpdata].[dbo].[TSVtblMRuleData]  Set ÿֻ���� = " & qtyTemp & " where ���Ϲ淶���='" & bomIDtTemp & "' and �Ϻ�='" & bomProductTemp & "'"
AddSql2 (sqlTemp)

 MsgBox "�޸������ɹ�!", vbInformation, "������ʾ"

cmdQuery_Click


End Sub

Private Sub cmdExport_Click()

Dim sqlTemp As String

'sqlTemp = "SELECT DISTINCT case ���߱�� when 1 then 'TSV' when 2 then 'WLO' when 3 then 'WLC' end as ��������,������, ���ϱ��,�Ϻ�,����,����ͺ�,�ͺ�,����,ʵ������ FROM erpbase..tbllljh WHERE �깤���<>2  ORDER BY ��������,������"

'sqlTemp = "SELECT DISTINCT case ���߱�� when 1 then 'TSV' when 2 then 'WLO' when 3 then 'WLC' end as ��������,������ FROM erpbase..tbllljh WHERE �깤���<>2  AND ������ NOT LIKE '%G-%' AND ������ NOT LIKE '%M-%' AND ������ NOT LIKE '%D-%' " & _
' "  AND ���� <> ʵ������  ORDER BY ��������,������"

sqlTemp = "SELECT DISTINCT case ���߱�� when 1 then 'TSV' when 2 then 'WLO' when 3 then 'WLC' end as ��������,������, �Ϻ�,����, ʵ������  FROM erpbase..tbllljh WHERE �깤���<>2  AND ������ NOT LIKE '%G-%' AND ������ NOT LIKE '%M-%' AND ������ NOT LIKE '%D-%' " & _
 "  AND ���� <> ʵ������  ORDER BY ��������,������"

SqlServer2ExporToExcel (sqlTemp)

End Sub

Private Sub cmdQuery_Click()
'��ѯ
Dim sqlTemp As String

Dim sqlTemp1 As String

Dim sqlTemp2 As String

Dim sqltemp3 As String



  
   sqlTemp = " select rtrim(ltrim(a.������)) ������, rtrim(ltrim(a.���ϱ��)) ���ϱ��,rtrim(ltrim(b.�Ϻ�)) �Ϻ�,rtrim(ltrim(b.��������)) ��������,rtrim(LTRIM(b.����ͺ�)) ����ͺ� " & _
              " ,rtrim(ltrim(b.�ͺ�)) �ͺ�,a.����,'' as ����2,a.ʵ������ , 0 as b  from  [erpbase].[dbo].[tblllplan] a,dbo.tblSmainM2 b where a.������='" & UCase(Trim(TxtID.Text)) & "' and a.���߱��=1 and b.���ϱ��=a.���ϱ�� "
  
  




Set bomRS = GetFpsBomQuery(sqlTemp)
If bomRS.RecordCount <= 0 Then
    MsgBox "����û��������ݣ���ȷ��"
    Exit Sub
End If

With fps(1)
        .MaxRows = 0
        If bomRS.RecordCount > 0 Then
            Set .DataSource = bomRS
        End If
End With


End Sub

Private Sub Form_Activate()
TxtID.SetFocus

End Sub

Private Sub Form_Load()
IniFpsBom

End Sub









Private Sub GetBomData(ptTemp As String)
'��ϸ����

Set bomRS = GetFpsBom(ptTemp)
If bomRS.RecordCount <= 0 Then
    MsgBox "��ϸ����û��������ݣ���ȷ��"
    Exit Sub
End If

With fps(1)
        .MaxRows = 0
        If bomRS.RecordCount > 0 Then
            Set .DataSource = bomRS
        End If
End With

End Sub









Private Sub IniFpsBom()
    With fps(1)
    
    
      .ReDraw = False
        .MaxCols = E_FPS1.E_END - 1
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
          

    
    

        
        
        
        .SetText E_FPS1.e_ID, 0, "������"
        .SetText E_FPS1.E_BomID, 0, "���ϱ��"
        .SetText E_FPS1.E_Createdby, 0, "�Ϻ�"
        .SetText E_FPS1.E_CreatedDate, 0, "��������"
        .SetText E_FPS1.E_PT, 0, "����ͺ�"
        .SetText E_FPS1.E_Mt, 0, "�ͺ�"
        .SetText E_FPS1.E_name, 0, "ԭ������"
        .SetText E_FPS1.E_NameNew, 0, "�޸ĺ�����"
        .SetText E_FPS1.E_GG, 0, "ʵ������"
        .SetText E_FPS1.E_CloseFlag, 0, " ѡ��"
        
        

        .ColWidth(E_FPS1.e_ID) = 10
        .ColWidth(E_FPS1.E_BomID) = 12
        .ColWidth(E_FPS1.E_Createdby) = 12
        .ColWidth(E_FPS1.E_CreatedDate) = 12
        
        .ColWidth(E_FPS1.E_PT) = 14
        .ColWidth(E_FPS1.E_Mt) = 14
        .ColWidth(E_FPS1.E_name) = 14
        .SetText E_FPS1.E_NameNew, 0, "����"
         .ColWidth(E_FPS1.E_GG) = 14
          .ColWidth(E_FPS1.E_CloseFlag) = 14
       
        
    

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
               .Col = E_FPS1.E_NameNew
        .Lock = False
        
           .Col = E_FPS1.E_CloseFlag
        .Lock = False
        
        
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub fps_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

Dim bomProductTemp As String

   If (Col = E_FPS1.E_PT And NewCol = E_FPS1.E_Mt And NewRow = Row) Then
        With fps(1)
            .Row = Row
            .Col = Col

        bomProductTemp = .Text
        bomProduct = bomProductTemp
        
        '�����Ϻţ���ѯ�����Ϣ
        
          Set oiRS = GetProductChildBomAdd(bomProductTemp)
  
            If (oiRS.RecordCount > 0) Then
              
            .Row = Row
            .Col = Col + 1
            .Text = Trim(oiRS.Fields("���ϱ��").Value)
            
            .Row = Row
            .Col = Col + 2
            .Text = Trim(oiRS.Fields("��������").Value)
            
            .Row = Row
            .Col = Col + 3
            .Text = Trim(oiRS.Fields("����ͺ�").Value)
            
            .Row = Row
            .Col = Col + 4
            .Text = Trim(oiRS.Fields("�ͺ�").Value)
            .Row = Row
            .Col = Col + 6
            .Text = "0"
            
            .Row = Row
            .Col = Col + 7
            .Text = Trim(oiRS.Fields("������λ����").Value)
                
        
            End If
    End With
    End If



End Sub


