VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmWLO_Bom 
   Caption         =   "WLO�¹���Bom��"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17085
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
   ScaleHeight     =   7005
   ScaleWidth      =   17085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdAddRow 
      Caption         =   "add"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton CmdDelRow 
      Caption         =   "del"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "����Bom"
      Height          =   6615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   16575
      Begin VB.CommandButton CmdReturn 
         Caption         =   "����ǰ����"
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Add�󱣴�"
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   6015
         Index           =   1
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   15975
         _Version        =   524288
         _ExtentX        =   28178
         _ExtentY        =   10610
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
         SpreadDesigner  =   "FrmWLO_Bom.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "FrmWLO_Bom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS1          'Bom�֭�
    E_ID = 1                 'id��
    E_BomID                  '���Ϲ淶���
    E_PT                     '�Ϻ�
    E_Mt                     '���ϱ��
    E_Name                   '���ƪ�
    E_Qty                    'ÿֻ����
    E_Rate                   '�����
    E_Unit                   '��λ
    E_Pt2                     '�Ϻ�2
    E_Mt2                     '���ϱ��2
    E_Name2                   '����2��
    E_Qty2                    'ÿֻ����2
    E_Rate2                   '�����
    E_Unit2                   '��λ2
    E_End
    
End Enum

Dim bomRS        As New ADODB.Recordset
Public ptTemp As String



Private Sub CmdAddRow_Click()
 With fps(1)
        .MaxRows = .MaxRows + 1
    End With
End Sub

Private Sub CmdDelRow_Click()
'ɾ����
'ȡ��ǰ�е�Id
Dim currentRowTemp As String

With fps(1)
    Set .DataSource = Nothing
    .Row = .ActiveRow
    .Col = 1
    currentRowTemp = .Text
    .DeleteRows .ActiveRow, 1
    .MaxRows = .MaxRows - 1
End With

'ɾ����ǰ��
Call DelRowBomData(currentRowTemp)

'Call GetBomData(ptTemp)



End Sub

Private Sub CmdReturn_Click()
FrmWLOApplyWO.Show

End Sub

Private Sub CmdSave_Click()
'������к󣬱��浽DB
Dim i As Integer
Dim tempID As String
Dim tempPt As String
Dim tempEMt As String
Dim tempName As String
Dim tempQty As String
Dim tempUnit As String
Dim tempRate As String

i = 1
With fps(1)
    For i = 1 To .MaxRows
        empInfAll = False
        
        .Row = i
        .Col = E_FPS1.E_ID
        tempID = .Text
        
        If tempID = "" Then
            .Row = i
            .Col = E_FPS1.E_PT
            tempPt = .Text
            
            .Row = i
            .Col = E_FPS1.E_Mt
            tempEMt = .Text
            
            .Row = i
            .Col = E_FPS1.E_Name
            tempName = .Text
            
            .Row = i
            .Col = E_FPS1.E_Qty
            tempQty = .Text
            
            .Row = i
            .Col = E_FPS1.E_Rate
            tempRate = .Text
            
            .Row = i
            .Col = E_FPS1.E_Unit
            tempUnit = .Text
            
            If tempPt <> "" And CDbl(tempQty) > 0 Then
            
                AddRecord tempPt, tempEMt, tempName, tempQty, tempUnit, tempRate
            End If
        
        End If
         
    Next i

End With

'ˢ��һ��fps

Call GetBomData(ptTemp)



End Sub

Public Sub AddRecord(tempPt As String, tempEMt As String, tempName As String, tempQty As String, tempUnit As String, tempRate As String)
'����ӵ��У����浽DB��
Dim cmdStr As String
         
cmdStr = "INSERT INTO [erpdata].[dbo].[WLOtblBillBomInitData]([PT],[WLID],[Name],[Qty],[Unit],[SHRateQty]) " & _
         " VALUES('" + tempPt + "','" + tempEMt + "','" + tempName + "','" + tempQty + "','" + tempUnit + "','" + tempRate + "')"


                             
 Call AddSql2(cmdStr)
 
End Sub


Private Sub Form_Load()
IniFpsBom

InitData

End Sub

'Private Sub fps_ComboSelChange(Index As Integer, ByVal Col As Long, ByVal Row As Long)
'Dim tempPT As String
'Dim tempE_Mt As String
'
'
'If Col = E_FPS1.E_Pt Then
'    With fps(1)
'            .Row = .ActiveRow
'            .Col = E_FPS1.E_BomId
'
'            If .Text <> "" Then
'            '���Ϲ淶���Ϊ�գ�˵�����¼ӵ���
'            .Row = .ActiveRow
'            .Col = E_FPS1.E_Pt
'            tempPT = UCase$(.Text)
'            If tempPT <> "" Then
'                tempE_Mt = GetE_Mt(tempPT)
'                .Row = .ActiveRow
'                .Col = E_FPS1.E_Mt
'                .Text = tempE_Mt
'
'
'            End If
'
'
'            End If
'
'    End With
'
'End If
'
'End Sub

Public Function GetE_Mt(tempPt As String) As String
'��ѯ���ϱ��
Dim cmdStr As String
Dim RSResult As String
cmdStr = "SELECT ���ϱ�� FROM dbo.tblSmainM2 WHERE �Ϻ�='" + tempPt + "' "
     
RSResult = GetSqlServerStr(cmdStr)
GetE_Mt = RSResult
End Function

Public Function GetE_Name(tempPt As String) As String
'��ѯ����
Dim cmdStr As String
Dim RSResult As String
cmdStr = "SELECT �������� FROM dbo.tblSmainM2 WHERE �Ϻ�='" + tempPt + "' "
     
RSResult = GetSqlServerStr(cmdStr)
GetE_Name = RSResult
End Function

Public Function GetE_Unit(tempPt As String) As String
'��ѯ��λ
Dim cmdStr As String
Dim RSResult As String
cmdStr = "SELECT ������λ���� FROM dbo.tblSmainM2 WHERE �Ϻ�='" + tempPt + "' "
     
RSResult = GetSqlServerStr(cmdStr)
GetE_Unit = RSResult
End Function





Private Sub InitData()
'ѡ���Ʒ�Ϻţ�����ʾBom

ptTemp = FrmWLOApplyWO.Text3.Text

'ptTemp = "18V117FD00CF"

'2012-04-18
'����һ����ʱ���������ÿ�β�ѯ������Bom;
'ɾ������

DelBomData
Call AddBomData(ptTemp)


Call GetBomData(ptTemp)



End Sub

Private Sub GetBomData(ptTemp As String)
'��ϸ����

Set bomRS = GetFpsWLOBom(ptTemp)
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

Private Sub DelBomData()
'ɾ����ʼ������

Dim sqlTemp As String

sqlTemp = "DELETE FROM [erpdata].[dbo].[WLOtblBillBomInitData]"
Call AddSql2(sqlTemp)

End Sub

Private Sub DelRowBomData(rowId As String)
'ɾ����ǰ������

Dim sqlTemp As String

sqlTemp = "DELETE FROM [erpdata].[dbo].[WLOtblBillBomInitData] where id=" & rowId & ""
Call AddSql2(sqlTemp)

End Sub


Private Sub AddBomData(ptTemp As String)
'��ӳ�ʼ������

Dim sqlTemp As String

sqlTemp = "INSERT INTO [erpdata].[dbo].[WLOtblBillBomInitData] ( [MID],[PT],[WLID],[Name],[Qty],[SHRateQty],[Unit],[PT1],[WLID1],[Name1],[Qty1],[SHRateQty1],[Unit1] ) " & _
" SELECT  b.���Ϲ淶���,b.�Ϻ�,b.���ϱ��,b.����,cast(b.ÿֻ���� as varchar),b.���,b.��λ,b.�Ϻ�1,b.���ϱ��1,b.����1,cast(b.�������� as varchar),b.���1,b.��λ1 " & _
" FROM [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b " & _
" Where a.���Ϲ淶��� = b.���Ϲ淶��� AND a.���ϱ��='" & ptTemp & "'   and a.���߱��=3  "
 
Call AddSql2(sqlTemp)

End Sub



Private Sub IniFpsBom()
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
        .Lock = False
        

        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
      
        
        .SetText E_FPS1.E_ID, 0, "���"
        .SetText E_FPS1.E_BomID, 0, "���Ϲ淶���"
        .SetText E_FPS1.E_PT, 0, "�Ϻ�"
        .SetText E_FPS1.E_Mt, 0, "���ϱ��"
        .SetText E_FPS1.E_Name, 0, "����"
        .SetText E_FPS1.E_Qty, 0, "ÿֻ����"
        .SetText E_FPS1.E_Rate, 0, "���"
        .SetText E_FPS1.E_Unit, 0, "��λ"
        
        .SetText E_FPS1.E_Pt2, 0, "�����Ϻ�"
        .SetText E_FPS1.E_Mt2, 0, "�������ϱ��"
        .SetText E_FPS1.E_Name2, 0, "��������"
        .SetText E_FPS1.E_Qty2, 0, "����ÿֻ����"
        .SetText E_FPS1.E_Rate2, 0, "�������"
        .SetText E_FPS1.E_Unit2, 0, "���ϵ�λ"
    
        
        
        .ColWidth(E_FPS1.E_ID) = 6
        .ColWidth(E_FPS1.E_BomID) = 12
        .ColWidth(E_FPS1.E_PT) = 14
        .ColWidth(E_FPS1.E_Mt) = 14
        .ColWidth(E_FPS1.E_Name) = 14
        .ColWidth(E_FPS1.E_Qty) = 10
        .ColWidth(E_FPS1.E_Rate) = 6
        .ColWidth(E_FPS1.E_Unit) = 8
        
        .ColWidth(E_FPS1.E_Pt2) = 14
        .ColWidth(E_FPS1.E_Mt2) = 14
        .ColWidth(E_FPS1.E_Name2) = 14
        .ColWidth(E_FPS1.E_Qty2) = 10
         .ColWidth(E_FPS1.E_Rate2) = 6
        .ColWidth(E_FPS1.E_Unit2) = 8
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .Col = E_FPS1.E_ID
        .Lock = True
        
        .Col = E_FPS1.E_BomID
        .Lock = True
        
        
        .ReDraw = True
    End With
    
    
    

End Sub

Private Sub fps_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim tempPt As String
Dim tempE_Mt As String
Dim tempName As String
Dim tempUnit As String


If Col = E_FPS1.E_PT Then
    With fps(1)
            .Row = .ActiveRow
            .Col = E_FPS1.E_BomID
            
            If .Text = "" Then
            '���Ϲ淶���Ϊ�գ�˵�����¼ӵ���
            .Row = .ActiveRow
            .Col = E_FPS1.E_PT
            tempPt = Trim(UCase$(.Text))
            If tempPt <> "" Then
                tempE_Mt = GetE_Mt(tempPt)
                .Row = .ActiveRow
                .Col = E_FPS1.E_Mt
                .Text = tempE_Mt
                
                
                '����
                tempName = GetE_Name(tempPt)
                .Row = .ActiveRow
                .Col = E_FPS1.E_Name
                .Text = tempName
                
                '��λ
                tempUnit = GetE_Unit(tempPt)
                .Row = .ActiveRow
                .Col = E_FPS1.E_Unit
                .Text = tempUnit
                
                
            
            End If
            
            
            End If

    End With
    
End If

End Sub
