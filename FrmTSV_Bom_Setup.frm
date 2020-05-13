VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmTSV_Bom_Setup 
   Caption         =   "TSV Bom �趨"
   ClientHeight    =   9765
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   17085
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox DcbProductName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "ȡ��"
      Height          =   480
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   990
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "��"
      Height          =   480
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   990
   End
   Begin VB.TextBox TxtBom2 
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox TxtBom1 
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo DcbTYBomName 
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   7335
      Index           =   0
      Left            =   720
      TabIndex        =   8
      Top             =   2280
      Width           =   12855
      _Version        =   524288
      _ExtentX        =   22675
      _ExtentY        =   12938
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
      SpreadDesigner  =   "FrmTSV_Bom_Setup.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ͨ��Bom��"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ƷBom��"
      Height          =   195
      Left            =   6360
      TabIndex        =   1
      Top             =   360
      Width           =   840
   End
   Begin VB.Label LblProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ʒ�Ϻţ�"
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "FrmTSV_Bom_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Bom�֭�
    E_id = 1                 'id��
    E_PRODUCT                '���Ϲ淶���
    E_BomID                  'Bom��
    E_BomName                 'Bom�Ϻ�
    E_BomDate                 '��������
    
    E_End
    
End Enum

Dim bomRS        As New ADODB.Recordset
Dim bomRS1        As New ADODB.Recordset
Dim oiRS        As New ADODB.Recordset
Public ptTemp As String




Private Sub cmdQuery_Click()
'��ѯ
Dim sqlTemp As String

Dim sqlTemp1 As String

Dim sqlTemp2 As String

Dim sqltemp3 As String


  sqlTemp1 = "select a.[���Ϲ淶���],a.[���ϱ��],a.����,a.��������,b.�Ϻ� , b.���ϱ��, b.����, b.���, b.�ͺ�, b.ÿֻ����, b.���, b.��λ, b.���, b.��������" & _
             " from [erpdata].[dbo].[TSVtblSetMRule] a,[erpdata].[dbo].[TSVtblMRuleData] b Where a.[���Ϲ淶���] = b.[���Ϲ淶���]"

  sqlTemp2 = ""

  sqltemp3 = " order by a.[���Ϲ淶���],a.[���ϱ��], b.�Ϻ�"


 If Trim(TxtID.Text) <> "" Then

 sqlTemp2 = sqlTemp2 & " and a.���Ϲ淶��� like '%" & UCase(Trim(TxtID.Text)) & "%'"

 End If

  If Trim(TxtPT.Text) <> "" Then

 'sqltemp2 = sqltemp2 & " and a.���ϱ��='" & UCase(Trim(TxtPT.Text)) & "'"

 sqlTemp2 = sqlTemp2 & " and a.���ϱ�� like '%" & UCase(Trim(TxtPT.Text)) & "%'"


 End If

 If Trim(TxtPT2.Text) <> "" Then

 sqlTemp2 = sqlTemp2 & " and b.�Ϻ� like '%" & UCase(Trim(TxtPT2.Text)) & "%'"

 End If

 sqlTemp = sqlTemp1 & sqlTemp2 & sqltemp3



Set bomRS = GetFpsBomQuery(sqlTemp)
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

Private Sub cmdExit_Click()
TxtBom1.Text = ""
TxtBom2.Text = ""
End Sub



Private Sub CmdOK_Click()
Dim productNameTemp As String
Dim childBomId As String
Dim tyBomId As String


If Trim(DcbProductName.Text) = "" Or Trim(TxtBom1.Text) = "" Or Trim(DcbTYBomName.Text) = "" Then

    MsgBox "����û��ѡȫ����ȷ�ϣ�", vbInformation, "������ʾ"

    Exit Sub
Else

productNameTemp = Trim(DcbProductName.Text)
childBomId = Trim(TxtBom1.Text)
tyBomId = Left(Trim(DcbTYBomName.Text), InStr(Trim(DcbTYBomName.Text), "/") - 1)


End If

'�ж��Ƿ��Ѵ��ڣ�������ڣ������
If JudgeTSVBomAdd(productNameTemp) = True Then

    MsgBox "�Ѵ��ڣ��������!", vbInformation, "������ʾ"

Else



'sqlTemp = "insert into [erpdata].[dbo].[TSVtblBomSetup](ProductName,childBomNameID,Flag,CreateDate) values('" & productNameTemp & "','" & childBomId & "','Y',GETDATE())"
'AddSql2 (sqlTemp)

sqlTemp2 = "insert into [erpdata].[dbo].[TSVtblBomSetup](ProductName,childBomNameID,Flag,CreateDate) values('" & productNameTemp & "','" & tyBomId & "','Y',GETDATE())"
AddSql2 (sqlTemp2)


 MsgBox "��ӳɹ�!", vbInformation, "������ʾ"

GetBomData
End If



End Sub

Private Sub DcbProductName_Change()
Dim productTemp As String
productTemp = Trim(DcbProductName.Text)

'���ݳ�Ʒ�Ϻţ���ѯ����ƷBom


  Set oiRS = GetProductChildBom(productTemp)

    If (oiRS.RecordCount > 0) Then

        TxtBom1.Text = Trim(oiRS.Fields("���Ϲ淶���").Value)
        TxtBom2.Text = Trim(oiRS.Fields("���ϱ��").Value)

    End If


End Sub


Private Sub DcbProductName_Click()

Dim productTemp As String
productTemp = Trim(DcbProductName.Text)

'���ݳ�Ʒ�Ϻţ���ѯ����ƷBom


  Set oiRS = GetProductChildBom(productTemp)

    If (oiRS.RecordCount > 0) Then

        TxtBom1.Text = Trim(oiRS.Fields("���Ϲ淶���").Value)
        TxtBom2.Text = Trim(oiRS.Fields("���ϱ��").Value)

    End If

IniProduct2
End Sub

Private Sub Form_Load()
IniFpsBom
GetBomData

IniProduct
IniProduct2

End Sub




Private Sub IniProduct()
Set mainItemRS1 = GetBomProductName()
'Set DcbProductName.RowSource = mainItemRS1
'DcbProductName.ListField = mainItemRS1("productname").Name
'DcbProductName.BoundColumn = mainItemRS1("PID").Name


 If Not mainItemRS1.EOF Then
      mainItemRS1.MoveFirst
    While Not mainItemRS1.EOF
        With DcbProductName
            .AddItem Trim(mainItemRS1!productName)
        End With
        mainItemRS1.MoveNext
       Wend

 End If





End Sub

Private Sub IniProduct2()
Set mainItemRS = GetBomTYName()
Set DcbTYBomName.RowSource = mainItemRS
DcbTYBomName.ListField = mainItemRS("productname").Name
DcbTYBomName.BoundColumn = mainItemRS("PID").Name


End Sub




Private Sub GetBomData()
'��ϸ����

Set bomRS = GetFpsBomSetUp()
If bomRS.RecordCount <= 0 Then
    MsgBox "��ϸ����û��������ݣ���ȷ��"
    Exit Sub
End If

With fps(0)
        .MaxRows = 0
        If bomRS.RecordCount > 0 Then
            Set .DataSource = bomRS
        End If
End With

End Sub









Private Sub IniFpsBom()
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
        .Lock = False
        

        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        

        .SetText E_FPS0.E_id, 0, "���"
        .SetText E_FPS0.E_PRODUCT, 0, "��Ʒ�Ϻ�"
        .SetText E_FPS0.E_BomID, 0, "Bom���"
        .SetText E_FPS0.E_BomName, 0, "Bom�Ϻ�"
        .SetText E_FPS0.E_BomDate, 0, "��������"
        

        .ColWidth(E_FPS0.E_id) = 10
        .ColWidth(E_FPS0.E_PRODUCT) = 12
        .ColWidth(E_FPS0.E_BomID) = 12
        .ColWidth(E_FPS0.E_BomName) = 12
        .ColWidth(E_FPS0.E_BomDate) = 18
        
       

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        
        .ReDraw = True
    End With
    
    
    

End Sub

