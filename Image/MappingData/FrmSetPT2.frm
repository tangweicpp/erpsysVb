VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmSetPT2 
   Caption         =   "�ͻ��Ϻ��볧���ϺŶ�Ӧ��ϵ���趨(��AA)"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13725
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
   ScaleHeight     =   9630
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   11640
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox TxtQtechPT 
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "����"
      Height          =   360
      Left            =   3960
      TabIndex        =   5
      Top             =   840
      Width           =   990
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�޸�"
      Height          =   360
      Left            =   6840
      TabIndex        =   4
      Top             =   840
      Width           =   990
   End
   Begin VB.TextBox TxtCustomerPT 
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   8055
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   12495
      _Version        =   196613
      _ExtentX        =   22040
      _ExtentY        =   14208
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
      SpreadDesigner  =   "FrmSetPT2.frx":0000
      TextTip         =   2
   End
   Begin MSDataListLib.DataCombo CmbCustomer 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������룺"
      Height          =   195
      Left            =   10800
      TabIndex        =   9
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ϻţ�"
      Height          =   195
      Left            =   6960
      TabIndex        =   3
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ����룺"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   900
   End
   Begin VB.Label LblPT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ��Ϻţ�"
      Height          =   195
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "FrmSetPT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail�֭�
'    E_ID = 1                 'id��
    E_Product = 1             '�ͻ�����
    E_Pf                      '�ͻ��Ϻ�
    E_Tray                    '��Ʒ�Ϻ�
    E_End
    
End Enum

Dim mainItemRS As New ADODB.Recordset
Dim reportRS As New ADODB.Recordset


Private Sub IniProductTestNo(productNameTemp As String)
Dim testval As String

Set mainItemRS = GetMainItemProduct(productNameTemp)
'Set DCbMainItem.RowSource = mainItemRS
testval = mainItemRS("typename").Name

'DCbMainItem.ListField = mainItemRS("typename").Name
'DCbMainItem.BoundColumn = mainItemRS("id").Name

TxtQtechPT.Text = mainItemRS.fields(1).Value

End Sub


'Private Sub IniTestNo()
'Set mainItemRS = GetMainItem()
'Set DCbMainItem.RowSource = mainItemRS
'DCbMainItem.ListField = mainItemRS("typename").Name
'DCbMainItem.BoundColumn = mainItemRS("id").Name
'
'End Sub

Private Sub CmdAdd_Click()
'����
Dim tempProductName As String
Dim tempPf As String
Dim tempTray As String
Dim tempTestNo As String


Dim sqlTemp As String

tempProductName = UCase(Trim(CmbCustomer.Text))
tempPf = TxtCustomerPT.Text
tempTray = TxtQtechPT.Text
secondCode = Trim(Text1.Text)



'�ж��Ƿ�������
 If tempProductName = "" Or tempPf = "" Or tempTray = "" Then
    MsgBox "�������������ύ��", vbInformation, "������ʾ"
    Exit Sub
 
 End If


 
sqlTemp = "insert into TBLSETPT(customercode,customerpt,qtechpt,flag,createby,createdate) values ('" & tempProductName & "','" & tempPf & "','" & tempTray & "','Y','Auto',sysdate)"
AddSql (sqlTemp)

 MsgBox "��ӳɹ�!", vbInformation, "������ʾ"
 
ShowData_Where

End Sub

Private Sub Command3_Click()
'�޸�

Dim tempProductName As String
Dim tempPf As String
Dim tempTray As String
Dim tempTestNo As String

Dim sqlTemp As String

tempProductName = UCase(Trim(TxtCustomerPT.Text))
tempPf = ComCustomer.Text
tempTray = CombTray.Text
tempTestNo = TxtQtechPT.Text



'�ж��Ƿ�������
 If tempProductName = "" Or tempPf = "" Or tempTray = "" Or tempTestNo = "" Then
    MsgBox "�������������ύ��", vbInformation, "������ʾ"
    Exit Sub
 
 End If
 
 '�ж��Ϻ��Ƿ����
If (Not JudgePtExist(tempProductName)) Then
   MsgBox "��ʣ�" & tempProductName & " �����ڣ������޸ģ�"
Exit Sub

End If


sqlTemp = " update TBLSETPT set pfStaus='" & tempPf & "',trayStaus='" & tempTray & "',testNo='" & tempTestNo & "',lastupdateby='Auto',lastupdatedate=sysdate where productName='" & tempProductName & "' and flag='Y' "

AddSql (sqlTemp)

 MsgBox "�޸ĳɹ�!", vbInformation, "������ʾ"
 
ShowData_Where



End Sub

Private Sub Form_Activate()
CmbCustomer.SetFocus



End Sub

Private Sub Form_Load()
'IniTestNo

IniCustomerName

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
        
        
        .SetText E_FPS0.E_Product, 0, "�ͻ�����"
        .SetText E_FPS0.E_Pf, 0, "�ͻ��Ϻ�"
        .SetText E_FPS0.E_Tray, 0, "��Ʒ�Ϻ�"

    
        
        .ColWidth(E_FPS0.E_Product) = 20
        .ColWidth(E_FPS0.E_Pf) = 20
        .ColWidth(E_FPS0.E_Tray) = 20

       
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        

        
        
        .ReDraw = True
    End With
    
    
ShowData_Where

End Sub

Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CmbCustomer.RowSource = mainItemRS
CmbCustomer.ListField = mainItemRS("productname").Name
CmbCustomer.BoundColumn = mainItemRS("PID").Name

End Sub


Private Sub ShowData_Where()
Set reportRS = GetptOtherCustPT()


With Fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub


Private Sub Label2_Click()

End Sub

Private Sub TxtCustomerPT_KeyPress(KeyAscii As Integer)
Dim tempProductName As String

If KeyAscii = 13 Then
'��ѯ���԰汾��
tempProductName = UCase(Trim(TxtCustomerPT.Text))
    If tempProductName = "" Then
    
     MsgBox "�������Ʒ�Ϻţ�", vbInformation, "������ʾ"
     
    Else
    IniProductTestNo tempProductName
    
    
    
    End If

End If

End Sub
