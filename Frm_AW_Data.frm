VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_AW_Data 
   Caption         =   "��Ϊ��������"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   22365
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
   ScaleHeight     =   9090
   ScaleWidth      =   22365
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   16695
      Begin VB.TextBox txtText1 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdOutput 
         BackColor       =   &H00FFFF00&
         Caption         =   "����"
         Height          =   720
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H0080FF80&
         Caption         =   "��ѯ"
         Height          =   720
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   420
      End
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   5655
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   2400
      Width           =   16575
      _Version        =   524288
      _ExtentX        =   29236
      _ExtentY        =   9975
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
      SpreadDesigner  =   "Frm_AW_Data.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
End
Attribute VB_Name = "Frm_AW_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ��ͷ������ʼ��POS
Enum E_FPS
E_CASE_NO = 1   ' ���
E_PRODUCT   ' ��Ʒ����
E_PKG_BATCH_NO ' ��װ���κ�
E_DATE         ' ����
E_QTY          ' ����
E_PKG_STYLE    ' ��װ��ʽ
E_ORDER_NO     ' ��������
E_INV_NO       ' ��Ʊ��
E_NG           ' ��Ʒ����Ʒ
E_END
End Enum

Private Sub cmdOutput_Click()
Dim OddNumber As String
OddNumber = UCase(Trim(txtText1))
Dim cmd_sql As String

If OddNumber = "" Then
    MsgBox "�����뵥��"
    Exit Sub
Else
    cmd_sql = "select distinct b.���, d.MPN_DESC as ��Ʒ����,d.ZX_INVOICE as ��װ���κ�, substring( convert(varchar(100), datepart(YY,f.ERPCREATEDATE)),3,2) + convert(varchar(100) , datepart(WW,f.ERPCREATEDATE)) as ����, " & _
    "b.����,d.comp_code as ��װ��ʽ,  d.PO_NUM as ��������,'' as ��Ʊ��,CASE b.�ϸ��� WHEN '0' THEN '��Ʒ' Else '����Ʒ' END  as ��Ʒ����Ʒ From erpdata..tblStockMove a ,erpdata..tblStockmovesub b ,ERPBASE..tblmappingData c ,ERPBASE..tblCustomerOI d,erpdata..tblTSVwaferlist e" & _
    ",erpdata..tblTSVworkorder f Where b.���ݱ�� = a.���ݱ�� and b.������� = a.��� and c.SUBSTRATEID = b.���̿���� and d.ID = c.FILENAME and e.WAFERID = c.SUBSTRATEID " & _
   " and f.ORDERNAME = e.ORDERNAME and a.���ݱ�� = '" & OddNumber & "'"
End If

SqlServer2ExporToExcel (cmd_sql)


End Sub

Private Sub cmdQuery_Click()
Dim cmd_sql As String
Dim OddNumber As String

OddNumber = UCase(Trim(txtText1))

If OddNumber = "" Then
    MsgBox "�����뵥��"
    Exit Sub
Else
    cmd_sql = "select distinct b.���, d.MPN_DESC,d.ZX_INVOICE, substring( convert(varchar(100), datepart(YY,f.ERPCREATEDATE)),3,2) + convert(varchar(100), datepart(WW,f.ERPCREATEDATE)), " & _
    "b.����,d.comp_code,  d.PO_NUM,'',CASE b.�ϸ��� WHEN '0' THEN '��Ʒ' Else '����Ʒ' END From erpdata..tblStockMove a ,erpdata..tblStockmovesub b ,ERPBASE..tblmappingData c ,ERPBASE..tblCustomerOI d,erpdata..tblTSVwaferlist e" & _
    ",erpdata..tblTSVworkorder f Where b.���ݱ�� = a.���ݱ�� and b.������� = a.��� and c.SUBSTRATEID = b.���̿���� and d.ID = c.FILENAME and e.WAFERID = c.SUBSTRATEID " & _
   " and f.ORDERNAME = e.ORDERNAME and a.���ݱ�� = '" & OddNumber & "'"
End If

Set mainItemRS = getSqlServerStr2(cmd_sql)

With fps(0)
        .MaxRows = 0
        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
        End If
End With

End Sub

' �������
Private Sub Form_Load()

Call InitFpsHeader

End Sub
' ��ʼ����ͷ
Private Sub InitFpsHeader()
With fps(0)
    .ReDraw = False
    .MaxCols = E_FPS.E_END - 1
    .MaxRows = 0
        
    .DAutoHeadings = False
    .DAutoCellTypes = False
    .DAutoSizeCols = DAutoSizeColsNone
        
    .Col = -1
    .Row = -1
    .Lock = True
    .OperationMode = OperationModeNormal
    .TypeVAlign = TypeVAlignCenter
    .SelForeColor = &HFF8080
            
    ' �����ͷ��
    .SetText E_FPS.E_CASE_NO, 0, "���"
    .SetText E_FPS.E_PRODUCT, 0, "��Ʒ����"
    .SetText E_FPS.E_PKG_BATCH_NO, 0, "��װ���κ�"
    .SetText E_FPS.E_DATE, 0, "����"
    .SetText E_FPS.E_QTY, 0, "����"
    .SetText E_FPS.E_PKG_STYLE, 0, "��װ��ʽ"
    .SetText E_FPS.E_ORDER_NO, 0, "��������"
    .SetText E_FPS.E_INV_NO, 0, "��Ʊ��"
    .SetText E_FPS.E_NG, 0, "��Ʒ/����Ʒ"
          
    ' ������
    .ColWidth(1) = 10
    .ColWidth(2) = 20
    .ColWidth(3) = 20
    .ColWidth(4) = 5
    .ColWidth(5) = 10
    .ColWidth(6) = 10
    .ColWidth(7) = 20
    .ColWidth(8) = 5
    .ColWidth(9) = 5
    
    ' ����߶�
    .RowHeight(0) = 20
    .RowHeight(-1) = 15
    
    .ReDraw = True
    
End With


End Sub
