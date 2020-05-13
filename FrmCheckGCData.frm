VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmCheckGCData 
   Caption         =   "GC数据确认"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdTestNo 
      BackColor       =   &H0080FF80&
      Caption         =   "测试版本号设定"
      Height          =   360
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "时间备注设定"
      Height          =   360
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "BC资料导入"
      Height          =   360
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "查询结果"
      Height          =   9135
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   11055
      Begin FPSpreadADO.fpSpread fps 
         Height          =   8895
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   11055
         _Version        =   196613
         _ExtentX        =   19500
         _ExtentY        =   15690
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
         SpreadDesigner  =   "FrmCheckGCData.frx":0000
         TextTip         =   2
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   345
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox TxtLotID 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lot ID："
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "FrmCheckGCData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
'    E_ID = 1                 'id
    E_LotId = 1                'LotId
    E_WaferId                'Waferid
    E_MarkingCode            'markingcode
    E_DieQty                 'DieQty
    E_Date                   '日期
    E_End
    
End Enum

Dim LotTdTemp As String
Dim reportRS As New ADODB.Recordset

Private Sub CmdTestNo_Click()
FrmTestSetUp.Show

End Sub

Private Sub Command1_Click()
If Trim(TxtLotID.Text) = "" Then
      MsgBox "请先输入要查询的LotID !", vbInformation, "友情提示"
      
Else
    LotTdTemp = UCase(Trim(TxtLotID.Text))
    ShowData_Where LotTdTemp
    
    
End If


End Sub

Private Sub ShowData_Where(str_temp As String)
Set reportRS = GetReportData_Where(str_temp)

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
        Else
            MsgBox "查询不到数据，请与市场部联系！", vbInformation, "友情提示"
        End If
End With

End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Form_Load()
IniFpsHeader
End Sub

Private Sub IniFpsHeader()
    With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        '砞竚Α
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
  
'        .SetText E_FPS0.E_ID, 0, "序号"
        .SetText E_FPS0.E_LotId, 0, "LotID"
        .SetText E_FPS0.E_WaferId, 0, "WaferID"
        .SetText E_FPS0.E_MarkingCode, 0, "MarkingCode"
        .SetText E_FPS0.E_DieQty, 0, "Die数量"
        .SetText E_FPS0.E_Date, 0, "日期"

        
        
'        .ColWidth(E_FPS0.E_ID) = 10
        .ColWidth(E_FPS0.E_LotId) = 15
        .ColWidth(E_FPS0.E_WaferId) = 15
        .ColWidth(E_FPS0.E_MarkingCode) = 10
        .ColWidth(E_FPS0.E_DieQty) = 12
        .ColWidth(E_FPS0.E_Date) = 15
 

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub

Private Sub TxtLotID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
