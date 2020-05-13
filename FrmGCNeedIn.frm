VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmGCNeedIn 
   Caption         =   "GC客户待入库资料"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13575
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
   ScaleHeight     =   9960
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin FPSpreadADO.fpSpread fps 
      Height          =   9615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   13335
      _Version        =   196613
      _ExtentX        =   23521
      _ExtentY        =   16960
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
      SpreadDesigner  =   "FrmGCNeedIn.frx":0000
      TextTip         =   2
   End
End
Attribute VB_Name = "FrmGCNeedIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
    E_ID = 0                 'id
    E_Wo                     '工单号
    E_LotId                  'LotID
    E_WaferId                'WaferID
    E_GoodDie                '数量
    E_Qbox                   '箱号
    E_ContainName            '主批
    E_End
    
End Enum

Private Sub Form_Load()
IniFpsHeader
ShowData
End Sub


Private Sub ShowData()
Set reportRS = GetGCNeedIn()

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

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
        
        .SetText E_FPS0.E_ID, 0, "序号"
        .SetText E_FPS0.E_Wo, 0, "工单号"
        .SetText E_FPS0.E_LotId, 0, "LotID"
        .SetText E_FPS0.E_WaferId, 0, "WaferID"
        .SetText E_FPS0.E_GoodDie, 0, "Die数量"
        .SetText E_FPS0.E_Qbox, 0, "小箱号"
        .SetText E_FPS0.E_ContainName, 0, "主批号"
      
        
        .ColWidth(E_FPS0.E_ID) = 6
        .ColWidth(E_FPS0.E_Wo) = 15
        .ColWidth(E_FPS0.E_LotId) = 15
        .ColWidth(E_FPS0.E_WaferId) = 15
        .ColWidth(E_FPS0.E_GoodDie) = 15
        .ColWidth(E_FPS0.E_Qbox) = 15
        .ColWidth(E_FPS0.E_ContainName) = 15


        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .ReDraw = True
    End With
    
    
    

End Sub
