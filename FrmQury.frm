VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmQury 
   Caption         =   "查询"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
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
   ScaleHeight     =   7320
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   360
      Left            =   6600
      TabIndex        =   2
      Top             =   360
      Width           =   990
   End
   Begin VB.TextBox TxtLotid 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   4695
      Index           =   0
      Left            =   720
      TabIndex        =   12
      Top             =   2520
      Width           =   11415
      _Version        =   524288
      _ExtentX        =   20135
      _ExtentY        =   8281
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
      SpreadDesigner  =   "FrmQury.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DateCode："
      Height          =   195
      Left            =   10560
      TabIndex        =   18
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label LbldCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mpn"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11520
      TabIndex        =   17
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MPN码："
      Height          =   195
      Left            =   7920
      TabIndex        =   16
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label LblMpn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mpn"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8640
      TabIndex        =   15
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MTRL_NUM："
      Height          =   195
      Left            =   7560
      TabIndex        =   14
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label LblMtrlNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mtrl_num"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8640
      TabIndex        =   13
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "最近两天的客户OI对应厂内料号显示:"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   2040
      Width           =   2940
   End
   Begin VB.Label LblProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5400
      TabIndex        =   10
      Top             =   1440
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "成品料号："
      Height          =   195
      Left            =   4320
      TabIndex        =   9
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "测试版本号："
      Height          =   195
      Left            =   720
      TabIndex        =   8
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label LblTestNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1920
      TabIndex        =   7
      Top             =   1440
      Width           =   45
   End
   Begin VB.Label LblTrayType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5400
      TabIndex        =   6
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tray类型："
      Height          =   195
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "贴膜标志："
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label LblPF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source_batch_id："
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1380
   End
End
Attribute VB_Name = "FrmQury"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum E_FPS0          'Detail

    E_id = 0                 'id
    E_POId                  'po
    E_LotID                  'LotId
    E_ProductName            '料号
    E_End
    
End Enum

Dim listRS As New ADODB.Recordset

Private Sub Command1_Click()

    On Error Resume Next

    Dim lotIdTemp As String

    lotIdTemp = UCase(Trim(TxtLotid.Text))

    lotIdTemp = Replace(lotIdTemp, vbCrLf, "")
    lotIdTemp = Replace(lotIdTemp, vbCr, "")
    lotIdTemp = Replace(lotIdTemp, vbLf, "")

    Dim pfType   As String

    Dim trayType As String

    Dim testno   As String

    Dim ptFirst  As String

    pfType = GetString(lotIdTemp)
    LblPF.Caption = pfType

    trayType = GetTrayString(lotIdTemp)
    LblTrayType.Caption = trayType

    testno = GetTestNoString(lotIdTemp)
    LblTestNo.Caption = testno

    '成品料号
    '根据OI，查出成品料号的前9位

    ptFirst = GetFirstPtString(lotIdTemp)
    LblProduct.Caption = GetAllPtString(ptFirst, pfType, trayType, testno)

    '2014-09-15 jiayun add 市场部新需求

    LblMpn.Caption = GetMPNString(lotIdTemp)

    LblMtrlNum.Caption = GetMtrlNumString(lotIdTemp)

    LbldCode.Caption = GetDateCodeString(lotIdTemp)

End Sub

Private Sub Form_Load()
    TxtLotid.Text = "7610679003"
    IniFpsHeader

    GetFpsData

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
  
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
          
        .SetText E_FPS0.E_id, 0, "序号"
        .SetText E_FPS0.E_POId, 0, "PONum"
        .SetText E_FPS0.E_LotID, 0, "LotId"
        .SetText E_FPS0.E_ProductName, 0, "料号"
        
        .ColWidth(E_FPS0.E_id) = 5
        .ColWidth(E_FPS0.E_POId) = 15
        .ColWidth(E_FPS0.E_LotID) = 15
        .ColWidth(E_FPS0.E_ProductName) = 20

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        .ReDraw = True

    End With

End Sub

Private Sub GetFpsData()

    Dim i As Integer

    Set listRS = GetOINewProduct()

    If (listRS.RecordCount > 0) Then

        fps(0).MaxRows = listRS.RecordCount

        For i = 0 To listRS.RecordCount - 1

            With fps(0)
                .Row = i + 1
                .Col = E_FPS0.E_id
                .Text = i + 1
        
                .Row = i + 1
                .Col = E_FPS0.E_POId
                .Text = listRS.Fields(0).Value
        
                .Row = i + 1
                .Col = E_FPS0.E_LotID
                .Text = listRS.Fields(1).Value
        
                .Row = i + 1
                .Col = E_FPS0.E_ProductName
                .Text = getAutoWo(listRS.Fields(1).Value)
        
            End With
    
            listRS.MoveNext

        Next

    End If

End Sub

Private Function getAutoWo(lotidTemp2 As String) As String

    Dim lotIdTemp As String

    lotIdTemp = lotidTemp2

    Dim pfType   As String

    Dim trayType As String

    Dim testno   As String

    Dim ptFirst  As String

    pfType = GetString(lotIdTemp)
    'LblPF.Caption = pfType

    trayType = GetTrayString(lotIdTemp)
    'LblTrayType.Caption = trayType

    testno = GetTestNoString(lotIdTemp)
    'LblTestNo.Caption = testno

    '成品料号
    '根据OI，查出成品料号的前9位

    ptFirst = GetFirstPtString(lotIdTemp)
    getAutoWo = GetAllPtString(ptFirst, pfType, trayType, testno)

End Function

