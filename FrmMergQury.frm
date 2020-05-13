VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form FrmMergQury 
   Caption         =   "合批信息查询"
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
   Begin VB.CommandButton Command2 
      Caption         =   "导出"
      Height          =   360
      Left            =   8880
      TabIndex        =   18
      Top             =   360
      Width           =   990
   End
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
      Height          =   5295
      Index           =   0
      Left            =   480
      TabIndex        =   11
      Top             =   1920
      Width           =   11655
      _Version        =   196613
      _ExtentX        =   20558
      _ExtentY        =   9340
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
      SpreadDesigner  =   "FrmMergQury.frx":0000
      TextTip         =   2
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DateCode："
      Height          =   195
      Left            =   10560
      TabIndex        =   17
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label LbldCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11520
      TabIndex        =   16
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MPN码："
      Height          =   195
      Left            =   7920
      TabIndex        =   15
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label LblMpn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8640
      TabIndex        =   14
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MTRL_NUM："
      Height          =   195
      Left            =   7560
      TabIndex        =   13
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label LblMtrlNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8640
      TabIndex        =   12
      Top             =   1440
      Width           =   45
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
Attribute VB_Name = "FrmMergQury"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail
    E_ID = 0                 'id
    E_LotId                  'LotId
    E_MPN                'mpn
    E_Mtrl_Num           'mtrl_num
    E_Tst_Program_Rev           'tst_program_rev
    E_DateCode           'datecode
    E_End
    
End Enum

Dim listRS        As New ADODB.Recordset



Private Sub Command1_Click()
On Error Resume Next

Dim lotIDtemp As String
lotIDtemp = UCase(Trim(TxtLotid.Text))


lotIDtemp = Replace(lotIDtemp, vbCrLf, "")
lotIDtemp = Replace(lotIDtemp, vbCr, "")
lotIDtemp = Replace(lotIDtemp, vbLf, "")


Dim pfType As String
Dim trayType As String
Dim testno As String

Dim ptFirst As String



pfType = GetString(lotIDtemp)
LblPF.Caption = pfType

trayType = GetTrayString(lotIDtemp)
LblTrayType.Caption = trayType

testno = GetTestNoString(lotIDtemp)
LblTestNo.Caption = testno

'成品料号
'根据OI，查出成品料号的前9位

ptFirst = GetFirstPtString(lotIDtemp)
LblProduct.Caption = GetAllPtString(ptFirst, pfType, trayType, testno)


'2014-09-15 jiayun add 市场部新需求

LblMpn.Caption = GetMPNString(lotIDtemp)

LblMtrlNum.Caption = GetMtrlNumString(lotIDtemp)

LbldCode.Caption = GetDateCodeString(lotIDtemp)

Dim mpnTemp As String
Dim mtrl_Num As String
Dim dateCodeTemp As String

'lotIDtemp
mpnTemp = LblMpn.Caption
mtrl_Num = LblMtrlNum.Caption
'testno
dateCodeTemp = LbldCode.Caption


sqlTemp = " insert into  TSV_AA_MergeQuery(lotID,mpn,mtrl_num,test_program_rev,dateCode) values ('" & lotIDtemp & "','" & mpnTemp & "','" & mtrl_Num & "','" & testno & "','" & dateCodeTemp & "')"
AddSql (sqlTemp)

 
GetFpsData


End Sub

Private Sub Command2_Click()

Dim sqlTemp As String

sqlTemp = "select lotID,mpn,mtrl_num,test_program_rev,dateCode from TSV_AA_MergeQuery order by createdate desc"
         
  ExporToExcel (sqlTemp)


End Sub

Private Sub Form_Load()

sqlTemp = " delete from  TSV_AA_MergeQuery"
AddSql (sqlTemp)


'TxtLotid.Text = "7610679003"
IniFpsHeader

'GetFpsData

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
          
        .SetText E_FPS0.E_ID, 0, "序号"
        .SetText E_FPS0.E_LotId, 0, "LotId"
        .SetText E_FPS0.E_MPN, 0, "MPN"
        .SetText E_FPS0.E_Mtrl_Num, 0, "Mtrl_Num"
        .SetText E_FPS0.E_Tst_Program_Rev, 0, "Tst_Program_Rev"
        .SetText E_FPS0.E_DateCode, 0, "DateCode"
        
        
        
    
        
        .ColWidth(E_FPS0.E_ID) = 10
        .ColWidth(E_FPS0.E_LotId) = 10
        .ColWidth(E_FPS0.E_MPN) = 10
        .ColWidth(E_FPS0.E_Mtrl_Num) = 10
        .ColWidth(E_FPS0.E_Tst_Program_Rev) = 18
        .ColWidth(E_FPS0.E_DateCode) = 10
     

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        

        
        .ReDraw = True
    End With
    
    
    

End Sub

Private Sub GetFpsData()
Set reportRS = GetMergAAQueryInf()

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With



End Sub

Private Function getAutoWo(lotidTemp2 As String) As String

Dim lotIDtemp As String
lotIDtemp = lotidTemp2
Dim pfType As String
Dim trayType As String
Dim testno As String

Dim ptFirst As String



pfType = GetString(lotIDtemp)
'LblPF.Caption = pfType

trayType = GetTrayString(lotIDtemp)
'LblTrayType.Caption = trayType

testno = GetTestNoString(lotIDtemp)
'LblTestNo.Caption = testno

'成品料号
'根据OI，查出成品料号的前9位

ptFirst = GetFirstPtString(lotIDtemp)
getAutoWo = GetAllPtString(ptFirst, pfType, trayType, testno)

End Function



