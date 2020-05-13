VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_PackingQty 
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   19005
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FraDN 
      Caption         =   "DNºÅ"
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   17295
      Begin VB.CommandButton cmdExport 
         Caption         =   "µ¼           ³ö"
         Height          =   600
         Left            =   6960
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "²é           Ñ¯"
         Height          =   600
         Left            =   4320
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtDN 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNºÅ:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   450
      End
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   6255
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   18135
      _Version        =   524288
      _ExtentX        =   31988
      _ExtentY        =   11033
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
      SpreadDesigner  =   "Frm_PackingQty.frx":0000
      TextTip         =   2
      AppearanceStyle =   0
   End
End
Attribute VB_Name = "Frm_PackingQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExport_Click()
Dim DNID As String

DNID = Trim$(txtDN.Text)

If DNID = "" Then
    MsgBox ("ÇëÊäÈëDNºÅ")
Else
    ExportPackInfo (DNID)


End If



End Sub

Private Sub cmdQuery_Click()

Dim DNID As String

DNID = Trim$(txtDN.Text)

If DNID = "" Then
    MsgBox ("ÇëÊäÈëDNºÅ")
Else
    QueryPackInfo (DNID)


End If



End Sub

Private Sub QueryPackInfo(DN As String)
Dim sql As String

sql = "select * from fh where Delivery= '" + DN + "'order by Delivery,BatchNumber,idd desc"

Set mainItemRs = getSqlStr2(sql)

With fps(0)
   .MaxRows = 0
        
    If mainItemRs.RecordCount > 0 Then
        Set .DataSource = mainItemRs
       
    End If
End With

End Sub

Private Sub ExportPackInfo(DN As String)
Dim sql As String

sql = "select * from fh where Delivery= '" + DN + "'order by Delivery,BatchNumber,idd desc"

SqlServer2ExporToExcel (sql)

End Sub
