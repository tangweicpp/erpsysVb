VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm57 
   Caption         =   "57出矽力杰"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16815
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
   ScaleHeight     =   10065
   ScaleWidth      =   16815
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   14775
      Begin VB.Frame Frame2 
         Caption         =   "当前外箱明细"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   14655
         Begin VB.CommandButton btnClose 
            BackColor       =   &H00FFC0C0&
            Caption         =   "合 箱(&C)"
            Height          =   360
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtTempTrayLblQrCode 
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   6
            Top             =   398
            Width           =   8535
         End
         Begin VB.CommandButton btnBegin 
            BackColor       =   &H00C0C0C0&
            Caption         =   "开 始(&B)"
            Height          =   360
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   840
            Width           =   975
         End
         Begin WMPLibCtl.WindowsMediaPlayer player1 
            Height          =   495
            Left            =   11520
            TabIndex        =   8
            Top             =   480
            Visible         =   0   'False
            Width           =   615
            URL             =   ""
            rate            =   1
            balance         =   0
            currentPosition =   0
            defaultFrame    =   ""
            playCount       =   1
            autoStart       =   -1  'True
            currentMarker   =   0
            invokeURLs      =   -1  'True
            baseURL         =   ""
            volume          =   50
            mute            =   0   'False
            uiMode          =   "full"
            stretchToFit    =   0   'False
            windowlessVideo =   0   'False
            enabled         =   -1  'True
            enableContextMenu=   -1  'True
            fullScreen      =   0   'False
            SAMIStyle       =   ""
            SAMILang        =   ""
            SAMIFilename    =   ""
            captioningID    =   ""
            enableErrorDialogs=   0   'False
            _cx             =   1085
            _cy             =   873
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "卷盘临时标签二维码"
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
            Left            =   240
            TabIndex        =   5
            Top             =   435
            Width           =   1620
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "当前外箱明细"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8175
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   1680
         Width           =   14655
         Begin FPSpreadADO.fpSpread fps 
            Height          =   7575
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   14175
            _Version        =   524288
            _ExtentX        =   25003
            _ExtentY        =   13361
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
            MaxCols         =   8
            MaxRows         =   0
            SpreadDesigner  =   "Frm57.frx":0000
            Appearance      =   1
            AppearanceStyle =   0
         End
      End
   End
End
Attribute VB_Name = "Frm57"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TARY_INFO

PO_NO As String
PART_NO As String
LOT_NO As String
DATE_CODE As String
QUANTITY As String
QR_CODE As String
SERIAL_NO As String

End Type

Private strTempTrayLblQrCodeList As String

'Begin scan
Private Sub btnBegin_Click()
txtTempTrayLblQrCode.Enabled = True
txtTempTrayLblQrCode.SetFocus
strTempTrayLblQrCodeList = ""
fps.MaxRows = 0

Call PlaySound("请依次扫描需要合箱的卷盘临时标签")
End Sub

Private Sub PlaySound(strSound As String)
player1.url = "\\10.160.1.84\public\media_source\" & strSound & ".wav"

End Sub

Private Sub Form_Load()
strTempTrayLblQrCodeList = ""

With fps
    .MaxRows = 0
    .MaxCols = 6
    .Col = -1
    .Row = -1
    .Lock = True
    
    .SetText 1, 0, "唯一码"
    .SetText 2, 0, "W/O#"
    .SetText 3, 0, "P/N"
    .SetText 4, 0, "QTY"
    .SetText 5, 0, "D/C"
    .SetText 6, 0, "LOT NO"
    .ColWidth(1) = 15
    .ColWidth(2) = 15
    .ColWidth(3) = 15

End With

End Sub

'Scan tray label qrcode
Private Sub txtTempTrayLblQrCode_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then Exit Sub
If Len(Trim$(txtTempTrayLblQrCode.text)) = 0 Then Exit Sub

Dim strTempTrayLblQrCode As String
strTempTrayLblQrCode = UCase$(Trim$(txtTempTrayLblQrCode.text))
'Check data
If Not CheckTempTrayLblQrCode(strTempTrayLblQrCode) Then
    txtTempTrayLblQrCode.text = ""
    Exit Sub
End If

'Get data
Call GetTempTrayLblInfo(strTempTrayLblQrCode)

Dim trayInfo As TARY_INFO
trayInfo = GetTempTrayLblInfo(strTempTrayLblQrCode)

Call ShowTrayInfo(trayInfo)
Call PlaySound("铝箔袋标签已打印,请更换标签")

txtTempTrayLblQrCode.text = ""
End Sub

'Check the tray qrcode
Private Function CheckTempTrayLblQrCode(strTempTrayLblQrCode As String) As Boolean
CheckTempTrayLblQrCode = False

'Check label format
If InStr(strTempTrayLblQrCode, "@$") = 0 Then
    MsgBox "请扫描正确的卷盘的临时标签二维码信息", vbCritical, "二维码格式错误"
    Exit Function
End If

'Check if repeated
If InStr(strTempTrayLblQrCodeList, strTempTrayLblQrCode) > 0 Then
    MsgBox "当前卷盘已扫描,请勿重复扫描", vbCritical, "重复扫描错误"
    Exit Function
End If







CheckTempTrayLblQrCode = True
End Function
 
Private Function GetTempTrayLblInfo(strTempTrayLblQrCode As String) As TARY_INFO
Dim trayInfo As TARY_INFO
Dim strArray() As String

strArray = Split(strTempTrayLblQrCode, "@$")
trayInfo.PO_NO = strArray(3)
trayInfo.PART_NO = strArray(1)
trayInfo.QUANTITY = strArray(7)
trayInfo.DATE_CODE = strArray(4)
trayInfo.LOT_NO = strArray(5)
trayInfo.SERIAL_NO = strArray(8)
trayInfo.QR_CODE = trayInfo.PO_NO + "!" + trayInfo.PART_NO + "!" + trayInfo.QUANTITY + "!" + trayInfo.DATE_CODE + "!" + trayInfo.LOT_NO

strTempTrayLblQrCodeList = strTempTrayLblQrCodeList + strTempTrayLblQrCode + "%%"
GetTempTrayLblInfo = trayInfo
End Function

Private Sub ShowTrayInfo(trayInfo As TARY_INFO)
Dim i As Integer

With fps
    .MaxRows = .MaxRows + 1
    .SetText 1, .MaxRows, trayInfo.SERIAL_NO
    .SetText 2, .MaxRows, trayInfo.PO_NO
    .SetText 3, .MaxRows, trayInfo.PART_NO
    .SetText 4, .MaxRows, trayInfo.QUANTITY
    .SetText 5, .MaxRows, trayInfo.DATE_CODE
    .SetText 6, .MaxRows, trayInfo.LOT_NO

End With

End Sub
