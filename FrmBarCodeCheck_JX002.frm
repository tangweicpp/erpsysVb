VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FrmBarCodeCheck_JX002 
   Caption         =   "JX002标签条形码核对"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11010
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
   ScaleHeight     =   9180
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.TextBox txtScan 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   1005
         Width           =   6975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描顺序: 外箱->内盒->铝箔袋"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   3360
         Width           =   2445
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   1050
         Width           =   360
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   3480
         TabIndex        =   1
         Top             =   5880
         Visible         =   0   'False
         Width           =   1215
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
         _cx             =   2143
         _cy             =   873
      End
   End
End
Attribute VB_Name = "FrmBarCodeCheck_JX002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strWXCode As String
Dim strNXCode As String
Dim strLVCode As String
Dim lQtyAdd As Long
Dim lQtyAll As Long


Rem: 播放音频提醒
Private Sub Play(sFileName As String)
Dim sPath   As String
Dim sSuffix As String

sPath = "\\10.160.1.84\public\media_source\"
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix

End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then Exit Sub
Dim strScan As String
Dim lQtyTmp As Long

strScan = UCase(Trim$(txtScan.Text))
If InStr(strScan, " ") = 0 Or UBound(Split(strScan, " ")) <> 2 Then
    MsgBox "扫描不正确", vbInformation, "警告"
    txtScan.Text = ""
    Exit Sub

End If

If strWXCode = "" Then
    lQtyAdd = 0
    lQtyAll = CLng(Split(strScan, " ")(2))
    strWXCode = strScan
    Play ("外箱标签已扫描,请扫描内箱标签")
ElseIf strNXCode = "" Then
    If Split(strScan, " ")(0) <> Split(strWXCode, " ")(0) Then
        MsgBox "内盒标签与外箱标签不一致", vbCritical, "警告"
        txtScan.Text = ""
        Exit Sub

    End If

    If Split(strScan, " ")(1) <> Split(strWXCode, " ")(1) Then
        MsgBox "内盒标签与外箱标签不一致", vbCritical, "警告"
        txtScan.Text = ""
        Exit Sub

    End If

    lQtyTmp = CLng(Split(strScan, " ")(2))
    
    If lQtyAdd + lQtyTmp > lQtyAll Then
        MsgBox "内盒数量大于外箱数量,标签出错", vbCritical, "警告"
        txtScan.Text = ""
        Exit Sub

    End If
    
    lQtyAdd = lQtyAdd + lQtyTmp

    strNXCode = strScan
    Play ("内箱标签已扫描,请扫描铝箔袋标签")
ElseIf strLVCode = "" Then
    If Split(strScan, " ")(0) <> Split(strNXCode, " ")(0) Then
        MsgBox "内盒标签与铝箔袋标签不一致", vbCritical, "警告"
        txtScan.Text = ""
        Exit Sub

    End If

    If Split(strScan, " ")(1) <> Split(strNXCode, " ")(1) Then
        MsgBox "内盒标签与铝箔袋标签不一致", vbCritical, "警告"
        txtScan.Text = ""
        Exit Sub

    End If

    If Split(strScan, " ")(2) <> Split(strNXCode, " ")(2) Then
        MsgBox "内盒标签与铝箔袋标签不一致", vbCritical, "警告"
        txtScan.Text = ""
        Exit Sub

    End If

    If lQtyAdd = lQtyAll Then
        Play ("该外箱已全部比对完成,请继续比对其他外箱")
        strWXCode = ""
        strNXCode = ""
        strLVCode = ""
    Else
        Play ("该内箱已核对完成, 请继续下一个内箱")
        strNXCode = ""
    End If

End If

txtScan.Text = ""

End Sub
