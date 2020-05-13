VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Frm37QRLblConverter 
   Caption         =   "37卷盘二维码标签转换工具"
   ClientHeight    =   11865
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10425
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
   ScaleHeight     =   11865
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   12015
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   10455
      Begin VB.TextBox txtFailed 
         ForeColor       =   &H000000FF&
         Height          =   9855
         Left            =   5280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtSuccess 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   9855
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtReelID 
         BackColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   390
         Width           =   2655
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   8640
         TabIndex        =   7
         Top             =   360
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
      Begin VB.Label lblFailed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转换失败的卷盘:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5280
         TabIndex        =   5
         Top             =   1200
         Width           =   1650
      End
      Begin VB.Label lblSuccess 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转换成功的卷盘:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘号"
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   540
      End
   End
End
Attribute VB_Name = "Frm37QRLblConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type STBox

    JOB As String
    DEV As String
    FactoryFlow As String
    lot As String
    QTY As String
    DATECODE As String
    testdateCode As String

End Type

Private lSuccess As Long

Private lFailed  As Long

Private Sub setPrintPath()
str37BCIDPath = "\\10.160.1.84\public\BarCode\37\37内盒带二维码\"   'QR
End Sub

Private Sub setTestPrintPath()
str37BCIDPath = "C:\test\"      ' 37B,C,R小标签
End Sub

Private Sub Form_Load()

Select Case gUserName

    Case "07885"
        Call setTestPrintPath

    Case Else
        Call setPrintPath

End Select

lSuccess = 0
lFailed = 0

End Sub

Private Sub txtReelID_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Or Len(Trim(txtReelID.Text)) = 0 Then Exit Sub
If PrintQrReelLbl(UCase(Trim$(txtReelID.Text))) = True Then
    lSuccess = lSuccess + 1
    lblSuccess.Caption = "转换成功的卷盘:" & lSuccess
    txtSuccess.Text = txtSuccess.Text & txtReelID.Text & vbCrLf
    Call PlaySound("转换成功")
Else
    lFailed = lFailed + 1
    lblFailed.Caption = "转换失败的卷盘:" & lFailed
    txtFailed.Text = txtFailed.Text & txtReelID.Text & vbCrLf
    Call PlaySound("转换失败")
End If
txtReelID.Text = ""
End Sub

Private Function PrintQrReelLbl(strReelID As String) As Boolean
Dim strSql      As String
Dim strTxt      As String
Dim strFlagTxt  As String
Dim StrFileName As String
Dim rsJobID     As New ADODB.Recordset
Dim tSTBox      As STBox
Dim strQrCode   As String

PrintQrReelLbl = False

'检查箱号
If Get_SqlserverCnt("select * from erpdata..tblPackMainInfSub where 箱号 = '" & strReelID & "' ") = 0 Then
    MsgBox "卷盘箱号不存在", vbCritical, "警告"
    Exit Function
End If

'获取数据,检查数据
tSTBox.lot = strReelID
If tSTBox.lot = "" Then
    MsgBox "卷盘号不可为空", vbCritical, "警告"
    Exit Function

End If

tSTBox.JOB = GetJobID(strReelID)
If tSTBox.JOB = "" Then
    MsgBox "JOB号不可为空", vbCritical, "警告"
    Exit Function

End If

tSTBox.DATECODE = GetDC(tSTBox.JOB)
If tSTBox.DATECODE = "" Then
    MsgBox "DC不可为空", vbCritical, "警告"
    Exit Function

End If

tSTBox.QTY = GetReelQty(strReelID)
If tSTBox.QTY = 0 Then
    MsgBox "数量不可0", vbCritical, "警告"
    Exit Function

End If

tSTBox.FactoryFlow = GetPN(strReelID)
If tSTBox.FactoryFlow = "" Then
    MsgBox "机种不可为空", vbCritical, "警告"
    Exit Function

End If

tSTBox.DEV = Replace(Replace(Replace$(tSTBox.FactoryFlow, ".P1", ""), ".P2", ""), ".P3", "")
If tSTBox.DEV = "" Then
    MsgBox "机种不可为空", vbCritical, "警告"
    Exit Function

End If

'拼接txt
strTxt = strTxt & tSTBox.DEV & "," & tSTBox.JOB & ",1T" & tSTBox.JOB & "," & tSTBox.DEV & "," & "1P" & tSTBox.DEV & "," & tSTBox.DATECODE & "," & tSTBox.DATECODE & "," & Mid(tSTBox.lot, 2) & "," & tSTBox.lot & "," & tSTBox.QTY & ",Q" & tSTBox.QTY & "," & tSTBox.DATECODE & "," & tSTBox.DATECODE & GetDevMark(tSTBox.DEV)
strTxt = strTxt & "," & tSTBox.FactoryFlow & "," & "6P" & tSTBox.FactoryFlow & "," & "10D" & tSTBox.DATECODE & ","
strQrCode = "[)>" & Chr(30) & "06" & Chr(29) & "1T" & tSTBox.JOB & Chr(29) & "1P" & tSTBox.DEV & Chr(29) & tSTBox.lot & Chr(29) & "Q" & tSTBox.QTY & Chr(29) & "6P" & tSTBox.FactoryFlow & Chr(29) & "10D" & tSTBox.DATECODE & Chr(30) & Chr(4)
strTxt = strTxt & strQrCode & vbCrLf
StrFileName = strReelID
Call CreateTxt(StrFileName, strTxt, str37BCIDPath)
PrintQrReelLbl = True

End Function

Private Function GetJobID(strReelID As String) As String
Dim strSql As String
Dim strRes As String

strSql = "select KEY_VALUE from erpdata..tblErpInStockDetailInfo a where SUBSTRING(a.KEY_VALUE,1,charindex('|',a.KEY_VALUE)-1) =  '" & strReelID & "' and a.KEY_NAME = 'CONTAINER_NAME' AND a.KEY_TYPE = 'T' and charindex('|',a.KEY_VALUE) > 0"
strRes = Get_SqlStr(strSql)
GetJobID = Mid(strRes, InStr(strRes, "|") + 1)
If GetJobID = "" Then
    GetJobID = Get_SqlStr("select customerlotid as jobid from erpdata..TblTSV_Tray_details where TRAYQBOXNUMBER = '" & strReelID & "'")

End If

End Function

Private Function GetDC(strJobID As String) As String
Dim strSql As String

If Right(strJobID, 1) = "M" Then
    strSql = "select DC from erptemp..TBL37TESTDC_m where JOBID = '" & strJobID & "'"
Else
    strSql = "select DC from erptemp..TBL37TESTDC where JOBID = '" & strJobID & "'"

End If

GetDC = Get_SqlStr(strSql)

End Function

Private Function GetReelQty(strReelID As String) As Long
Dim strSql As String

strSql = " select SUM(数量) from erpdata..tblPackMainInfSub where 箱号 = '" & strReelID & "' "
GetReelQty = Get_SqlserverNo(strSql)

End Function

Private Function GetPN(strReelID As String) As String
GetPN = UCase(Get_SqlStr("SELECT distinct t2.customerpn FROM erpdata..tblPackMainInfSub t1  inner join [erpdata].[dbo].tblTSVworkorder t2 on t1.大工单 = t2.ORDERNAME where t1.箱号 = '" & strReelID & "' "))

End Function

Private Sub CreateTxt(filename As String, msgTxt As String, dirtemp As String)
Dim fileNameTemp As String
Dim dirNameTemp  As String
Dim fileTemp     As String
On Error GoTo hErr

dirNameTemp = dirtemp
fileNameTemp = Replace(filename, "'", "") & ".txt"
fileTemp = dirNameTemp & fileNameTemp
Open fileTemp For Output As #1
Print #1, msgTxt
Close #1
Exit Sub

hErr:
MsgBox Err.DESCRIPTION
Close #1
End Sub

Private Sub PlaySound(sFileName As String)
Dim sPath   As String
Dim sSuffix As String
sPath = "\\10.160.1.84\public\media_source\"
sSuffix = ".wav"
media.url = sPath & sFileName & sSuffix

End Sub

