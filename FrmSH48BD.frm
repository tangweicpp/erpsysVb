VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FrmSH48BD 
   Caption         =   "SH48标签补打工具"
   ClientHeight    =   11475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11145
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
   ScaleHeight     =   11475
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   13455
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   11175
      Begin VB.OptionButton Option2 
         Caption         =   "外包补打"
         Height          =   495
         Left            =   8520
         TabIndex        =   11
         Top             =   2040
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "内包补打"
         Height          =   495
         Left            =   8520
         TabIndex        =   10
         Top             =   1560
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.TextBox txtPces 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   5400
         TabIndex        =   8
         Text            =   "3"
         Top             =   450
         Width           =   375
      End
      Begin VB.TextBox txtFailed 
         ForeColor       =   &H000000FF&
         Height          =   9855
         Left            =   4560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtSuccess 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   9975
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtReelID 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   435
         Width           =   2655
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   7800
         TabIndex        =   9
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补打份数:"
         Height          =   195
         Left            =   4560
         TabIndex        =   7
         Top             =   480
         Width           =   780
      End
      Begin VB.Label lblFailed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补打失败的标签:"
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
         Left            =   4560
         TabIndex        =   4
         Top             =   960
         Width           =   1650
      End
      Begin VB.Label lblSuccess 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补打成功的标签:"
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
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘号:"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmSH48BD"
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
Dim strSql     As String
Dim strContent As String
Dim i          As Integer
Dim iMax       As Integer
Dim strArray() As String
Dim strData    As String
Dim strDataNew As String
Dim strID      As String
Dim rs         As New ADODB.Recordset
Dim iPces      As Integer

PrintQrReelLbl = False
'检查箱号
If Option1.Value = True Then
    If PrintIBoxLbl(strReelID) = False Then
        Exit Function

    End If

ElseIf Option2.Value = True Then
    If PrintOBoxLbl(strReelID) = False Then
        Exit Function

    End If

Else
    Exit Function

End If

'获取数据,检查数据
PrintQrReelLbl = True

End Function

Private Function PrintIBoxLbl(strReelID As String) As Boolean
Dim strSql     As String
Dim strContent As String
Dim i          As Integer
Dim iMax       As Integer
Dim strArray() As String
Dim strData    As String
Dim strDataNew As String
Dim strID      As String
Dim rs         As New ADODB.Recordset
Dim iPces      As Integer

PrintIBoxLbl = False

strSql = "select top 1 Content,ID from erpdata.dbo.tblME_PrintInfo where BartenderName = 'SH48IN1.btw' and PrinterNameID = 'W_IN_2B5F_2'  and EVENT_SOURCE = 'PKG' and LABEL_ID = 'SH48IN1' and charindex('" & strReelID & "',Content) > 0 "
Set rs = Get_SqlserveRs(strSql)
If rs.RecordCount = 0 Then
    MsgBox "查询不到", vbCritical, "警告"
    Exit Function
End If

strContent = Trim("" & rs!Content)
strID = Trim("" & rs!ID)
If InStr(strContent, ";") = 0 Then
    MsgBox "查询不到该标签的打印记录,无法补打", vbCritical, "提示"
    Exit Function

End If

strArray = Split(strContent, ";")
iMax = UBound(strArray)
If InStr(strArray(5), "QR_CODE_SH48_IN") > 0 Then
    strData = Split(strArray(5), ",")(1)
Else

    For i = 0 To iMax - 1
        If InStr(strArray(i), "QR_CODE_SH48_IN") > 0 Then
            strData = Split(strArray(i), ",")(1)

        End If

    Next

End If

strDataNew = Replace$(strData, "/", "||")
strContent = Replace$(strContent, strData, strDataNew)
iPces = CInt(Trim(txtPces.Text))
strSql = "INSERT INTO erpdata..tblME_PrintInfo(PrinterNameID,BartenderName,Content,Content2,Content3,flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) SELECT a.PrinterNameID,a.BartenderName,'" & strContent & "',a.Content2,a.Content3,'0' AS flag ,GETDATE() AS create_date,a.EVENT_SOURCE AS EVENT_SOURCE,a.EVENT_ID,a.LABEL_ID,a.PRINT_QTY FROM erpdata..tblME_PrintInfo a WHERE a.ID = '" & strID & "' "
AddSql2 (strSql)

PrintIBoxLbl = True
End Function

Private Function PrintOBoxLbl(strReelID As String) As Boolean
Dim strSql     As String
Dim strContent As String
Dim i          As Integer
Dim iMax       As Integer
Dim strArray() As String
Dim strData    As String
Dim strDataNew As String
Dim strID      As String
Dim rs         As New ADODB.Recordset
Dim iPces      As Integer

PrintOBoxLbl = False

strSql = "select top 1 ID from erpdata.dbo.tblME_PrintInfo where BartenderName = 'SH48OUT1.btw' and PrinterNameID = 'ALL_OUT_2B1F_2'  and EVENT_SOURCE = 'PKG' and LABEL_ID = 'SH48OUT1' and charindex('" & strReelID & "',Content) > 0 "
Set rs = Get_SqlserveRs(strSql)
If rs.RecordCount = 0 Then
    MsgBox "查询不到", vbCritical, "警告"
    Exit Function
End If

strID = Trim("" & rs!ID)

strSql = "INSERT INTO erpdata..tblME_PrintInfo(PrinterNameID,BartenderName,Content,Content2,Content3,flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) SELECT a.PrinterNameID,a.BartenderName,a.Content ,a.Content2,a.Content3,'0' AS flag ,GETDATE() AS create_date,a.EVENT_SOURCE AS EVENT_SOURCE,a.EVENT_ID,a.LABEL_ID,a.PRINT_QTY FROM erpdata..tblME_PrintInfo a WHERE a.ID = '" & strID & "' "
AddSql2 (strSql)

PrintOBoxLbl = True
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
