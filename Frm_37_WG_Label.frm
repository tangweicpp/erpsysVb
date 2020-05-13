VERSION 5.00
Begin VB.Form Frm_37_WG_Label 
   Caption         =   "37标签外挂"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10590
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
   ScaleHeight     =   5925
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnCannel 
      Caption         =   "取消"
      Height          =   360
      Left            =   4920
      TabIndex        =   5
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton btnCommit 
      Caption         =   "确定"
      Height          =   360
      Left            =   2400
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txtDescription 
      Height          =   4095
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   8655
   End
   Begin VB.TextBox txtUrl 
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblUrl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Txt路径:"
      Height          =   195
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   660
   End
   Begin VB.Label lblTheme 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "扫入的WaferID"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1155
   End
End
Attribute VB_Name = "Frm_37_WG_Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCannel_Click()
 txtDescription.Text = ""
 txtDescription.SetFocus
End Sub

Private Sub btnCommit_Click()
 Dim txtStr              As String
 Dim WaferID             As String
 Dim Result              As String
 Dim FLAG                As String
 Dim LotNO               As String
 Dim QueryWaferID        As String
 Dim strsql              As String
 Dim rs                  As New ADODB.Recordset
 Dim WaferNum            As String
 Dim fileNameTemp        As String
 Dim msgTxtTemp          As String
 Dim dirtemp             As String
 Dim msgBoxReturn        As String
 
 txtStr = txtDescription.Text
 If txtStr = "" Then
   msgBoxReturn = MsgBox("请输入WaferID", vbOKCancel, "系统提示")
   Exit Sub
 End If
 txtStr = Replace(txtStr, vbCrLf, "','")
 
 
 Dim arr
 '处理只有一行数据截取报错异常
 If InStr(txtStr, ",") = 0 Then
    ReDim arr(1) As String
    arr(0) = txtStr
 Else
  txtStr = Mid(txtStr, 1, InStr(txtStr, ",") - 1) & "," & Right(txtStr, Len(txtStr) - InStr(txtStr, ","))
  arr = Split(Replace(txtStr, "'", "") & ",", ",")
 End If
 '初始化参数数据
 Result = ""
 QueryWaferID = ""
 WaferNum = ""
    For i = 0 To UBound(arr) - 1
       WaferID = ""
       WaferID = Replace(arr(i), Chr(10), "")
     
        '检查WaferID是否存在
       strsql = "SELECT 工单号" & _
             " FROM erpdata..tblstocknumsub " & _
             " WHERE 流程卡编号='" & WaferID & "'"
      If rs.State = adStateOpen Then rs.Close
      rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        If rs.RecordCount > 0 Then
            QueryWaferID = QueryWaferID & "'" & WaferID & "',"
            WaferNum = WaferNum & Trim$(Right(WaferID, 2)) & " "
        Else
            Result = Result & "WaferID:" & WaferID + ";" & vbCrLf
        End If
      rs.Close
    Next i
  
  
  If Result <> "" Then
     msgBoxReturn = MsgBox(Result & "系统不存在", vbOKCancel, "系统提示")
  Else
    QueryWaferID = Left(QueryWaferID, Len(QueryWaferID) - 1)
    strsql = "sELECT COUNT(a.流程卡编号) PCS,SUM(a.数量) QTY,a.工单号 LOTNO,c.MPN_DESC  " & _
     " FROM erpdata..tblstocknumsub a, " & _
   "  ERPBASE..tblmappingData b, " & _
  "   ERPBASE..tblCustomerOI c " & _
  "   Where b.SubstrateId = a.流程卡编号 " & _
   "  and convert(varchar(20),c.ID) = b.FILENAME " & _
   "   AND a.流程卡编号 in (" & QueryWaferID & ") and 合格标记 = 0 " & _
  " GROUP BY a.工单号,c.MPN_DESC  "
         
    If rs.State = adStateOpen Then rs.Close
      rs.Open strsql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        If rs.RecordCount = 1 Then
           msgTxtTemp = Trim$(rs!MPN_DESC) & "," & Trim$(rs!LotNO) & "," & Left(WaferNum, Len(WaferNum) - 1) & "," & Trim$(rs!pcs) & "," & Trim$(rs!qty) & "," & Format(DateTime.Now, "yyyy-MM-dd")
           dirtemp = txtUrl.Text
           fileNameTemp = Format(DateTime.Now, "yyyyMMddHHmmss")
           Call addLabelTxt(fileNameTemp, msgTxtTemp, dirtemp)
           txtDescription.Text = ""
           txtDescription.SetFocus
        ElseIf rs.RecordCount > 1 Then
           msgBoxReturn = MsgBox("WaferID必须为同一LOTNO", vbOKCancel, "系统提示")
        Else
           msgBoxReturn = MsgBox("系统无资料", vbOKCancel, "系统提示")
        End If
      rs.Close

  End If
End Sub

Private Sub Form_Load()
  txtDescription.Text = ""
  txtUrl.Text = "\\10.160.1.14\BarCode\37\37DIEOUT\"
  txtUrl.Enabled = False
End Sub


