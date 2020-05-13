VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Frm_MatchLabel 
   Caption         =   "标签比对系统(BUMPING)"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
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
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   10815
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   20775
      Begin VB.TextBox txtThis 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   8640
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   1215
         Left            =   15360
         TabIndex        =   12
         Top             =   2640
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "清除记录"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1920
         TabIndex        =   11
         Top             =   9120
         Width           =   2775
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "导出比对记录"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   7440
         TabIndex        =   10
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出界面"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   11400
         TabIndex        =   9
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         Height          =   6375
         Left            =   7440
         TabIndex        =   7
         Top             =   2160
         Width           =   6855
      End
      Begin VB.TextBox txtLog 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   6375
         Left            =   1920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtIPCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtOPCode 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   735
         Left            =   15360
         TabIndex        =   15
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
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
         _cx             =   1931
         _cy             =   1296
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内箱数据:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   720
         TabIndex        =   13
         Top             =   8640
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描状态:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   6360
         TabIndex        =   8
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱原始数据:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   5280
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内箱:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   600
      End
   End
End
Attribute VB_Name = "Frm_MatchLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    ExporToExcel ("select * from LABELMATCH01 order by match_date desc")
End Sub

Private Sub Command1_Click()
    txtLog.Text = ""
    txtOPCode.Enabled = True
    txtOPCode.Locked = False
    txtOPCode.SetFocus
End Sub

Private Sub Form_Activate()
    txtOPCode.SetFocus
End Sub

Private Sub txtIPCode_KeyPress(KeyAscii As Integer)

    Dim strSql As String

    If KeyAscii <> vbKeyReturn Then

        Exit Sub

    End If

    If txtIPCode.Text = "" Then

        Exit Sub

    End If

    Dim strCode As String

    strCode = "<" & Trim(txtIPCode.Text) & ">"
    
    txtThis.Text = strCode
    
    If InStr(strCode, "<999999999>") Then
        Play ("nextWaixiang")
        txtOPCode.Enabled = True
        txtOPCode.Locked = False
        txtOPCode.SetFocus
        
        txtIPCode.Enabled = False
            
        txtIPCode.Text = ""
        txtLog.Text = ""
        Text2.Text = ""

        Exit Sub

    End If

    If InStr(txtLog.Text, strCode) > 0 Then
        Play ("right")
        Text1.BackColor = vbBlue
        txtLog.Text = Replace(txtLog.Text, strCode, "", , 1)
  
        txtIPCode.Text = ""
        strSql = "insert into LABELMATCH01 values('" & strCode & "', 'Y', '" & gUserName & "',sysdate)"
        AddSql (strSql)
        
        If Replace(txtLog.Text, vbCrLf, "") = "" Then
        ' next IP
            Play ("nextNeixiang")
            
            txtIPCode.Text = ""
            txtLog.Text = Text2.Text

            Exit Sub

        End If
      
    Else
        If InStr(Text2.Text, strCode) > 0 Then
            txtIPCode.Text = ""
            Exit Sub
        End If
    
        Text1.BackColor = vbRed
        Play ("wrong")
        MsgBox "核对错误,没有该外箱标签", vbCritical, "警告"
    
        txtIPCode.Text = ""
        strSql = "insert into LABELMATCH01 values('" & strCode & "', 'N', '" & gUserName & "',sysdate)"
        AddSql (strSql)

        Exit Sub

    End If

    txtIPCode.Text = ""

End Sub

Private Sub txtOPCode_KeyPress(KeyAscii As Integer)

    If KeyAscii <> vbKeyReturn Then

        Exit Sub

    End If

    If txtOPCode.Text = "" Then

        Exit Sub

    End If

    Dim strCode As String

    strCode = "<" & Trim(txtOPCode.Text) & ">"

    If InStr(strCode, "<" & gUserName & ">") Then
        Play ("scaninbox")
        Text2.Text = txtLog.Text
        
        txtIPCode.Enabled = True
        txtIPCode.Locked = False
        txtIPCode.SetFocus
        
        txtOPCode.Enabled = False
        txtOPCode.Text = ""

        Exit Sub

    End If

    If InStr(txtLog.Text, strCode) > 0 Then
        Play ("repScan")
        MsgBox "请勿重复扫描", vbExclamation, "提示"
        txtOPCode.Text = ""

        Exit Sub

    End If

    txtLog.Text = txtLog.Text & strCode & vbCrLf

    txtOPCode.Text = ""

End Sub

Rem: 播放音频提醒
Private Sub Play(sFileName As String)

    Dim sPath   As String

    Dim sSuffix As String

    sPath = "\\10.160.1.84\public\media_source\"
    sSuffix = ".wav"
    media.url = sPath & sFileName & sSuffix
    
End Sub
