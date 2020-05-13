VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_DA69_CARTON 
   Caption         =   "DA69外箱标签"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10410
   ScaleWidth      =   13920
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   13935
      Begin VB.TextBox txtCartonQty 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   2235
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FF8080&
         Caption         =   "结束出货"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   42
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox txtScan 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1755
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00FF8080&
         Caption         =   "开始出货"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   42
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   6855
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   2640
         Width           =   12255
         _Version        =   524288
         _ExtentX        =   21616
         _ExtentY        =   12091
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
         MaxCols         =   2
         MaxRows         =   0
         SpreadDesigner  =   "Frm_DA69_CARTON.frx":0000
         Appearance      =   1
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总箱数"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   2235
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "箱号扫描"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1755
         Width           =   960
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   10920
         TabIndex        =   1
         Top             =   1560
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
   End
End
Attribute VB_Name = "Frm_DA69_CARTON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Play(sFileName As String)

    Dim sPath   As String

    Dim sSuffix As String

    sPath = "\\10.160.1.84\public\media_source\"
    sSuffix = ".wav"
    media.url = sPath & sFileName & sSuffix
    
End Sub

Private Sub cmdClose_Click()

    If txtCartonQty.Text = "0" Then
        MsgBox "当前未有箱号录入", vbInformation, "提示"
        Exit Sub

    End If

    Call saveLblData

    Play ("外箱箱号录入完毕,请操作MES打印标签")

    txtScan.Visible = False

End Sub

Private Sub cmdStart_Click()
    Play ("开始录入外箱箱号")
    txtCartonQty.Text = "0"
    fpS(0).MaxRows = 0
    txtScan.Visible = True
    txtScan.SetFocus

End Sub

Private Sub Form_Load()
    InitCtrl

End Sub

Private Sub InitCtrl()

    With fpS(0)
        .Col = -1
        .Row = -1
        .Lock = True
        
        .Col = 1
        .Row = 0
        .FontSize = 10
        
        .Col = 2
        .Row = 0
        .FontSize = 10
        
        .SetText 1, 0, "第几箱"
        .SetText 2, 0, "外箱箱号"
        
        .ColWidth(1) = 10
        .ColWidth(2) = 31
        
        .Col = 3
        .Lock = False

    End With

End Sub

Private Function checkLblData(strData As String) As Boolean

    Dim i As Integer

    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 2

            If strData = .Text Then
                MsgBox "该箱号:" & strData & "已经扫描过, 请确认是否扫描出错", vbInformation, "提示"
                Exit Function

            End If

        Next

    End With

    checkLblData = True

End Function

Private Function showLblData(strData As String)

    Dim i As Integer

    With fpS(0)
        .MaxRows = .MaxRows + 1
        i = .MaxRows
        .SetText 1, i, i
        .SetText 2, i, strData
      
    End With

    txtCartonQty.Text = fpS(0).MaxRows

    Play ("箱号已扫描")

End Function

Private Sub saveLblData()

    Dim i           As Integer

    Dim strno       As String

    Dim strCartonNO As String

    Dim strTotal    As String

    Dim strSql      As String

    With fpS(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            strno = Trim$(.Text)
        
            .Col = 2
            strCartonNO = UCase(Trim$(.Text))
        
            strTotal = Trim(txtCartonQty.Text)
        
            strSql = "select * from DA69_CARTON_DATA_TBL where CARTON_NO = '" & strCartonNO & "'"

            If Get_OracleCnt(strSql) > 0 Then
                If MsgBox("存在相同箱号重复记录, 请问此次是否为更新?", vbYesNoCancel) = vbYes Then
                    AddSql ("update DA69_CARTON_DATA_TBL set NO = '" & strno & "', Total = '" & strTotal & "', CREATED_DATE = sysdate, CREATED_BY = '" & gUserName & "' where CARTON_NO = '" & strCartonNO & "'  ")
                    MsgBox "箱号记录已更新", vbInformation, "提示"
                
                End If
                
            Else
                AddSql ("insert into DA69_CARTON_DATA_TBL(CARTON_NO, NO, Total, CREATED_DATE, CREATED_BY) values('" & strCartonNO & "', '" & strno & "','" & strTotal & "', sysdate, '" & gUserName & "')")
                Play ("箱号已记录")
            
            End If

        Next

    End With

End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)

    Dim strScan As String

    If KeyAscii <> vbKeyReturn Or Len(Trim(txtScan.Text)) = 0 Then Exit Sub

    strScan = UCase$(Trim$(txtScan.Text))

    If checkLblData(strScan) = True Then
        Call showLblData(strScan)

    End If

    txtScan.Text = ""

End Sub
