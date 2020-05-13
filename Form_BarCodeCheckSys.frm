VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Form_BarCodeCheckSys 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "通用三条码比对系统"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15090
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
   ScaleHeight     =   9675
   ScaleWidth      =   15090
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Form_BarCodeCheckSys 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   15135
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "清空记录"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "导出记录"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         Caption         =   "铝箔袋标签"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   7560
         TabIndex        =   5
         Top             =   1665
         Width           =   2175
      End
      Begin VB.CheckBox chk 
         Caption         =   "卷盘标签"
         Enabled         =   0   'False
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
         Left            =   4140
         TabIndex        =   4
         Top             =   1665
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         Caption         =   "临时标签"
         Enabled         =   0   'False
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
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtScan 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   10815
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "开始扫码"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   6855
         Index           =   0
         Left            =   720
         TabIndex        =   6
         Top             =   2280
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
         MaxCols         =   3
         MaxRows         =   0
         SpreadDesigner  =   "Form_BarCodeCheckSys.frx":0000
         Appearance      =   1
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   615
         Left            =   13200
         TabIndex        =   7
         Top             =   1920
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
         _cy             =   1085
      End
   End
End
Attribute VB_Name = "Form_BarCodeCheckSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
    InitStatus

End Sub

Private Sub InitStatus()
    txtScan.Visible = True
    txtScan.SetFocus

    chk(0).Value = 0
    chk(1).Value = 0
    chk(2).Value = 0
    fpS(0).MaxRows = 0

End Sub

Private Sub cmdDel_Click()
If InStr("0788510354 朱丽花6219陈静14403王银苹15204刘小倩15034席江艳15725刘艳丽15952周玉燕14367", gUserName) > 0 Then
     AddSql ("delete from unique_tbl ")
    MsgBox "历史记录已经清空", vbInformation, "提示"
Else
    MsgBox "你没有删除的权限", vbInformation, "警告"
    Exit Sub
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdExport_Click()
ExporToExcel ("select * from unique_tbl order by update_time desc")
End Sub

Private Sub Form_Load()

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
        
        .Col = 3
        .Row = 0
        .FontSize = 10
        
        .SetText 1, 0, "临时标签"
        .SetText 2, 0, "卷盘标签"
        .SetText 3, 0, "铝箔袋标签"
        
        .ColWidth(1) = 31
        .ColWidth(2) = 31
        .ColWidth(3) = 31

    End With

End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)

    If KeyAscii <> vbKeyReturn Or txtScan.Text = "" Then Exit Sub

    Call CheckHandle(UCase$(Trim$(txtScan.Text)))

    txtScan.Text = ""

End Sub

Private Sub CheckHandle(strCode As String)

    ListData (strCode)

End Sub

Private Sub ListData(strCode As String)

    If InStr(strCode, "-R") = 0 Then
        MsgBox "条码不正确", vbInformation, "提示"
        Exit Sub

    End If

    If chk(0).Value = 0 Then
        fpS(0).MaxRows = 1
        
        With fpS(0)
            .SetText 1, 1, strCode
 
        End With
        
        chk(0).Value = 1
        Play ("临时标签已扫描")

    ElseIf chk(1).Value = 0 Then

        With fpS(0)

            .SetText 2, 1, strCode

        End With
        
        chk(1).Value = 1
        Play ("卷盘标签已扫描")
        
    ElseIf chk(2).Value = 0 Then

        With fpS(0)
           
            .SetText 3, 1, strCode

        End With
        
        chk(2).Value = 1
        Play ("铝箔袋标签已扫描")
        
        CheckIsSame

    Else

    End If

End Sub

Private Sub CheckIsSame()

    Dim strcarton As String

    Dim strBox    As String
    
    Dim strReel   As String

    With fpS(0)

        .Col = 1
        .Row = 1
        strcarton = Trim$(.Text)
        
        .Col = 2
        .Row = 1
        strBox = Trim$(.Text)
            
        .Col = 3
        .Row = 1
        strReel = Trim$(.Text)
            
        If strcarton <> strBox Then
            .Row = 1
            .Col = 1
            .BackColor = vbRed
            
            .Row = 1
            .Col = 2
            .BackColor = vbRed
        
            Play ("标签不一致")
            MsgBox "标签不一致,请确认是否标签异常", vbInformation, "警告"
            
            Exit Sub

        End If
            
        If strReel <> strBox Then
            .Row = 1
            .Col = 2
            .BackColor = vbRed
            
            .Row = 1
            .Col = 3
            .BackColor = vbRed
        
            Play ("标签不一致")
            MsgBox "标签不一致,请确认是否标签异常", vbInformation, "警告"
            
            Exit Sub

        End If
        
        If Get_OracleStr("select * from unique_tbl where KEY_VALUE = '" & strReel & "'") <> "" Then
            .Row = 1
            .Col = 1
            .BackColor = vbRed
            .Col = 2
            .BackColor = vbRed
            .Col = 3
            .BackColor = vbRed
            
            MsgBox "该标签条码在历史记录中已经扫描过,请勿重复扫描, 标签出错", vbInformation, "警告"
            
            Exit Sub
        End If


    End With
    
    Play ("三条码比对一致")
    Dim lId As Long
    lId = Get_OracleStr("select UNIQUE_SEQ.NEXTVAL from dual")
    AddSql ("insert into unique_tbl(KEY_ID, KEY_VALUE,UPDATE_TIME,UPDATE_BY) values('" & lId & "', '" & strReel & "', sysdate, '" & gUserName & "') ")
    
    InitStatus

End Sub

Private Sub Play(sFileName As String)

    Dim sPath   As String

    Dim sSuffix As String

    sPath = "\\10.160.1.84\public\media_source\"
    sSuffix = ".wav"
    media.url = sPath & sFileName & sSuffix
    
End Sub
