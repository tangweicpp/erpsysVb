VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_Bumping_LOT_PACKCODE_COMP 
   Caption         =   "Bumping���"
   ClientHeight    =   11985
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15975
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
   ScaleHeight     =   11985
   ScaleWidth      =   15975
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   12015
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   15975
      Begin VB.TextBox txtPackID 
         Height          =   285
         Left            =   6000
         TabIndex        =   17
         Top             =   3720
         Width           =   3135
      End
      Begin VB.CommandButton cmd 
         Caption         =   "ɾ����ż�¼"
         Height          =   720
         Index           =   3
         Left            =   6000
         TabIndex        =   15
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmd 
         Caption         =   "������ʷ���"
         Height          =   720
         Index           =   2
         Left            =   3240
         TabIndex        =   14
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmd 
         Caption         =   "�˳�"
         Height          =   720
         Index           =   1
         Left            =   8760
         TabIndex        =   13
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmd 
         Caption         =   "����"
         Height          =   720
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   4320
         Width           =   1335
      End
      Begin VB.OptionButton Opt 
         Caption         =   "������"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   5640
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton Opt 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Top             =   1440
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtScanCode 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   458
         Width           =   6495
      End
      Begin VB.Line Line1 
         X1              =   6720
         X2              =   6720
         Y1              =   4320
         Y2              =   3960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   195
         Index           =   1
         Left            =   5640
         TabIndex        =   16
         Top             =   3720
         Width           =   360
      End
      Begin MSForms.TextBox txtPackCode 
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   11
         Top             =   2280
         Width           =   3495
         VariousPropertyBits=   746604567
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "6165;661"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtLotID 
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   10
         Top             =   1755
         Width           =   3495
         VariousPropertyBits=   746604567
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "6165;661"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin WMPLibCtl.WindowsMediaPlayer media 
         Height          =   495
         Left            =   10200
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
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
         _cx             =   2566
         _cy             =   873
      End
      Begin MSForms.TextBox txtPackCode 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   2280
         Width           =   3495
         VariousPropertyBits=   746604567
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "6165;661"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblPackCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pack Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   2325
         Width           =   1095
      End
      Begin MSForms.TextBox txtLotID 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   4
         Top             =   1755
         Width           =   3495
         VariousPropertyBits=   746604567
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "6165;661"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblLoTID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Wafer No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Top             =   1800
         Width           =   1395
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ɨ���"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm_Bumping_LOT_PACKCODE_COMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_Click(Index As Integer)
Select Case Index

    Case 0
        txtScanCode.Text = ""
        txtScanCode.SetFocus
        Opt(0).Value = True
        txtLotID(0).Text = ""
        txtLotID(1).Text = ""
        txtPackCode(0).Text = ""
        txtPackCode(1).Text = ""
        
    Case 1
        Unload Me
        
    Case 2
        ExporToExcel ("select content as ���, Match_By as �ϴ�����, Match_Date as �ϴ����� from LABELMATCH01 where Flag = 'BUMPING_LOT_PACKID' order by match_date desc ")
    Case 3
        If Len(Trim(txtPackID.Text)) = 0 Then
            MsgBox "������Ҫɾ�����������", vbInformation, "��ʾ"
        Else
            AddSql ("delete from LABELMATCH01 where content = '" & UCase(Trim$(txtPackID.Text)) & "' and flag = 'BUMPING_LOT_PACKID' ")
            MsgBox "��ʷ�ȶԼ�¼�Ѿ�û�и����" & UCase(Trim$(txtPackID.Text)), vbInformation, "��ʾ"
        End If
End Select

End Sub

Private Sub Form_Activate()
txtScanCode.SetFocus
End Sub

Private Sub txtScanCode_KeyPress(KeyAscii As Integer)

    If KeyAscii <> vbKeyReturn Or Len(Trim(txtScanCode.Text)) = 0 Then Exit Sub

    If Opt(0).Value = True Then
   
        If txtLotID(0).Text = "" Then
            txtLotID(0).Text = UCase(Trim$(txtScanCode.Text))
            Play ("LOT����ɨ��,��ɨ�����")
        Else

            If Left(UCase(Trim$(txtScanCode.Text)), 1) <> "Q" Then
                MsgBox "������Ŵ���", vbInformation, "��ʾ"
                Exit Sub

            End If
            
            If Get_OracleStr("select * from LABELMATCH01 where content = '" & UCase(Trim$(txtScanCode.Text)) & "' and flag = 'BUMPING_LOT_PACKID' ") <> "" Then
                MsgBox "�������" & UCase(Trim$(txtScanCode.Text)) & "�Ѿ�ɨ���, �뵼����ʷ���,ȷ���Ƿ����쳣����;" & vbCrLf & "�������ɾ����ʷ���", vbInformation, "��ʾ"
                Exit Sub
            End If
        
            txtPackCode(0).Text = UCase(Trim$(txtScanCode.Text))
            Play ("���������ɨ��, ��ɨ������������")
            
            Opt(1).Value = True
  
        End If

    Else
        If txtLotID(1).Text = "" Then
            
            If UCase(Trim$(txtScanCode.Text)) <> txtLotID(0).Text Then
                MsgBox "������������LOT�Ų�һ��", vbInformation, "��ʾ"
                Exit Sub

            End If
            
            txtLotID(1).Text = UCase(Trim$(txtScanCode.Text))
            
            Play ("������LOT����ɨ��,��ɨ����ű�ǩ")
        Else

            If Left(UCase(Trim$(txtScanCode.Text)), 1) <> "Q" Then
                MsgBox "��������ű�ǩ����", vbInformation, "��ʾ"
                Exit Sub

            End If
        
            If UCase(Trim$(txtScanCode.Text)) <> txtPackCode(0).Text Then
                MsgBox "��������������Ų�һ��", vbInformation, "��ʾ"
                Exit Sub

            End If
            
            txtPackCode(1).Text = UCase(Trim$(txtScanCode.Text))
            
            Play ("�������������ǩһ��, ��ɨ����������")
            
            AddSql ("insert into LABELMATCH01(content, Flag,MAtch_By, match_date) values('" & Trim(txtPackCode(0).Text) & "','BUMPING_LOT_PACKID','" & gUserName & "', sysdate)")
            Call cmd_Click(0)
        End If

    End If

    txtScanCode.Text = ""

End Sub

Rem: ������Ƶ����
Private Sub Play(sFileName As String)

    Dim sPath   As String

    Dim sSuffix As String

    sPath = "\\10.160.1.84\public\media_source\"
    sSuffix = ".wav"
    media.url = sPath & sFileName & sSuffix
    
End Sub
