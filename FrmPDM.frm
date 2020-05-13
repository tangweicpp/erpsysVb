VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmPDM 
   Caption         =   "�ͻ�PDM�ϴ�"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "ɾ���ɵ�PDM"
      Height          =   600
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "PDM_xls"
      Height          =   2295
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   9015
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   495
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   4935
      End
      Begin VB.CommandButton Command2 
         Caption         =   ".."
         Height          =   495
         Left            =   6120
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�ϴ�DB"
         Height          =   480
         Left            =   1200
         TabIndex        =   2
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "��������"
         Height          =   480
         Left            =   4080
         TabIndex        =   1
         Top             =   1560
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3000
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ����ϴ���xls��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmPDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim cmdStr As String

    cmdStr = " delete from  tblPDM "
                        
    AddSql (cmdStr)

    MsgBox "ɾ���ɹ���", vbInformation, "��ʾ"

End Sub

Private Sub Command2_Click()

    On Error Resume Next

    Dim FName

    '˧ѡ�ļ�
    CommonDialog1.Filter = "EXCEL�ļ�(*.xls)|*.xls"
    CommonDialog1.ShowOpen
    '�õ��ļ���
    FName = CommonDialog1.filename

    If FName <> "" Then
        Text2.Text = FName

    End If

End Sub

Private Sub Command3_Click()

    '�ϴ�����

    Dim source_batch_id_Temp As String

    If Text2.Text = "" Then
        MsgBox "��ѡ����ϴ����ļ�"
        Exit Sub

    End If

    Dim dirName  As String

    Dim filename As String

    'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.Open(Text2.Text)    '���ļ�

    Set xlSheet = xlBook.Worksheets("PDM")        '��sheet�еı�

    '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 61 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If

    Dim i       As Integer

    Dim j       As Integer

    Dim id      As Long

    Dim temp    As String

    Dim temp2   As String

    Dim tempVal As String

    SumCount = 0
    BCResultFlag = False

    For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
        temp = ""
        source_batch_id_Temp = ""

        For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count

            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)

            End If
        
            tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
           
            If j = 1 Then
                source_batch_id_Temp = Trim(tempVal)  'LotId

            End If
      
            temp = temp & "," & newStr("" & tempVal)

        Next j

        temp = Mid$(temp, 2, Len(temp))
  
        Call AddPDM(temp)
        SumCount = SumCount + 1

    Next i
     
    xlBook.Close      '������ʾ�Ƿ񱣴�   ����Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

    '    VBExcel.Quit

    If SumCount > 0 Then
        MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�", , "��������"
    
    Else

        If BCResultFlag = True Then
            MsgBox "�ϴ�ʧ�ܣ���ȷ�����ϸ�ʽ��", , "��������"
            Exit Sub

        End If
    
    End If

End Sub

Public Sub AddPDM(temp As String)

    Dim cmdStr As String

    cmdStr = " insert into tblPDM values(" & temp & ")"
                        
    AddSql (cmdStr)

End Sub

Private Function newStr(temp As String)

    If temp <> "" Then
        newStr = "'" & temp & "'"
    Else
        newStr = "''"

    End If

End Function

Private Sub Command5_Click()

    ExporToExcel (" select  SC3PNAME01,SC3PNAME02 ,SC2PNAME01 ,SC2PNAME02 ,SC4PNAME01 ,SC4PNAME02  ,PROBE_MAT_DESC ,PROBE_MATERIAL ,TURNKEY_MAT_DESC,TURNKEY_MATERIAL," & _
       " TURNKEY_MAT_DESC2 ,TURNKEY_MATERIAL2 ,Test_Mat_Desc,Test_MATERIAL ,FG_MAT_DESC ,FG_MATERIAL ,MICRON_MATERIAL ,CMOS_IMAGER_TYPE ,DESIGN_ID  ,RESOLUTION ,MKT_ID ,MAX_DPW," & _
       " PIXEL_SIZE ,CHROMATICITY , TEMPERATURE_SPEC,ASSY_PROCESS_ID , DARK_BOND_PAD_ASSY , ASSY_SERIAL_TYPE ,PRODUCT_FORM, OPTICAL_QUALITY , ENCODED_MARK_ID, DIE_ATTACH_METHOD   ," & _
       " PACKAGE_LID_TYPE , LID_ATTACH_METHOD ,PACKAGE_TYPE , PACKAGE_LENGTH , PACKAGE_WIDTH , LEAD_COUNT , PB_FREE_PACKAGE , ENCAP_COMPOUND_TYPE , EPOXY_TYPE  , INTERPOSER_MATERIAL  ," & _
       " TARGET_WAF_THICKNESS , WAFER_BOX_TYPE , TEST_SITE ,TST_PROCESS_ID ,ELEC_SPECIAL_TEST  , BOX_TYPE , PROTECTIVE_FILM_APLD , SHIPPING_MST_260 ,SHIPPING_MST_LEVEL ,BOX_SIZE ,   " & _
       "  PKG_LID_ADHES_TYPE , CRA, RECON_OUT_DPW, GLASS_THICKNESS, SPECIAL_REMARK_5, SPECIAL_REMARK_3, SPECIAL_REMARK_2, SPECIAL_REMARK_4, CUSTOM_PART_NO from TBLPDM ")

End Sub
