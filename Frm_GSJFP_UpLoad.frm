VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_GSJFP_UpLoad 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "��˰�ַ�Ʊ�ϴ�"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra 
      Height          =   1455
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton cmdExcelIn 
         Caption         =   "��˰��Ʊ�ϴ�"
         Height          =   480
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "��˰��Ʊ����"
         Height          =   480
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   1095
         Left            =   3240
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "��        ��"
         Height          =   480
         Left            =   8280
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·����"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Fra 
      Height          =   4095
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   10455
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6975
         _Version        =   524288
         _ExtentX        =   12303
         _ExtentY        =   3413
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
         MaxCols         =   5
         MaxRows         =   5
         SpreadDesigner  =   "Frm_GSJFP_UpLoad.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   10560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_GSJFP_UpLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum FpsDetail
    e_Invoice = 1       '��˰�ַ�Ʊ
    
    e_FHDH = 2          '��������
    e_Qty = 3          '����
    e_Unit = 4         '��λ
    e_Price = 5        '����
    e_JE = 6            '���
    e_BB = 7            '�ұ�
    e_HL = 8            '����
    e_SL = 9            '˰��
    e_Invoice1 = 10       '���۷�Ʊ
    
    e_MCol = 10
End Enum

Private Sub cmdExcelIn_Click()
On Error GoTo ErrHandler

Dim FName
    'ɸѡ�ļ�
    Com.Filter = "�����ļ�(*.*)|*.*|Excel�ļ�(*.xls;*.xlsx)|*.xls;*.xlsx"
    Com.ShowOpen
    '�õ��ļ���
    FName = Com.FileName
    If FName <> "" Then
       txtPath.Text = FName  '·����ʾ����
       '������д��FPS
       FileExportInFps
    End If
    
Exit Sub
ErrHandler:
    ' �û����ˡ�ȡ������ť
    Exit Sub
End Sub
Private Sub InitCtrl()
Dim i               As Integer

    'Fps��ʼ���ǥ\��
    With Fps(0)
        .ReDraw = False
        .DAutoSizeCols = DAutoSizeColsBest
        .MaxRows = 0
        .MaxCols = FpsDetail.e_MCol
        .ColsFrozen = 1
        .Row = -1
        .Col = -1
        .Lock = True
        
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ZOrder
        .ReDraw = True
    End With

End Sub

Private Sub cmdExit_Click()
    Unload Me

End Sub

Private Sub cmdSave_Click() '���ϱ���
On Error GoTo ErrHandle
Dim strSql                          As String
Dim Rs                              As New adodb.Recordset
Dim i                               As Integer
Dim strTmp(FpsDetail.e_MCol)        As String


    '�������
    If Fps(0).MaxRows <= 0 Then
        MsgBox "û��Ҫ���������", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If MsgBox("�Ƿ�Ҫ������", vbInformation + vbYesNo, "��ʾ") = vbNo Then Exit Sub
    '��������ϣ���ʼ�������ݿ�
    '��������ģʽ
    MousePointer = 11
    With Fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsDetail.e_Invoice          'Invoice Number
            strTmp(0) = Trim$(.Text)
             .Col = FpsDetail.e_Invoice1          'Invoice Number
            strTmp(9) = Trim$(.Text)
            .Col = FpsDetail.e_FHDH             '��������
            strTmp(1) = Trim$(.Text)
            .Col = FpsDetail.e_Qty              '����
            strTmp(2) = Val(Trim$(.Text))
            .Col = FpsDetail.e_Unit             '��λ
            strTmp(3) = Trim$(.Text)
            .Col = FpsDetail.e_Price            '����
            strTmp(4) = Val(.Text)
            .Col = FpsDetail.e_JE               '���
            strTmp(5) = Val(.Text)
            .Col = FpsDetail.e_BB               '�ұ�
            strTmp(6) = Trim(.Text)
            .Col = FpsDetail.e_HL               '����
            strTmp(7) = Val(.Text)
            .Col = FpsDetail.e_SL               '˰��
            strTmp(8) = Val(.Text)
            '�ж���û������-----------------------
            If strTmp(2) <= 0 Then
                MousePointer = 0
                MsgBox "��" + Trim(i) + "������Ϊ0�����ܱ��棡", vbInformation, "��ʾ"
                Exit Sub
            End If
            
            '------------------------------------------------

            '������Ų�ѯ�Ƿ��Ѿ��ϴ����������
            strSql = "Select * From erptemp..tblBB_GSJFP Where Invoice_Number='" & Trim(strTmp(0)) & "' And Send_Number='" & Trim(strTmp(1)) & "'"
            If Rs.State = adStateOpen Then Rs.Close
            Rs.open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            If Not Rs.EOF Then  '��ʾ��������
                MousePointer = 0
                MsgBox "��" + Trim(i) + "��Invoice_Number:" + Trim(strTmp(0)) + ",��������:" + Trim(strTmp(1)) + "�Ѿ��ϴ����ˣ������ظ��ϴ���", vbInformation, "��ʾ"
                Exit Sub
            End If
            Rs.Close
            
            'У����ϣ��������ݿ�
            If Val(strTmp(2)) > 0 Then
                strSql = "Insert Into erptemp..tblBB_GSJFP(Invoice_Number,Send_Number,����,��λ,����,���,�ұ�,����,˰��, Create_by,sale_Invoice_Number) " & _
                         " Values('" & strTmp(0) & "','" & strTmp(1) & "'," & strTmp(2) & ",'" & strTmp(3) & "'," & strTmp(4) & "," & strTmp(5) & ",'" & strTmp(6) & "'," & strTmp(7) & "," & strTmp(8) & ",'" & gUserName & "','" & strTmp(9) & "')"
                INIadoCon2.Execute strSql
            End If
            
        Next
    End With
    MousePointer = 0
    
    MsgBox "���ϱ���ɹ���"
    
Exit Sub
    
ErrHandle:
    MousePointer = 0
    MsgBox Err.Description, vbCritical + vbInformation, "����"
End Sub

Private Sub Form_Load()
    '��ʼ��
    InitCtrl
End Sub
'Form��С�Զ�����
Private Sub Form_Resize()
    Fra(0).Move Fra(0).Left, Fra(0).Top, Me.ScaleWidth - 200, Fra(0).Height
    Fra(1).Move Fra(0).Left, Fra(1).Top, Me.ScaleWidth - 200, Me.ScaleHeight - Fra(0).Height - 50
    Fps(0).Move Fps(0).Left, Fps(0).Top, Fra(1).Width - 300, Fra(1).Height - 300
End Sub
'��������
Private Sub FileExportInFps()
On Error GoTo ErrHandle
Dim VBExcel                         As Excel.Application
Dim xlBook                          As Excel.Workbook
Dim xlSheet                         As Excel.Worksheet
Dim strFileName                     As String
Dim i                               As Integer
Dim j                               As Integer
Dim strChar                         As String
Dim strTmp(FpsDetail.e_MCol)        As Variant
    
    MousePointer = 11
    'Fps
    Fps(0).MaxRows = 0
    '��ȡ�ļ���
    If InStrRev(Trim(txtPath.Text), "\") > 0 Then
        strFileName = Mid(Trim(txtPath.Text), InStrRev(Trim(txtPath.Text), "\") + 1)
        If InStr(strFileName, ".") > 0 Then
            strFileName = Mid(strFileName, 1, InStr(strFileName, ".") - 1)
        End If
    End If
    'Excel�ļ�����
    '1)��Excel
    Set VBExcel = CreateObject("excel.application")      '����Excle����
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.open(txtPath.Text)    '���ļ�
    Set xlSheet = xlBook.Worksheets("Sheet1")            '��sheet�еı�
    '�ж������Excel�еĺ��趨���Ƿ���ͬ
    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> FpsDetail.e_MCol Then
        MousePointer = 0
        MsgBox "Excel�е��������趨��ģ��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        GoTo ExitPro
        Exit Sub
    End If
    '����ExcelExcel
    With Fps(0)
        For i = 1 To xlSheet.Range("A1").CurrentRegion.Rows.Count         '2)�õ�Excel�����
            strTmp(0) = Trim(xlSheet.Range("A" & i).Value)
            If Len(strTmp(0)) > 0 Then
                If i <> 1 Then .MaxRows = .MaxRows + 1  '��һ�б�ʾ���⣬����������
                For j = 1 To FpsDetail.e_MCol
                    'ѭ��i,j 26(�õ�A.B.C.)
                    If j > 26 Then
                        strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
                    Else
                        strChar = Chr(96 + j)
                    End If
'                    strTmp(j) = xlSheet.Range(strChar & i).Value   '�����Σ�������д
                    If i = 1 Then '�õ���һ��
                        .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))  '��ֵ��FPS
'                        For j = 0 To UBound(strTmp) - 1
'                            .SetText j + 1, i - 1, Trim$(strTmp(j + 1))
'                        Next
                    Else
                        .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))   '��ֵ��FPS
'                        For j = 0 To UBound(strTmp) - 1
'                            .SetText j + 1, i - 1, Trim$(strTmp(j + 1))
'                        Next
                    End If
                Next
                
            End If
        Next
    End With
    MousePointer = 0  '���״̬��ԭ
    
    xlBook.Close      '������ʾ�Ƿ񱣴�
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
    VBExcel.Quit
    
Exit Sub
ExitPro:
    On Error Resume Next
    MousePointer = 0
    If Not VBExcel Is Nothing Then
        xlBook.Close
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing
        VBExcel.Quit
    End If
    Exit Sub
ErrHandle:
    GoTo ExitPro
End Sub


