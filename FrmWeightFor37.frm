VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmWeightFor37 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "37WaferID����"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
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
   ScaleHeight     =   7830
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra 
      Height          =   735
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton cmdExportOut 
         Caption         =   "��     ��"
         Height          =   360
         Left            =   6360
         TabIndex        =   19
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ     ��"
         Height          =   360
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "��     ��"
         Height          =   360
         Left            =   7800
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H000000FF&
         Caption         =   "��     ��"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         TabIndex        =   11
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "��     ѯ"
         Height          =   360
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdExportIn 
         Caption         =   "��     ��"
         Height          =   360
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   990
      End
      Begin MSComDlg.CommonDialog Com 
         Left            =   9480
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Fra 
      ForeColor       =   &H00FF0000&
      Height          =   7335
      Index           =   1
      Left            =   3840
      TabIndex        =   3
      Top             =   720
      Width           =   9615
      Begin VB.OptionButton Opt 
         Caption         =   "��ά����Ϣ"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton Opt 
         Caption         =   "��ά����Ϣ"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6015
         _Version        =   524288
         _ExtentX        =   10610
         _ExtentY        =   3836
         _StockProps     =   64
         EditEnterAction =   5
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
         SpreadDesigner  =   "FrmWeightFor37.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   2175
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6015
         _Version        =   524288
         _ExtentX        =   10610
         _ExtentY        =   3836
         _StockProps     =   64
         EditEnterAction =   4
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
         SpreadDesigner  =   "FrmWeightFor37.frx":050B
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "��ѯ����"
      ForeColor       =   &H00FF0000&
      Height          =   7335
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3735
      Begin VB.TextBox txtPath 
         Height          =   2490
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "˫��ѡ��Ӧ��"
         Top             =   1680
         Visible         =   0   'False
         Width           =   3315
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   15
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   250478593
         CurrentDate     =   42739
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "˫��ѡ��Ӧ��"
         Top             =   240
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   17
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   250478593
         CurrentDate     =   42739
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wafer   ID"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   750
      End
   End
End
Attribute VB_Name = "FrmWeightFor37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum FpsDetail
    e_WaferID = 1               'WaferID
    e_Weight                    '����
    e_Stand                     '��׼
    e_NUM                       '����
    e_Cust                      '�ͻ�
    e_MCol
End Enum
'�˳�
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExportIn_Click()
On Error GoTo ErrHandler

Dim FName
    'ɸѡ�ļ�
    com.Filter = "�����ļ�(*.*)|*.*|Excel�ļ�(*.xls;*.xlsx)|*.xls;*.xlsx"
    com.ShowOpen
    '�õ��ļ���
    FName = com.filename
    If FName <> "" Then
        txtPath.Text = FName
       '������д��FPS
       FileExportInFps
    End If
    
Exit Sub
ErrHandler:
    ' �û����ˡ�ȡ������ť
    Exit Sub
End Sub
'��������
Private Sub FileExportInFps()
On Error GoTo ErrHandle
Dim VBExcel                         As Excel.Application
Dim xlBook                          As Excel.Workbook
Dim xlSheet                         As Excel.Worksheet
Dim strFilename                     As String
Dim I                               As Integer
Dim J                               As Integer
Dim strChar                         As String
Dim strTmp(FpsDetail.e_MCol - 1)    As Variant
    
    MousePointer = 11
    '���FPS(0)
    Fps(0).ClearRange FpsDetail.e_WaferID, FpsDetail.e_WaferID, Fps(0).MaxCols, Fps(0).MaxRows, True
    '��ȡ�ļ���
    If InStrRev(Trim(txtPath.Text), "\") > 0 Then
        strFilename = Mid(Trim(txtPath.Text), InStrRev(Trim(txtPath.Text), "\") + 1)
        If InStr(strFilename, ".") > 0 Then
            strFilename = Mid(strFilename, 1, InStr(strFilename, ".") - 1)
        End If
    End If
    'Excel�ļ�����
    '1)��Excel
    Set VBExcel = CreateObject("excel.application")      '����Excle����
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(txtPath.Text)    '���ļ�
    Set xlSheet = xlBook.Worksheets(1)            '��sheet�еı�
    '�ж������Excel�еĺ��趨���Ƿ���ͬ
    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> FpsDetail.e_MCol - 1 Then
        MousePointer = 0
        MsgBox "Excel�е��������趨��ģ��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        GoTo EXITPRO
        Exit Sub
    End If
    '����ExcelExcel
    With Fps(0)
        For I = 1 To xlSheet.Range("A1").CurrentRegion.Rows.Count         '2)�õ�Excel�����
            strTmp(0) = Trim(xlSheet.Range("A" & I).Value)
            If Len(strTmp(0)) > 0 Then
                For J = 1 To FpsDetail.e_MCol - 1
                    'ѭ��i,j 26(�õ�A.B.C.)
                    If J > 26 Then
                        strChar = Chr(96 + Int(J / 26 - 0.001)) & IIf(J Mod 26 = 0, "Z", Chr(96 + (J Mod 26)))
                    Else
                        strChar = Chr(96 + J)
                    End If
'                    strTmp(j) = xlSheet.Range(strChar & i).Value   '�����Σ�������д
                    If I = 1 Then '�õ���һ��
'                        .SetText j, .MaxRows, Trim$(xlSheet.Range(strChar & i))  '��ֵ��FPS
'                        For j = 0 To UBound(strTmp) - 1
'                            .SetText j + 1, i - 1, Trim$(strTmp(j + 1))
'                        Next
                    Else
                        .SetText J, I - 1, Trim$(xlSheet.Range(strChar & I)) '��ֵ��FPS
'                        For j = 0 To UBound(strTmp) - 1
'                            .SetText j + 1, i - 1, Trim$(strTmp(j + 1))
'                        Next
                    End If
                Next
                
            End If
        Next
    End With
    MousePointer = 0  '���״̬��ԭ
    
    MsgBox "����ɹ���"
    
    xlBook.Close      '������ʾ�Ƿ񱣴�
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
    
Exit Sub
EXITPRO:
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
    MousePointer = 0  '���״̬��ԭ
    MsgBox "ִ��ʧ�ܣ�" + Chr(13) + "ԭ��:" + Err.Description, vbInformation, Me.Caption
    GoTo EXITPRO
End Sub

Private Sub cmdExportOut_Click()
    If Opt(0).Value = True Then
        Export 0
    Else
        Export 1
    End If
End Sub

'ˢ��
Private Sub cmdRefresh_Click()
    If Opt(0).Value = True Then
        Fps(0).ClearRange FpsDetail.e_WaferID, FpsDetail.e_WaferID, Fps(0).MaxCols, Fps(0).MaxRows, True
    Else
        Fps(1).MaxRows = 0
    End If
End Sub

Private Sub CmdSave_Click()
    'У������
    If Not CheckData Then Exit Sub
    '�������ݵ����Ͽ�
    saveData
End Sub
'��ѯ
Private Sub cmdSearch_Click()
    Call Search(IIf(Opt(0).Value = True, 0, 1))
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Fra(1).Move Fra(1).Left, Fra(1).Top, Me.ScaleWidth - Fra(0).Width - 120, Me.ScaleHeight - Fra(2).Height - 120
    Fra(2).Move 60, Fra(2).Top, Me.ScaleWidth - 120, Fra(2).Height
    Fra(0).Move 60, Fra(0).Top, Fra(0).Width, Me.ScaleHeight - Fra(2).Height - 120
    Fps(0).Move 60, Fps(0).Top, Fra(1).Width - 120, Me.ScaleHeight - Fra(2).Height - 4 * 120
    Fps(1).Move 60, Fps(0).Top, Fra(1).Width - 120, Me.ScaleHeight - Fra(2).Height - 4 * 120
    
End Sub
Private Sub Form_Load()
    '��ʼ���ؼ�
    InitCtrl
End Sub
'��ʼ���ؼ�
Private Sub InitCtrl()
Dim I                   As Integer
Dim strsql              As String
Dim rs                  As New ADODB.Recordset

    'Fps��ʼ��
    With Fps(0)
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsBest
        .MaxRows = 500
        .ColsFrozen = 1
        .ButtonDrawMode = 1
        .MaxCols = FpsDetail.e_MCol - 1
        .Row = -1
        .Col = -1
        .Lock = True
        '�趨������
        .Col = FpsDetail.e_WaferID
        .Lock = False
        .Col = FpsDetail.e_Weight
        .CellType = CellTypeNumber
        .TypeNumberDecPlaces = 4
        .Lock = False
        .Col = FpsDetail.e_Stand
        .TypeHAlign = TypeHAlignRight
        .TypeVAlign = TypeVAlignCenter
'        .CellType = CellTypeNumber
'        .TypeNumberDecPlaces = 6
        .Col = FpsDetail.e_NUM
        .CellType = CellTypeNumber
        .TypeNumberShowSep = True
        .TypeNumberSeparator = ","
        .TypeNumberDecPlaces = 0
        .Col = FpsDetail.e_Cust
        .CellType = CellTypeComboBox
        .TypeComboBoxList = "37"
'        .TypeComboBoxList = .TypeComboBoxList & "68"
'        .TypeComboBoxList = .TypeComboBoxList & "95"
        .TypeHAlign = TypeHAlignRight
        .TypeVAlign = TypeVAlignCenter
        .SetText FpsDetail.e_Cust, -1, "37"
        '�趨�ɱ༭�߿���ɫ
        .SetCellBorder FpsDetail.e_WaferID, -1, FpsDetail.e_Weight, -1, 15, vbBlue, CellBorderStyleDot
        '������ͷ
        .SetText FpsDetail.e_WaferID, 0, "Wafer ID"
        .SetText FpsDetail.e_Weight, 0, "����"
        .SetText FpsDetail.e_Stand, 0, "��׼����"
        .SetText FpsDetail.e_NUM, 0, "����"
        .SetText FpsDetail.e_Cust, 0, "�ͻ�����"
        '����Ĭ��ֵ
        .SetText FpsDetail.e_Stand, -1, "0.000106"
        
        '�趨�п�
        .ColWidth(-1) = 10
        .RowHeight(-1) = 15
'        '�趨�Ƿ�����
'        .UserColAction = UserColActionSort
'        For i = 1 To .MaxCols
'            .Col = i
'            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
'        Next
'        .ZOrder
        .ReDraw = True
    End With
    
    With Fps(1)
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsBest
        .MaxRows = 0
        .Row = -1
        .Col = -1
        .Lock = True
        .Col = FpsDetail.e_NUM
        .CellType = CellTypeNumber
        .TypeNumberDecPlaces = 0
        .TypeNumberShowSep = True
        .TypeNumberSeparator = ","
        
        .ColWidth(-1) = 10
        .RowHeight(-1) = 15
        '�趨�Ƿ�����
        .UserColAction = UserColActionSort
        For I = 1 To .MaxCols
            .Col = I
            .ColUserSortIndicator(I) = ColUserSortIndicatorAscending
        Next
        .Visible = False
        .ZOrder
        .ReDraw = True
    End With
    
    DTP(0).Value = Format(Now(), "YYYY/MM/01")
    DTP(1).Value = Format(Now(), "YYYY/MM/DD")
    
End Sub

Private Sub fps_Change(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim DblWeight       As Double
Dim DblStand        As Double
Dim DblNum          As Double

    If Index = 0 Then
        If Col <> FpsDetail.e_Weight Then Exit Sub '�༭���������¼�
        With Fps(Index)
            .Row = Row
            .Col = FpsDetail.e_Weight   '����
            DblWeight = Val(.Text)
            .Col = FpsDetail.e_Stand    '��׼����
            DblStand = IIf(Val(.Text) = 0, 1, Val(.Text))
            DblNum = DblWeight / DblStand '�õ�����
            .SetText FpsDetail.e_NUM, Row, DblNum
        End With
    End If
End Sub

'Fps�༭�¼�
Private Sub Fps_EditChange(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim DblWeight       As Double
Dim DblStand        As Double
Dim DblNum          As Double

    If Index = 0 Then
        If Col <> FpsDetail.e_Weight Then Exit Sub '�༭���������¼�
        With Fps(Index)
            .Row = Row
            .Col = FpsDetail.e_Weight   '����
            DblWeight = Val(.Text)
            .Col = FpsDetail.e_Stand    '��׼����
            DblStand = IIf(Val(.Text) = 0, 1, Val(.Text))
            DblNum = DblWeight / DblStand '�õ�����
            .SetText FpsDetail.e_NUM, Row, DblNum
        End With
    End If
End Sub

Private Sub Opt_Click(Index As Integer)
    If Index = 0 Then
        Fps(0).Visible = True
        Fps(1).Visible = False
        cmdExportIn.Enabled = True
        cmdSave.Enabled = True
    Else
        Fps(0).Visible = False
        Fps(1).Visible = True
        cmdExportIn.Enabled = False
        cmdSave.Enabled = False
    End If
End Sub

'Private Sub Fps_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Index = 1 Then
'        If KeyCode = 46 Then    '��ʾ������del��
'            If MsgBox("�Ƿ�Ҫɾ��������", vbInformation + vbDefaultButton2 + vbYesNo, "��ʾ") = vbNo Then Exit Sub
'            With Fps(0)
'                If .MaxRows <= 0 Then Exit Sub
'                Set .DataSource = Nothing
'                .DeleteRows .ActiveRow, 1
'                .MaxRows = .MaxRows - 1
'            End With
'        End If
'    End If
'End Sub

'��������
Public Sub Export(intBJ As Integer)
    If intBJ = 0 Then
        If Not ExportFpspreadToExcel(Fps(intBJ), "��ά����Ϣ", "��ά����Ϣ") Then Exit Sub
    Else
        If Not ExportFpspreadToExcel(Fps(intBJ), "��ά����Ϣ", "��ά����Ϣ") Then Exit Sub
    End If
End Sub
'У������
Private Function CheckData() As Boolean
On Error GoTo ErrHandle
Dim I               As Integer
Dim J               As Integer
Dim strsql          As String
Dim rs              As New ADODB.Recordset
Dim strTmp(4)       As String
Dim strWaferID      As String
Dim strWaferInfo    As String

    CheckData = False
    strWaferID = ""
    Screen.MousePointer = 11
    With Fps(0)
        If .MaxRows <= 0 Then Exit Function
        For I = 1 To .MaxRows
            .Row = I
            .Col = FpsDetail.e_WaferID          'wafer id
            strTmp(0) = Replace(Replace(Trim$(.Text), vbCrLf, ""), "'", "")
            
            .Col = FpsDetail.e_Weight           '����
            strTmp(1) = Val(.Text)
            .Col = FpsDetail.e_Stand            '��׼
            strTmp(2) = Val(.Text)
            .Col = FpsDetail.e_NUM              '����
            strTmp(3) = Val(.Text)
            .Col = FpsDetail.e_Cust             '�ͻ�
            strTmp(4) = Trim$(.Text)
            'Wafer id ����Ϊ��
            If Len(strTmp(0)) > 0 Then
                '��ѯ���ݿ����У��
'                strSql = "select containername from container Where containername='" & strTmp(0) & "'"
'                If rs.State = adStateOpen Then rs.Close
'                rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
'                If rs.EOF Then  '���Mesû�����ݱ�ʾ������ WaferID
'                    Screen.MousePointer = 0
'                    MsgBox "��" & i & "�е�WaferID:" & strTmp(0) & "��MESϵͳ�в����ڣ�"
'                    Exit Function
'                End If
'                rs.Close
                '�����Ƿ�ά����ȷ
                If strTmp(1) <= 0 Then          '����
                    Screen.MousePointer = 0
                    MsgBox "��" & I & "�е�WaferID:" & strTmp(1) & "û��������"
                    Exit Function
                End If
                If strTmp(2) <= 0 Then          '��׼
                    Screen.MousePointer = 0
                    MsgBox "��" & I & "�е�WaferID:" & strTmp(2) & "û�б�׼������"
                    Exit Function
                End If
                If strTmp(3) <= 0 Then          '����
                    Screen.MousePointer = 0
                    MsgBox "��" & I & "�е�WaferID:" & strTmp(3) & "û��������"
                    Exit Function
                End If
                If strTmp(4) <= 0 Then          '�ͻ�
                    Screen.MousePointer = 0
                    MsgBox "��" & I & "�е�WaferID:" & strTmp(4) & "û�пͻ���"
                    Exit Function
                End If
                '��¼���е�WaferID,�������У���Ƿ�������ݿ�
                strWaferID = strWaferID + strTmp(0) + ","
                '�ڲ�ѭ��
                For J = I + 1 To .MaxRows
                    .Row = J
                    .Col = FpsDetail.e_WaferID
                    If strTmp(0) = Replace(Replace(Trim$(.Text), vbCrLf, ""), "'", "") Then    '����ӵ�һ�е�WaferID ���������ظ�����ʾ����
                        Screen.MousePointer = 0
                        MsgBox "��" & I & "�е�WaferID:" & strTmp(0) & "�͵�" & J & "�е�WaferID��ͬ��"
                        Exit Function
                    End If
                Next J
            End If
        Next I
    End With
    '��ȡWaferID
    If Len(strWaferID) > 0 Then
        strWaferID = Mid$(strWaferID, 1, Len(strWaferID) - 1)
        '��ѯ���ݿ����У��
        strsql = "Select WaferID From Weight37 Where WaferID In('" & Replace$(strWaferID, ",", "','") & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not rs.EOF Then  '��������ݱ�ʾ�д��ڵ�WaferID
            Do While Not rs.EOF
                strWaferInfo = strWaferInfo + Trim$("" & rs!WaferID) + ","
                rs.MoveNext
            Loop
        End If
        rs.Close
        If Len(strWaferInfo) > 0 Then
            Screen.MousePointer = 0
            strWaferInfo = Mid$(strWaferInfo, 1, Len(strWaferInfo) - 1)
            MsgBox "WaferID:" & strWaferInfo & "�Ѿ��������ݿ��У������ظ��������ݣ�"
            Exit Function
        End If
    End If
    
    CheckData = True
    Screen.MousePointer = 0
    
Exit Function
ErrHandle:
    CheckData = False
    Screen.MousePointer = 0
    MsgBox "ִ��ʧ�ܣ�" + Chr(13) + "ԭ��:" + Err.Description, vbInformation, Me.Caption
End Function

'�������ݵ����Ͽ���
Public Sub saveData()

    On Error GoTo ErrHandle

    Dim I           As Integer

    Dim strsql      As String

    Dim strsql2     As String

    Dim rs          As New ADODB.Recordset

    Dim strTmp(4)   As String

    Dim bln         As Boolean

    Dim strDatecode As String

    bln = False
    Screen.MousePointer = 11

    With Fps(0)

        If .MaxRows <= 0 Then Exit Sub

        For I = 1 To .MaxRows
            .Row = I
            .Col = FpsDetail.e_WaferID          'wafer id
            strTmp(0) = Trim$(.Text)
            .Col = FpsDetail.e_Weight           '����
            strTmp(1) = Val(.Text)
            .Col = FpsDetail.e_Stand            '��׼
            strTmp(2) = Val(.Text)
            .Col = FpsDetail.e_NUM              '����
            strTmp(3) = Val(Replace$(.Text, ",", ""))
            .Col = FpsDetail.e_Cust             '�ͻ�
            strTmp(4) = Trim$(.Text)

            'Wafer id ����Ϊ��
            If Len(strTmp(0)) > 0 Then
                bln = True
                '�������Ͽ�
                strsql = "Insert Into weight37(WAFERID,WEIGHT,STANDWEIGHT,DIE,CUSTOMER) Values('" & strTmp(0) & "','" & strTmp(1) & "','" & strTmp(2) & "','" & strTmp(3) & "','" & strTmp(4) & "')"
            
                Cnn.Execute strsql
                
                Dim sOra As String
        
                sOra = "select mes_dn_pkg.MES_WEIGHT_37('" & strTmp(0) & "') from dual"
                AddSql (sOra)
        
                strDatecode = Get_OracleStr("select case when create_date >= to_date(to_char(create_date,'yyyy') || '-12-31','yyyy-mm-dd') - mod(to_char(create_date,'YYYY'),7) - 5 " & " then to_char(create_date,'yyww') else to_char(create_date +  mod(mod(to_char(create_date,'YYYY'),7) + 5,7),'yyww') end as PODATECODE from weight37 where waferid = '" & strTmp(0) & "'")
                
                ' strDateCode = Get_OracleStr("select to_char(create_date+1,'YYWW') from weight37 where WAFERID = '" & strTmp(0) & "'")
                strsql2 = "insert into erpbase..WEIGHT37(WAFERID,CREATE_DATE) values('" & strTmp(0) & "', '" & strDatecode & "') "
                AddSql2 (strsql2)
            
            End If

        Next

    End With

    If bln = True Then
        MsgBox "���ϱ���ɹ���", vbInformation, "��ʾ"

    End If
    
    Screen.MousePointer = 0
    Exit Sub
ErrHandle:
    Screen.MousePointer = 0
    MsgBox "ִ��ʧ�ܣ�" + Chr(13) + "ԭ��:" + Err.Description, vbInformation, Me.Caption

End Sub

'��ѯ����
Public Sub Search(intBJ As Integer)
On Error GoTo ErrHandle
Dim I               As Long
Dim J               As Integer
Dim rs              As New ADODB.Recordset
Dim strsql          As String

    Screen.MousePointer = 11
    Fps(intBJ).MaxRows = 0
    If intBJ = 0 Then '��ά����Ϣ
        Screen.MousePointer = 0
        MsgBox "�����ѯ��ֱ��������Ϣ�����漴�ɣ�"
        Exit Sub
    Else
        strsql = "Select * From weight37 Where Create_Date>=to_date('" & DTP(0).Value & "','YYYY/MM/DD') And  Create_Date<to_date('" & DTP(1).Value + 1 & "','YYYY/MM/DD') "
    End If
    If txt(0).Text <> "" Then
        strsql = strsql & " And WAFERID like '" & Trim(txt(0).Text) & "%'"
    End If
    
    rs.Open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        With Fps(intBJ)
            Set .DataSource = rs
            .MaxRows = rs.RecordCount
        End With
    Else
        Screen.MousePointer = 0
        MsgBox "��������Ϣ��", vbInformation, "��ʾ"
        Exit Sub
    End If
    rs.Close
    Screen.MousePointer = 0
    
Exit Sub
ErrHandle:
    Screen.MousePointer = 0
    MsgBox "ִ��ʧ�ܣ�" + Chr(13) + "ԭ��:" + Err.Description, vbInformation, Me.Caption
End Sub






