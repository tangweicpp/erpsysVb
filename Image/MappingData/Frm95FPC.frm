VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm95FPC 
   Caption         =   "Frm95FPC"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11460
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
   ScaleHeight     =   6105
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTFPC 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10610
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "95FPCί�ⶩ������"
      TabPicture(0)   =   "Frm95FPC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmd"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtText1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "95FPC�ػ����ϵ���"
      TabPicture(1)   =   "Frm95FPC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Com"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Fra(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Fra(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Frm95FPC.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox txtText1 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Frame Fra 
         Height          =   4095
         Index           =   1
         Left            =   -74880
         TabIndex        =   4
         Top             =   1800
         Width           =   10455
         Begin FPSpreadADO.fpSpread Fps 
            Height          =   1935
            Index           =   0
            Left            =   120
            TabIndex        =   5
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
            SpreadDesigner  =   "Frm95FPC.frx":0054
            TextTip         =   2
            AppearanceStyle =   0
         End
      End
      Begin VB.Frame Fra 
         Height          =   1455
         Index           =   0
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   10455
         Begin VB.CommandButton cmdExit 
            Caption         =   "��        ��"
            Height          =   480
            Left            =   8280
            TabIndex        =   10
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtPath 
            BackColor       =   &H8000000B&
            Enabled         =   0   'False
            Height          =   1095
            Left            =   3240
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   240
            Width           =   4935
         End
         Begin VB.CommandButton CmdSave 
            Caption         =   "�ػ����ϱ���"
            Height          =   480
            Left            =   480
            TabIndex        =   7
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton cmdExcelIn 
            Caption         =   "�ػ������ϴ�"
            Height          =   480
            Left            =   480
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "·����"
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   9
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.CommandButton cmd 
         Caption         =   "������������"
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   2640
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog Com 
         Left            =   -64320
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
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
         Left            =   960
         TabIndex        =   1
         Top             =   2040
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm95FPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum FpsDetail
    e_OrderName = 1     '������
    e_PONUM = 2         '�ͻ�������
    e_POXH = 3          '��Ʒ�ͺ�
    e_Lot = 5           '�ͻ�lot1
    e_Lot2 = 6          '�ͻ�lot2
    e_Qty = 11           '����
    e_Hgs = 14          '�ϸ���
    e_Bls = 15          '������
    e_Qbox = 16         '���
    
    e_MCol = 30
End Enum

Private Sub cmd_Click()
Dim strSql                          As String
Dim RS                              As New ADODB.Recordset
Dim xlApp       As New Excel.Application
Dim xlBook      As Excel.Workbook
Dim xlSheet     As Excel.Worksheet
Dim i           As Integer, j           As Integer
Dim strFileName As String
Dim g_Path As String
g_Path = "E:\EngData" '����·��


If txtText1.Text = "" Then
MsgBox ("���빤����")
End If
'��ѯ��Ҫ��������Ϣ��ϸ
strSql = "select  row_number() OVER ( ORDER BY x.�Ϻ�) ID,'' �ͻ�����,X.PRODUCT,x.�Ϻ�,x.LOT��,x.wafer��,X.QTY,x.����,x.���Ϲ�����  from " & _
            " (select C.PRODUCT,sum(ʵ����Ʒ��*���)����,b.�Ϻ�,a.������ LOT��,dbo.FPC95_PO3(a.������,'" & txtText1.Text & "') wafer��, " & _
            " (len(replace(dbo.FPC95_PO3(a.������,'" & txtText1.Text & "'),'/','--'))-len(dbo.FPC95_PO3(a.������,'" & txtText1.Text & "')))+1 QTY,a.���Ϲ����� " & _
            " from tblstockmove a ,dbo.tblSmainM2 b,tblTSVworkorder c" & _
            " Where a.���ϱ�� = b.���ϱ��" & _
            " and c.ORDERNAME=a.���Ϲ�����" & _
            " and a.���Ϲ�����='" & txtText1.Text & "'" & _
            " and a.���ݱ��  NOT like 'F%'" & _
            " group by C.PRODUCT,b.�Ϻ�,a.������,a.���Ϲ�����" & _
            " )x" & _
            " where x.����<>0"
            
 If RS.State = adStateOpen Then
 RS.Close
 End If
 RS.open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Name = "95FpcPo3"
    xlSheet.Activate
    xlApp.Visible = False
     '��һ�б���
    
    xlApp.Range("A2:P2").MergeCells() = True '�ϲ���Ԫ��
    xlSheet.Cells(2, 1) = "�� �� ί �� ��"
    xlApp.Range("A2:P2").HorizontalAlignment = xlCenter 'ֵ����
    xlApp.Range("A2:P2").Font.Size = 12
   
   
    xlApp.Range("A3:K3").MergeCells() = True '�ϲ���Ԫ��
    xlSheet.Cells(3, 1) = "TO:����Ƽ�(����)���޹�˾"
    xlApp.Range("L3:P3").MergeCells() = True '�ϲ���Ԫ��
    xlSheet.Cells(3, 12) = "From:����Ƽ�(��ɽ)���޹�˾"
    xlApp.Range("A3:P3").Font.Size = 12
    
    xlApp.Range("A4:K4").MergeCells() = True
    xlSheet.Cells(4, 1) = "TEL:02986189323"
    xlApp.Range("L4:P4").MergeCells() = True
    xlSheet.Cells(4, 12) = "TEL:0512-50353793"
    xlApp.Range("A4:P4").Font.Size = 12
    
    xlApp.Range("A5:K5").MergeCells() = True
    xlSheet.Cells(5, 1) = "FAX:02986210606"
    xlApp.Range("L5:P5").MergeCells() = True
    xlSheet.Cells(5, 12) = "FAX:0512-50353886"
    xlApp.Range("A5:P5").Font.Size = 12
    
    xlApp.Range("A6:K6").MergeCells() = True
    xlSheet.Cells(6, 1) = "ATTN:����"
    xlApp.Range("L6:P6").MergeCells() = True
    xlSheet.Cells(6, 12) = "ATTN��½����"
    xlApp.Range("A6:P6").Font.Size = 12
    
    xlApp.Range("A7:K7").MergeCells() = True
    xlSheet.Cells(7, 1) = "ó�׷�ʽ:��˰"
    xlApp.Range("L7:P7").MergeCells() = True
    xlSheet.Cells(7, 12) = "������ţ�"
    xlApp.Range("A7:P7").Font.Size = 12
    
     xlSheet.Cells(8, 1) = "���"
     xlSheet.Cells(8, 2) = "�ͻ�������"
     xlSheet.Cells(8, 3) = "�Ϻ�"
     xlSheet.Cells(8, 4) = "оƬ����"
     xlSheet.Cells(8, 5) = "����"
     xlSheet.Cells(8, 6) = "Ƭ��"
     xlSheet.Cells(8, 7) = "Ƭ��"
     xlSheet.Cells(8, 8) = "����"
     xlSheet.Cells(8, 9) = "������"
     
    xlApp.Range("A8:H9").Font.Size = 12
    
    '������ֵѭ����ֵ��EXCEL��
    If RS.RecordCount > 0 Then
    For i = 0 To RS.RecordCount - 1
      For j = 0 To RS.fields.Count - 1
       xlSheet.Cells(9 + i, j + 1) = Trim(RS.fields(j).Value)
       xlSheet.Columns(j + 1).AutoFit '����ֵ�ô�С�Զ�����
      
      Next
      RS.MoveNext
    Next
    End If
    
    xlApp.Range("A" & i + 9 & ":P" & i + 1 + 9).MergeCells() = True '�ϲ���Ԫ��
    xlSheet.Cells(i + 9, 1) = "������Ϣ��ע��"
    xlApp.Range("A" & i + 9 & ":P" & i + 1 + 9).Font.Size = 15
    
    
    xlApp.Range("A2:P" & i + 9 + 1).Borders.Weight = xlThin '���б߿�Ӵ�
     '����ļ�
    strFileName = "95FpcPo3_" & Format(Date, "YYYYMMDD") & ".xlsx"
    If Len(strFileName) = 0 Then
        Exit Sub
    Else
        If Len(Dir(g_Path & "\" & strFileName)) > 0 Then
            On Error Resume Next
            Kill g_Path & "\" & strFileName
        End If
    End If
    
    xlSheet.SaveAs g_Path & "\" & strFileName
    xlBook.Close
    
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    RS.Close
    Set RS = Nothing
    MsgBox ("���ϵ����ɹ��� E:\EngData ����鿴")
   
End Sub


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

Private Sub CmdSave_Click() '���ϱ���
On Error GoTo ErrHandle
Dim strSql                          As String
Dim RS                              As New ADODB.Recordset
Dim i                               As Integer
Dim strTmp(FpsDetail.e_MCol)        As String
Dim strWLBH                         As String
Dim strLCK                          As String
Dim strLOT                          As String
Dim strLH                           As String
Dim intBJ                           As Integer

    '�������
    If Fps(0).MaxRows <= 0 Then
        MsgBox "û��Ҫ���������", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If MsgBox("�Ƿ�Ҫ������", vbInformation + vbYesNo, "��ʾ") = vbNo Then Exit Sub
    '��������ϣ���ʼ�������ݿ�
    '��������ģʽ
    MousePointer = 11
    On Error GoTo ErrRollback
    INIadoCon2.BeginTrans
    With Fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsDetail.e_PONUM            'PO_NUM
            strTmp(0) = Trim$(.Text)
            .Col = FpsDetail.e_POXH             'PO_CPXH
            strTmp(1) = Trim$(.Text)
            .Col = FpsDetail.e_Lot              'LOT1
            strTmp(2) = Trim$(.Text)
            .Col = FpsDetail.e_Lot2             'LOT2
            strTmp(3) = Trim$(.Text)
            .Col = FpsDetail.e_Qty              '����
            strTmp(4) = Val(.Text)
            .Col = FpsDetail.e_Hgs              '�ϸ���
            strTmp(7) = Val(.Text)
            .Col = FpsDetail.e_Bls              '������
            strTmp(8) = Val(.Text)
            '�ж���������Ʒ���ǲ���Ʒ------------------------
            If strTmp(4) <= 0 Then
                INIadoCon2.RollbackTrans
                MousePointer = 0
                MsgBox "��" + i + "����װ����Ϊ0�����ܱ��棡", vbInformation, "��ʾ"
                Exit Sub
            End If
            
            intBJ = 0
            If strTmp(7) > 0 Then   '�ϸ���
                strTmp(4) = strTmp(7)
                intBJ = 0
            Else
                strTmp(4) = strTmp(8)
                intBJ = 1
            End If
            '------------------------------------------------
            .Col = FpsDetail.e_Qbox             '���
            strTmp(5) = Trim$(.Text)
            '������Ų�ѯ�Ƿ��Ѿ��ϴ����������
            strSql = "Select * From tblFPC_BackInfo Where QBox='" & Trim(strTmp(5)) & "'"
            If RS.State = adStateOpen Then RS.Close
            RS.open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            If Not RS.EOF Then  '��ʾ�������
                INIadoCon2.RollbackTrans
                MousePointer = 0
                MsgBox "��" + i + "���" + Trim(strTmp(5)) + "�Ѿ��ϴ����ˣ������ظ��ϴ���", vbInformation, "��ʾ"
                Exit Sub
            End If
            RS.Close
            .Col = FpsDetail.e_OrderName        '������
            strTmp(6) = Trim$(.Text)
            '���ݹ����Ų�ѯ���������Ϻ�
            strWLBH = ""
            strLH = ""
            strSql = "SELECT b.���ϱ��,b.�Ϻ� FROM tblTSVworkorder a INNER JOIN tblSmainM2 b ON a.PRODUCT=b.�Ϻ� WHERE a.ORDERNAME='" & Trim(strTmp(6)) & "'"
            If RS.State = adStateOpen Then RS.Close
            RS.open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
            If Not RS.EOF Then  '������
                strWLBH = Trim$("" & RS!���ϱ��)
                strLH = Trim$("" & RS!�Ϻ�)
            Else                'û�����ϣ���ʾ�˳�
                INIadoCon2.RollbackTrans
                MousePointer = 0
                MsgBox "��" + i + "������" + Trim(strTmp(6)) + "�����ϺŲ����ڣ���ȷ�ϣ�", vbInformation, "��ʾ"
                Exit Sub
            End If
            RS.Close
            'У����ϣ��������ݿ�
            '�ػ�����
            If Len(strTmp(2)) > 0 Then
                strSql = "Insert Into tblFPC_BackInfo(PO_NUM,PO_CPXH,LOTNO,Qty,QBox,OrderName,ERPNO) " & _
                         " Values('" & strTmp(0) & "','" & strTmp(1) & "','" & strTmp(2) & "'," & strTmp(4) & ",'" & strTmp(5) & "','" & strTmp(6) & "','" & gUserName & "')"
                INIadoCon2.Execute strSql
            End If
            If Len(strTmp(3)) > 0 Then
                strSql = "Insert Into tblFPC_BackInfo(PO_NUM,PO_CPXH,LOTNO,Qty,QBox,OrderName) " & _
                         " Values('" & strTmp(0) & "','" & strTmp(1) & "','" & strTmp(3) & "'," & strTmp(4) & ",'" & strTmp(5) & "','" & strTmp(6) & "')"
                INIadoCon2.Execute strSql
            End If
            '������ű�
            strLCK = strTmp(2) + "_" + Right$("00" + Trim$(i), 2)
            strLOT = strTmp(2) + "_" + strTmp(3)
            strSql = "INSERT INTO TBLPACKMAININF(���,�ͻ�����,����,���߱��,�ϸ���,װ����) " & _
                   " Values('" & strTmp(5) & "','95'," & strTmp(4) & ",1," & intBJ & ",1)"
            INIadoCon2.Execute strSql
            strSql = "INSERT INTO TBLPACKMAININFSUB(���,���̿����,������,����,�Ϻ�,���ϱ��,�ϸ���,�󹤵�) " & _
                   " Values('" & strTmp(5) & "','" & strLCK & "','" & strLOT & "'," & strTmp(4) & ",'" & strLH & "','" & strWLBH & "'," & intBJ & ",'" & strTmp(6) & "')"
            INIadoCon2.Execute strSql
            strSql = "INSERT INTO tblPackTreeInf(���) " & _
                   " Values('" & strTmp(5) & "')"
            INIadoCon2.Execute strSql
        Next
    End With
    INIadoCon2.CommitTrans
    MousePointer = 0
    
    MsgBox "���ϱ���ɹ���"
    
Exit Sub
ErrRollback:
    '���ִ�������ع�
    MousePointer = 0
    INIadoCon2.RollbackTrans
    
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
    SSTFPC.Move SSTFPC.Left, SSTFPC.Top, Me.ScaleWidth, Me.ScaleHeight
    Fra(0).Move Fra(0).Left, Fra(0).Top, Me.ScaleWidth - 200, Fra(0).Height
    Fra(1).Move Fra(0).Left, Fra(1).Top, Me.ScaleWidth - 200, Me.ScaleHeight - Fra(0).Height - 500
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

