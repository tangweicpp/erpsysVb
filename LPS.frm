VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form LPS 
   BackColor       =   &H00FFFFFF&
   Caption         =   "VT�ػ�����"
   ClientHeight    =   12990
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16080
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
   ScaleHeight     =   12990
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTTab0 
      Height          =   12375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   21828
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "��Ϣ�ϴ�"
      TabPicture(0)   =   "LPS.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtPath"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Fps(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cbMode"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmd(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmd(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CommonDialog1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdQuery"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdSplit"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "LPS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.CommandButton cmdSplit 
         BackColor       =   &H80000015&
         Caption         =   "����"
         Height          =   600
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H80000015&
         Caption         =   "��ѯ"
         Height          =   360
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5400
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H80000015&
         Caption         =   "�˳�"
         Height          =   360
         Index           =   2
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   990
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H80000010&
         Caption         =   "���������"
         Height          =   600
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H80000010&
         Caption         =   "�ϴ�"
         Height          =   360
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cbMode 
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "LPS.frx":0038
         Left            =   1320
         List            =   "LPS.frx":003F
         TabIndex        =   1
         Top             =   1800
         Width           =   2415
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   9015
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   2520
         Width           =   15975
         _Version        =   524288
         _ExtentX        =   28178
         _ExtentY        =   15901
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
         SpreadDesigner  =   "LPS.frx":004D
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin MSForms.TextBox txtPath 
         Height          =   315
         Left            =   5400
         TabIndex        =   9
         Top             =   1800
         Width           =   5655
         VariousPropertyBits=   746604563
         ForeColor       =   12582912
         BorderStyle     =   1
         Size            =   "9975;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ļ�·��"
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
         Index           =   1
         Left            =   4320
         TabIndex        =   8
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴ�����"
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
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   1800
         Width           =   900
      End
   End
End
Attribute VB_Name = "LPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbMode_Click()

    With Fps(0)

        Select Case cbMode.ListIndex

            Case 0  'ί�����
                        
                .SetText 1, 0, "WAFER_ID"
                .SetText 2, 0, "��Ʒ����"
                .SetText 3, 0, "����Ʒ����"
            
        End Select
        
        

    End With

End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index

        Case 0  '�ϴ�
        
            If Len(Trim(cbMode.Text)) = 0 Then
                MsgBox "��ѡ���ϴ�����", vbInformation, "��ʾ"
                Exit Sub

            End If
            
            CommonDialog1.Filter = "�����ļ�(*.*)|*.*|Excel�ļ�(*.xls;*.xlsx)|*.xls;*.xlsx"
            CommonDialog1.ShowOpen
            
            If CommonDialog1.filename = "" Then
                Exit Sub

            End If

            txtPath.Text = CommonDialog1.filename
    
            CommonDialog1.filename = ""
            
            If txtPath.Text = "" Then
                MsgBox "��ѡ��Ҫ�ϴ����ļ�", vbInformation, "��ʾ"
                Exit Sub

            End If
            
            Select Case cbMode.ListIndex

                Case 0  'ί�����
                    Call Upload_0
                    
                Case Else
                    MsgBox "ѡ�����", vbInformation, "��ʾ"

            End Select
    
        Case 1  '����
           ' SqlServer2ExporToExcel ("SELECT * FROM erptemp..tblvt_back  order by WAFER_ID ")
             Qbox_Split
            
            
        Case 2  '�˳�
            Unload Me
            
    End Select

End Sub

Private Sub Upload_0()

    On Error GoTo ErrHandle

    Dim VBExcel     As Excel.Application

    Dim xlBook      As Excel.Workbook

    Dim xlSheet     As Excel.Worksheet
    
    Dim strWaferID  As String

    Dim strGoodDies As String

    Dim strBadDies  As String
    
    Dim User As String
    
    Dim rs         As New ADODB.Recordset

    Dim strSql     As String
    
    
    User = gUserName
    
         AddSql2 ("  UPDATE erptemp..tblvt_back  SET flag = 1 WHERE  flag = 2  ")
     
          strSql = " SELECT '' AS ѡ�� ,a.WAFER_ID, replace(b.���,' ',''),a.GOOD_DIE,a.NG_DIE ,b.���� AS �����, b.���� -a.GOOD_DIE - a.NG_DIE AS �������� ,c.��� AS �ػ���ʷ,'' as �����  FROM erptemp..tblvt_back a " & _
                 "   LEFT JOIN erpdata..tblStockNumSub b ON b.���̿���� = a.WAFER_ID  AND b.�ϸ��� = 0  LEFT JOIN erpdata..tblStockNumTree c  ON c.��� = REPLACE(b.���,' ','') + '_VT'  WHERE a.flag = '0'  " & _
                 "  ORDER BY a.WAFER_ID, b.���� -a.GOOD_DIE - a.NG_DIE "


    
    Fps(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType(rs)
          MsgBox "ϵͳ�������ϴ�δ�������ݣ�����������ϴ����ݲ�����", vbInformation, "��ʾ"
        
    Else
        
    Set VBExcel = CreateObject("excel.application")
    VBExcel.Visible = False
    Set xlBook = VBExcel.Workbooks.Open(txtPath.Text)

     Set xlSheet = xlBook.Worksheets(1)
 
  
    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 3 Then
        
        MsgBox "Excel�е��������趨��ģ��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        GoTo EXITPRO
        Exit Sub

    End If
    
    Fps(0).MaxRows = 0
    
    For I = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
        strWaferID = Replace(Trim(xlSheet.Range("A" & I)), Chr(13) + Chr(10), "")
        strGoodDies = Replace(Trim(xlSheet.Range("B" & I)), Chr(13) + Chr(10), "")
        strBadDies = Replace(Trim(xlSheet.Range("C" & I)), Chr(13) + Chr(10), "")
       
        AddSql2 ("insert into erptemp..tblvt_back select MAX(id) + 1, '" & strWaferID & "','" & strGoodDies & "','" & strBadDies & "', GETDATE()  ,'" & User & "' ,'','0' from erptemp..tblvt_back")
       

    Next
    
    MsgBox "�ϴ����", vbInformation, "��ʾ"
    
    Query
   
EXITUPLOAD:

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set VBExcel = Nothing
   
    Exit Sub
EXITPRO:

    On Error Resume Next

    MousePointer = 0

    If Not VBExcel Is Nothing Then

        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set VBExcel = Nothing

    End If
 End If
    Exit Sub
ErrHandle:
    GoTo EXITPRO

End Sub

Private Sub cmdQuery_Click()

   Query
   cmdSplit.Enabled = True
    
End Sub


Private Sub Query()


    Dim rs         As New ADODB.Recordset

    Dim strSql     As String
    
        AddSql2 ("  UPDATE erptemp..tblvt_back  SET flag = 1 WHERE  flag = 2  ")

        strSql = " SELECT '' AS ѡ�� ,a.WAFER_ID, replace(b.���,' ',''),a.GOOD_DIE,a.NG_DIE ,b.���� AS �����, b.���� -a.GOOD_DIE - a.NG_DIE AS �������� ,c.��� AS �ػ���ʷ,'' as �����  FROM erptemp..tblvt_back a " & _
                 "   LEFT JOIN erpdata..tblStockNumSub b ON b.���̿���� = a.WAFER_ID  AND b.�ϸ��� = 0   LEFT JOIN erpdata..tblStockNumTree c  ON c.��� = REPLACE(b.���,' ','') + '_VT'  WHERE a.flag = '0'  " & _
                 "  ORDER BY c.��� "

    
    Fps(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType(rs)
    Else
        
        MsgBox "û����Ҫ��������", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub


Private Sub EQuery()


    Dim rs         As New ADODB.Recordset

    Dim strSql     As String

        strSql = "  SELECT '' AS ѡ�� ,a.WAFER_ID, replace(b.���,' ',''),a.GOOD_DIE,a.NG_DIE ,b.���� AS �����, b.���� -a.GOOD_DIE - a.NG_DIE AS �������� ,c.��� AS �ػ���ʷ,'' as �����   " & _
                 " FROM erptemp..tblvt_back a  LEFT JOIN erpdata..tblStockNumSub b ON b.���̿���� = a.WAFER_ID LEFT JOIN erpdata..tblStockNumTree c  ON c.��� = REPLACE(b.���,' ','') + '_VT'  WHERE a.flag = '0'  " & _
                "  Union " & _
                "  SELECT '' AS ѡ�� ,a.WAFER_ID, replace(b.���,' ',''),'','' ,b.���� AS �����, '' AS �������� ,'' AS �ػ���ʷ,replace(b.���,' ','') as �����  FROM erptemp..tblvt_back a  " & _
                "  LEFT JOIN erpdata..tblStockNumSub b ON b.���̿���� = a.WAFER_ID   WHERE a.flag = '2'  "

    
    Fps(0).MaxRows = 0

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType(rs)
    Else
        
        MsgBox "��ѯ�����ÿͻ�����", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub





Private Sub ListDataType(rs As ADODB.Recordset)

    Dim I As Long

    With Fps(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
    With Fps(0)

        For I = 1 To .MaxRows
            .Row = I
            .Col = 1
            .ColWidth(1) = 2
            .CellType = CellTypeCheckBox
            .Text = 1
        Next
        
    End With

End Sub



Private Sub Qbox_Split()

   Dim strSql As String
   Dim rs         As New ADODB.Recordset
   Dim Qbox As String
   Dim nqbox As String
   Dim qnum As String
   Dim FLAG As String
   
     Qbox = ""
     nqbox = ""
     FLAG = ""
   
     With Fps(0)

        For I = 1 To .MaxRows
            .Row = I
            .Col = 1
             FLAG = .Text
             
            .Col = 7
            If FLAG = "1" And Val(.Text) < 0 Then
            
             MsgBox "��" & I & "�п���������������������ȷ��!", vbInformation, "��ʾ"
             Exit Sub
             
            End If
            
        Next
    
    End With
    
    
     With Fps(0)

        For I = 1 To .MaxRows
            .Row = I
            .Col = 1
           
            
            If .Text = "1" Then
                
                .Col = 3
                
                If Qbox <> .Text Then
                
                Qbox = .Text
                
                .Col = 8
                If InStr(.Text, "_VT") > 0 Then
                 
                 strSql = " SELECT COUNT(*) FROM erpdata..tblStockNumTree c WHERE  c.��� LIKE '" & Trim(.Text) & "' + '%' "
                 If rs.State = adStateOpen Then rs.Close
                 rs.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
                 
                 qnum = Trim(Str(rs.Fields(0).Value))
                 
                .Col = 9
                 .Text = Qbox + "_VT" + qnum
                 nqbox = .Text
                 
                Else
                 .Col = 9
                 .Text = Qbox + "_VT"
                 nqbox = .Text
                 
                End If
                Else
                  .Col = 9
                  .Text = nqbox
                End If
            
            End If
            qnum = ""
            
        Next
    
    End With
        
      

End Sub



Private Sub cmdSplit_Click()

   Dim nqbox As String
   Dim good_die  As Long
   Dim ng_die  As Long
   Dim wafer As String
   Dim Qbox As String
   
   
 nqbox = ""
 Qbox = ""

  With Fps(0)

        For I = 1 To .MaxRows
            .Row = I
            .Col = 1
           
            
            If .Text = "1" Then
                .Col = 9
               If nqbox <> .Text Then
               
                nqbox = .Text
                  
                AddSql2 ("INSERT INTO erpdata..TBLPACKMAININF(���,�ͻ�����,����,���߱��,�ϸ���,װ����)  VALUES ('" & nqbox & "','KR001',1,'1','0','1');INSERT INTO erpdata..tblPackTreeInf(���) VALUES ('" & nqbox & "')")
                
                AddSql2 ("INSERT INTO erpdata..tblStockNumTree ( ���,���,�ϼ����,������,������) SELECT b.���,b.���,b.�ϼ����,b.������,'0' FROM erpdata..tblPackTreeInf b WHERE b.��� = '" & nqbox & "' ")
               
               End If
              
            
            End If
            
        Next
    
    End With

 
  With Fps(0)

        For I = 1 To .MaxRows
            .Row = I
            .Col = 1
           
            
            If .Text = "1" Then
                .Col = 2
                wafer = .Text
                .Col = 3
                Qbox = .Text
                .Col = 4
                good_die = Val(.Text)
                .Col = 5
                ng_die = Val(.Text)
                .Col = 9
                nqbox = .Text
                
               If good_die <> 0 Then
               
              
                  
                AddSql2 (" INSERT INTO erpdata..tblStockNumSub  SELECT '" & nqbox & "',a.���̿����,a.������,'" & good_die & "',a.�Ϻ�,a.���ϱ��,a.�ϸ���,a.������� " & _
                        "  ,a.ID,a.�ⷿ���,GETDATE(),a.�󹤵� FROM erpdata..tblStockNumSub a  WHERE a.��� = '" & Qbox & "' AND a.���̿���� = '" & wafer & "' ; " & _
                        "  UPDATE erpdata..tblStockNumSub SET ���� = ���� - " & good_die & " WHERE ��� = '" & Qbox & "' AND ���̿���� = '" & wafer & "'; " & _
                        "  UPDATE erptemp..tblvt_back  SET flag = 2 WHERE WAFER_ID = '" & wafer & "' AND flag = 0  ")
                        
               
               ElseIf good_die = 0 Then
               
                      AddSql2 ("INSERT INTO erpdata..tblStockNumSub  SELECT '" & nqbox & "',a.���̿����,a.������,'" & ng_die & "',a.�Ϻ�,a.���ϱ��,a.�ϸ���,a.������� " & _
                        " ,a.ID,a.�ⷿ���,GETDATE(),a.�󹤵� FROM erpdata..tblStockNumSub a  WHERE a.��� = '" & Qbox & "' AND a.���̿���� = '" & wafer & "' ; " & _
                        " UPDATE erpdata..tblStockNumSub SET ���� = ���� - " & ng_die & " WHERE ��� = '" & Qbox & "' AND ���̿���� = '" & wafer & "' ; " & _
                        "  UPDATE erptemp..tblvt_back  SET flag = 2 WHERE WAFER_ID = '" & wafer & "' AND flag = 0  ")
               
               End If
                

             
              
            
            End If
            
        Next
    
    End With
    
    EQuery
    cmdSplit.Enabled = False

End Sub






