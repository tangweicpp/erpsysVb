VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_vtreport 
   Caption         =   "ί�ⱨ��"
   ClientHeight    =   9915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20070
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   20070
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "������"
      Height          =   495
      Left            =   15960
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Cob_cust 
      Height          =   300
      ItemData        =   "Frm_vtreport.frx":0000
      Left            =   1080
      List            =   "Frm_vtreport.frx":000D
      TabIndex        =   13
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "������"
      Height          =   495
      Left            =   14520
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ѯ"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.CheckBox Chk_bonded 
      Caption         =   "����˰"
      Height          =   255
      Left            =   11280
      TabIndex        =   10
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ί��δ��ȫ�ػ�"
      Height          =   375
      Left            =   9360
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton CmdOutput 
      Caption         =   "����"
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   106823681
      CurrentDate     =   43822
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   106823681
      CurrentDate     =   43822
   End
   Begin FPSpreadADO.fpSpread fpS 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   19575
      _Version        =   524288
      _ExtentX        =   34528
      _ExtentY        =   13150
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
      SpreadDesigner  =   "Frm_vtreport.frx":0020
   End
   Begin VB.Label Label5 
      Caption         =   "�ϼ�"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   9360
      Width           =   6135
   End
   Begin VB.Label Label4 
      Caption         =   "�ͻ�����"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   735
   End
   Begin MSForms.ComboBox CobPn 
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3625;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "����"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "�Ϻ�"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "����ʱ��"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "��ʼʱ��"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Frm_vtreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOutput_Click()
  CmdOutput.Caption = "������..."
  CmdOutput.Enabled = False
  
  FpsToExcel
  CmdOutput.Caption = "����"
  CmdOutput.Enabled = True
End Sub







Private Sub Command2_Click()
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer

    CobPn.Clear
    
    If SMR.State = adStateOpen Then SMR.Close

 
strSql = " SELECT distinct rtrim(t3.�Ϻ�) AS �Ϻ� ,rtrim(t1.���̿����) AS ���̿���� ,t1.����ʱ�� AS ��һ�η���ʱ��,t1.�ϸ��� AS ��������,t2.����ʱ�� AS ���һ�λػ�ʱ��,t2.�ϸ��� AS �ػ����� FROM  " & _
" (SELECT b.���̿����,sum(b.�ϸ���+b.�Ƴ̲�����+b.���ϲ�����) AS �ϸ��� ,min(a.����ʱ��) AS ����ʱ��  " & _
" FROM erpdata..tblstockdb  a  " & _
" LEFT JOIN erpdata..tblstockdbsub b ON a.�������=b.������� AND a.���=b.���  " & _
" INNER JOIN erpbase..tblstock e ON a.ԭ�ֿ�=e.�ⷿ���� " & _
" WHERE a.ԭ�ֿ�<>'72' AND a.Ŀ��ֿ�='72'  "
If Chk_bonded.Value = 1 Then
    strSql = strSql & " and  e.�ⷿ����<>'�Ǳ�˰' "
End If

strSql = strSql & " GROUP BY b.���̿���� ) t1  " & _
" LEFT JOIN  " & _
" (SELECT d.���̿����,sum(d.�ϸ���+d.�Ƴ̲�����+d.���ϲ�����)  AS �ϸ���,max(c.����ʱ��) AS ����ʱ��  " & _
" FROM erpdata..tblstockdb  c   " & _
" LEFT JOIN erpdata..tblstockdbsub d ON c.�������=d.������� AND c.���=d.���   " & _
" INNER JOIN erpbase..tblstock f ON c.Ŀ��ֿ�=f.�ⷿ���� " & _
" WHERE c.ԭ�ֿ�='72' AND c.Ŀ��ֿ�<>'72'"
If Chk_bonded.Value = 1 Then
    strSql = strSql & " and  f.�ⷿ����<>'�Ǳ�˰' "
End If

strSql = strSql & "  GROUP BY d.���̿����  ) t2 ON t1.���̿����=t2.���̿����  " & _
" LEFT JOIN erpdata ..tblPackMainInfSub t3   ON t3.���̿����=t1.���̿����  " & _
" WHERE isnull(t1.�ϸ���,0)>isnull(t2.�ϸ���,0)   " & _
" ORDER BY rtrim(t3.�Ϻ�),t1.����ʱ��"


    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If SMR.RecordCount > 0 Then
        With Fps
           .MaxRows = 0
           Set .DataSource = SMR
           
       End With

    Else
   
      With Fps
           .MaxRows = 0
          
        End With

   End If

    MsgBox "����ѯ��" & Fps.MaxRows & "����¼", vbInformation, "��ʾ"
    SMR.Close
    
    Set SMR = Nothing
End Sub

Private Sub Command3_Click()
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer

    Dim pnlist     As String
    Command3.Enabled = False
    If SMR.State = adStateOpen Then SMR.Close
    
    pnlist = ""
   getdatafromstockdb

 ' If Trim(CobPn.Text) = "" Then
    ' strSql = "SELECT distinct c.�Ϻ� FROM erpdata..tblstockdbsub a " & _
    ' " LEFT JOIN erpdata..tblstockdb b ON a.�������=b.������� AND a.���=b.��� " & _
    ' " LEFT JOIN erpdata ..tblPackMainInfSub c   ON a.���̿����=c.���̿���� " & _
    ' " INNER JOIN erpbase..tblstock e ON b.Ŀ��ֿ�=e.�ⷿ���� " & _
    ' " INNER JOIN erpbase..tblstock f ON b.ԭ�ֿ�=f.�ⷿ���� " & _
    ' " WHERE  e.�ⷿ����<>'�Ǳ�˰' AND f.�ⷿ����<>'�Ǳ�˰' " & _
    ' " AND  c.�Ϻ� IS NOT NULL AND  (b.Ŀ��ֿ�='72'  OR b.ԭ�ֿ�='72') " & _
    ' " And b.����ʱ��<'" & Format(DTP2.Value + 1, "yyyy/mm/dd") & "' and  b.����ʱ��>'" & Format(DTP1.Value, "yyyy/mm/dd") & "'"
        
        ' SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
        ' If SMR.RecordCount > 0 Then
            ' SMR.MoveFirst
            ' For i = 1 To SMR.RecordCount
            ' If pnlist = "" Then
                ' pnlist = "'" & Trim(SMR("�Ϻ�")) & "'"
            ' Else
                ' pnlist = pnlist & "," & "'" & Trim(SMR("�Ϻ�")) & "'"
            ' End If
                ' SMR.MoveNext
            ' Next
        ' End If
        ' SMR.Close
        
        ' Set SMR = Nothing
    ' Else
        ' pnlist = "'" & Trim(CobPn.Text) & "'"
    ' End If
    date_start = Format(DTP1.Value, "yyyy/mm/dd")
    date_end = Format(DTP2.Value, "yyyy/mm/dd")
    date_MID = "2020/01/01 00:00:00"
    If DateDiff("D", date_end, date_MID) > 0 Then
    'ֻ��2020/1/1֮ǰ������
        strSql = "  select �ͻ����� , �Ϻ�, qty�ⷢ, qty�ػ�, ���, ����ʱ�� , ������Ա, ������, rtrim(���̿����) AS ���̿����,ί�з���ҵ����, �������, ���, ԭ�ֿ�, Ŀ��ֿ� FROM  erpdata..zh_tblVT_modify   where  (flag=1 or flag=2 or flag=9) and ����ʱ��<'" & Format(DTP2.Value + 1, "yyyy/mm/dd") & "' and  ����ʱ��>'" & Format(DTP1.Value, "yyyy/mm/dd") & "'"
        
        If UCase(Trim(Cob_cust.text)) = "GC" Then
            strSql = strSql & " and �ͻ�����='GC'"
       ElseIf UCase(Trim(Cob_cust.text)) = "GD108" Then
            strSql = strSql & " and  �Ϻ� Like '%GD108%'"
        ElseIf UCase(Trim(Cob_cust.text)) = "KR" Then
            strSql = strSql & " and  �Ϻ� Like '%KR%'"
        End If
        
        If Trim(CobPn.text) <> "" Then
            strSql = strSql & " AND �Ϻ� ='" & Trim(CobPn.text) & "'"
        End If
        
        strSql = strSql & "  ORDER BY �Ϻ�,rtrim(���̿����) ,����ʱ�� "
            
        
        
        
    ElseIf DateDiff("D", date_start, date_MID) > 0 Then
    '��Խ2020/1/1
         
         
        strSql = " select �ͻ����� , �Ϻ�, qty�ⷢ, qty�ػ�, ���, ����ʱ�� , ������Ա, ������, rtrim(���̿����) AS ���̿����,ί�з���ҵ����, �������, ���, ԭ�ֿ�, Ŀ��ֿ� from  erpdata..zh_tblVT_new where flag=1  and ����ʱ��>='" & date_MID & "' and  ����ʱ��<'" & Format(DTP2.Value, "yyyy/mm/dd") & "'"
        If UCase(Trim(Cob_cust.text)) = "GC" Then
            strSql = strSql & " and �ͻ�����='GC'"
        ElseIf UCase(Trim(Cob_cust.text)) = "GD108" Then
            strSql = strSql & " and  �Ϻ� Like '%GD108%'"
        ElseIf UCase(Trim(Cob_cust.text)) = "KR" Then
            strSql = strSql & " and  �Ϻ� Like '%KR%'"
        End If
        If Trim(CobPn.text) <> "" Then
            strSql = strSql & " AND �Ϻ� ='" & Trim(CobPn.text) & "'"
        End If
        
    
         
         strSql = strSql & " union select �ͻ����� , �Ϻ�, qty�ⷢ, qty�ػ�, ���, ����ʱ�� , ������Ա, ������, rtrim(���̿����) AS ���̿����,ί�з���ҵ����, �������, ���, ԭ�ֿ�, Ŀ��ֿ� FROM  erpdata..zh_tblVT_modify   where  (flag=1 or flag=2 or flag=9) and ����ʱ��<'" & date_MID & "' and  ����ʱ��>'" & Format(DTP1.Value, "yyyy/mm/dd") & "'"
        
        If UCase(Trim(Cob_cust.text)) = "GC" Then
            strSql = strSql & " and �ͻ�����='GC'"
        ElseIf UCase(Trim(Cob_cust.text)) = "GD108" Then
            strSql = strSql & " and  �Ϻ� Like '%GD108%'"
        ElseIf UCase(Trim(Cob_cust.text)) = "KR" Then
            strSql = strSql & " and  �Ϻ� Like '%KR%'"
        End If
        If Trim(CobPn.text) <> "" Then
            strSql = strSql & " AND �Ϻ� ='" & Trim(CobPn.text) & "'"
        End If
  


        strSql = strSql & "   ORDER BY �Ϻ�,rtrim(���̿����) ,����ʱ�� "
   
    
    Else
    'ֻ��2020/1/1֮�������
    
         
        strSql = " select �ͻ����� , �Ϻ�, qty�ⷢ, qty�ػ�, ���, ����ʱ�� , ������Ա, ������, rtrim(���̿����) AS ���̿����,ί�з���ҵ����, �������, ���, ԭ�ֿ�, Ŀ��ֿ�  from  erpdata..zh_tblVT_new where flag=1  and ����ʱ��>='" & date_start & "' and  ����ʱ��<'" & Format(DTP2.Value, "yyyy/mm/dd") & "'"
        If UCase(Trim(Cob_cust.text)) = "GC" Then
            strSql = strSql & " and �ͻ�����='GC'"
        ElseIf UCase(Trim(Cob_cust.text)) = "GD108" Then
            strSql = strSql & " and  �Ϻ� Like '%GD108%'"
        ElseIf UCase(Trim(Cob_cust.text)) = "KR" Then
            strSql = strSql & " and  �Ϻ� Like '%KR%'"
        End If
        If Trim(CobPn.text) <> "" Then
            strSql = strSql & " AND �Ϻ� ='" & Trim(CobPn.text) & "'"
        End If
        
        
        strSql = strSql & "  ORDER BY �Ϻ�,rtrim(���̿����) ,����ʱ�� "
        
     
    End If
    
 

  If SMR.State = adStateOpen Then SMR.Close
  SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
  If SMR.RecordCount > 0 Then
        With Fps
           .MaxRows = 0
           Set .DataSource = SMR
           
       End With

    Else
   
      With Fps
           .MaxRows = 0
'
          
        End With

   End If

 MsgBox "����ѯ��" & Fps.MaxRows & "����¼", vbInformation, "��ʾ"
   
    sumQty
    Command3.Enabled = True

End Sub

Private Sub Command4_Click()
    
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    
    Dim stritem As String
    Dim strlot As String
    Dim strWafer As String
    Dim Strqty1 As String
    Dim Strqty2 As String
    Dim strtime1 As String
    Dim strperson1 As String
    Dim strtime2 As String
    Dim strperson2 As String

    

    
    strSql = " select item,lot,wafer,cast(qty1 as int) as qty1,cast(qty2 as int) as qty2,time1,person1,time2,person2   from  erpdata..zh_tblVT_temp where isnull(item,'')<>'' and isdate(time1)=1  and isdate(time2)=1      order by wafer  "
    If SMR.State = adStateOpen Then SMR.Close

    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
             stritem = Trim(SMR("item"))
             strlot = Trim(SMR("lot"))
             strWafer = Trim(SMR("wafer"))
             Strqty1 = SMR("qty1")
             Strqty2 = SMR("qty2")
             strtime1 = Trim(SMR("time1"))
             strperson1 = Trim(SMR("person1"))
             strtime2 = Trim(SMR("time2"))
             strperson2 = Trim(SMR("person2"))
             strSql = "insert into erpdata..zh_tblVT_modify(�Ϻ�,qty�ⷢ,qty�ػ�,������Ա,����ʱ��,������,���̿����,Ŀ��ֿ�) values('" & stritem & "'," & Strqty1 & ",0,'" & strperson1 & "','" & strtime1 & "','" & strlot & "','" & strWafer & "','72')"
           
             AddSql2 (strSql)
             strSql = "insert into erpdata..zh_tblVT_modify(�Ϻ�,qty�ⷢ,qty�ػ�,������Ա,����ʱ��,������,���̿����,ԭ�ֿ�) values('" & stritem & "',0," & Strqty1 & ",'" & strperson2 & "','" & strtime2 & "','" & strlot & "','" & strWafer & "','72')"
          
             AddSql2 (strSql)
             
             
             SMR.MoveNext
        Next
    End If
    SMR.Close
    
    Set SMR = Nothing
    
    
    
    
End Sub

Private Sub Command5_Click()
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    
    Dim stritem As String
    Dim strlot As String
    Dim strWafer As String
    Dim Strqty1 As String
    Dim Strqty2 As String
    Dim strtime1 As String
    Dim strperson1 As String
    Dim strtime2 As String
    Dim strperson2 As String

    

    
    strSql = " SELECT distinct a.���̿����,c.CUSTOMERSHORTNAME,c.MPN_DESC FROM erpdata..zh_tblVT_modify  a left JOIN ERPBASE..tblmappingData b ON  rtrim(a.���̿����)=rtrim(b.SUBSTRATEID) left JOIN ERPBASE..tblCustomerOI   c ON c.SOURCE_BATCH_ID =b.LOTID and CONVERT(nvarchar(20),c.id)=b.FILENAME   "
    If SMR.State = adStateOpen Then SMR.Close

    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount

             strSql = "update  erpdata..zh_tblVT_modify set �ͻ�����='" & Trim(SMR("CUSTOMERSHORTNAME")) & "'  where  ���̿���� ='" & SMR("���̿����") & "'"
           
             AddSql2 (strSql)

             
             SMR.MoveNext
        Next
    End If
    SMR.Close
    
    Set SMR = Nothing
    

    
End Sub



Private Sub DTP1_Change()
 Call DTP2_Change
End Sub

Private Sub DTP2_Change()

    
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    
    Exit Sub
    
    CobPn.Clear
    
    If SMR.State = adStateOpen Then SMR.Close

 
strSql = "SELECT distinct c.�Ϻ� FROM erpdata..tblstockdbsub a " & _
" LEFT JOIN erpdata..tblstockdb b ON a.�������=b.������� AND a.���=b.��� " & _
" LEFT JOIN erpdata ..tblPackMainInfSub c   ON a.���̿����=c.���̿���� " & _
" INNER JOIN erpbase..tblstock e ON b.Ŀ��ֿ�=e.�ⷿ���� " & _
" INNER JOIN erpbase..tblstock f ON b.ԭ�ֿ�=f.�ⷿ���� " & _
" WHERE  e.�ⷿ����<>'�Ǳ�˰' AND f.�ⷿ����<>'�Ǳ�˰' " & _
" AND  c.�Ϻ� IS NOT NULL AND  (b.Ŀ��ֿ�='72'  OR b.ԭ�ֿ�='72') " & _
" And b.����ʱ��<'" & Format(DTP2.Value + 1, "yyyy/mm/dd") & "' and  b.����ʱ��>'" & Format(DTP1.Value, "yyyy/mm/dd") & "'"
    
    SMR.Open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If SMR.RecordCount > 0 Then
        SMR.MoveFirst
        For i = 1 To SMR.RecordCount
            CobPn.AddItem (Trim(SMR("�Ϻ�")))
            SMR.MoveNext
        Next
    End If
    SMR.Close
    
    Set SMR = Nothing
End Sub



Private Sub FpsToExcel()
    If Fps.MaxRows = 0 Then
        MsgBox "û�����ݿ��Ե���", vbInformation, "��ʾ"
        Exit Sub
    End If

    Dim i As Long
    Dim j As Long
    
    Dim xlApp      As Excel.Application
    Dim xlBook     As Excel.Workbook
    Dim xlSheet    As Excel.Worksheet
    

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

    With xlApp
        .Rows(1).Font.Bold = True
    End With
    
' On Error GoTo Ert
    With Fps
        If .MaxRows > 5000 Then MsgBox "����" & .MaxRows & "����¼��Ҫ������������Ҫ�����ӣ������ĵȴ�", vbInformation, "��ʾ"
        If .MaxCols > 10 Then
            For i = 0 To .MaxRows
                For j = 1 To .MaxCols - 4
                    .Col = j
                    .Row = i
                    xlSheet.Cells(i + 1, j) = Trim$(("'" & .text))
                        
                Next j
           
            Next i
        Else
            For i = 0 To .MaxRows
                For j = 1 To .MaxCols
                    .Col = j
                    .Row = i
                    xlSheet.Cells(i + 1, j) = Trim$(("'" & .text))
                Next j
           
            Next i
        End If
    End With

    '�����и�ʽ����
   For j = 1 To Fps.MaxCols
       If Trim(xlSheet.Cells(1, j)) = "qty�ⷢ" Or Trim(xlSheet.Cells(1, j)) = "qty�ػ�" Or Trim(xlSheet.Cells(1, j)) = "���" Or Trim(xlSheet.Cells(1, j)) = "��������" Or Trim(xlSheet.Cells(1, j)) = "�ػ�����" Then
            For i = 2 To Fps.MaxRows + 1
                xlSheet.Cells(i, j) = Replace(xlSheet.Cells(i, j), "'", "")
            Next
        End If
    Next
    With xlSheet.Range("2:" & Fps.MaxRows + 1)
        .horizontalAlignment = xlLeft
    End With
    xlSheet.Range("A1").Select
    xlApp.Columns.AutoFit
    
    xlApp.Application.Visible = True
    
    
    Set xlApp = Nothing  '"���ٱ��Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing
'Ert:

 '   If Not (xlApp Is Nothing) Then
        
 '   Set xlApp = Nothing  '"���ٱ��Excel
  '  Set xlBook = Nothing
  '  Set xlSheet = Nothing
  '  End If
    
    
End Sub



Private Sub Form_Load()
DTP1.Value = Now - 1
DTP2.Value = Now
'Call DTP2_Change
End Sub

Private Sub sumQty()
    Dim qty_out As Long
    Dim qty_back As Long
    Dim i As Long
    qty_out = 0
    qty_back = 0
    

    With Fps

            For i = 0 To .MaxRows
                .Row = i
                .Col = 3
                If IsNumeric(.text) = True Then
                    qty_out = qty_out + .text
                End If
                 .Row = i
                .Col = 4
                
                If IsNumeric(.text) = True Then
                    qty_back = qty_back + .text
                End If
                
            Next
     End With
     Label5.Caption = "�ϼƣ�ί��" & qty_out & "; �ػ� " & qty_back
            
            
End Sub

Private Sub getdatafromstockdb()

    
    Dim SMR        As New ADODB.Recordset
    Dim strSql     As String
    Dim i          As Integer
    Dim strWafer     As String
    Dim maxtime1 As String
    Dim time2 As String

strSql = " SELECT isnull(max(����ʱ��),'') FROM erpdata..zh_tblvt_new "
maxtime1 = GetSqlServerStr(strSql)

'ͬ�����µ�ί��ػ���������
 strSql = " INSERT INTO erpdata..zh_tblvt_new " & _
" SELECT  DISTINCT j.MPN_DESC as �ͻ�����, c.�Ϻ�, " & _
" CASE b.Ŀ��ֿ� WHEN '72' THEN a.�ϸ���+a.�Ƴ̲�����+a.���ϲ�����   ELSE 0 END AS 'qty(�ⷢ)'  ,CASE b.Ŀ��ֿ� WHEN '72' THEN 0   ELSE a.�ϸ���+a.�Ƴ̲�����+a.���ϲ����� END AS  'qty(�ػ�)' , " & _
" b.������Ա ,b.����ʱ��,rtrim(a.������) AS ������ ,rtrim(a.���̿����) AS ���̿���� ,'', b.�������,b.���,b.ԭ�ֿ� ,b.Ŀ��ֿ� ,j.customershortname,1,'',0 " & _
" FROM erpdata..tblstockdbsub a " & _
" LEFT JOIN erpdata..tblstockdb b ON a.�������=b.������� AND a.���=b.��� " & _
" LEFT JOIN erpdata ..tblPackMainInfSub c   ON a.���̿����=c.���̿���� " & _
" left JOIN ERPBASE..tblmappingData i ON  a.���̿����=i.SUBSTRATEID " & _
" left JOIN ERPBASE..tblCustomerOI   j ON j.SOURCE_BATCH_ID =i.LOTID and CONVERT(nvarchar(20),j.id)=i.FILENAME " & _
" WHERE isnull(c.�Ϻ�,'')<>'' and left(c.�󹤵�,1)='A' AND  ( b.ԭ�ֿ�='72' OR b.Ŀ��ֿ�='72' ) " & _
" AND b.����ʱ��>='" & maxtime1 & "' and a.������� not in (SELECT rtrim(�������) FROM erptemp..InvalidStockDb) and a.������� not in (SELECT rtrim(�����������) FROM erptemp..InvalidStockDb) " & _
" ORDER BY RTRIM(a.���̿����),b.����ʱ��"


  AddSql2 (strSql)
  'ͬ���ػ�������Ӧ���ⷢʱ��
 strSql = " UPDATE t1 SET t1.�ⷢʱ��=t2.����ʱ�� ,t1.flag =case year(t2.����ʱ��) when '2018' then 0 when '2019' then '0' else 1 end " & _
" FROM erpdata..zh_tblvt_new t1 " & _
" INNER JOIN (SELECT a.���̿����,CONVERT(varchar(100),max(c.����ʱ��) ,23) AS ����ʱ�� " & _
" FROM erpdata..zh_tblvt_new a " & _
" INNER JOIN erpdata..tblstockdbsub b ON a.���̿����=b.���̿���� " & _
" INNER JOIN erpdata..tblstockdb c ON b.�������=c.������� AND b.���=c.��� " & _
" WHERE c.Ŀ��ֿ� ='72' AND a.ԭ�ֿ�='72'   and  b.������� not in (SELECT rtrim(�������) FROM erptemp..InvalidStockDb) and b.������� not in (SELECT rtrim(�����������) FROM erptemp..InvalidStockDb )" & _
" GROUP BY a.���̿���� )t2 ON t1.���̿����=t2.���̿���� " & _
" WHERE t1.ԭ�ֿ�='72' and t1.����ʱ��>='" & maxtime1 & "' "

 AddSql2 (strSql)
 
 
  'ͬ�����
 strSql = " UPDATE t1 set t1.���=h.��˰���� " & _
" FROM erpdata..zh_tblvt_new t1 " & _
" inner JOIN ERPBASE..tblToInRec_Wafer g ON  t1.���̿����=g.��ԲID  " & _
" inner JOIN ERPBASE..TblToInSub  h ON g.��ⵥ��� =h.��ⵥ��� and g.���� =h.��������  " & _
" where h.��ⵥ��� not in ( select ������ⵥ��� from  ERPBASE..TblToInRec ) and  t1.����ʱ��>='" & maxtime1 & "'"

  AddSql2 (strSql)

'ͬ���ջ���ַ��Ӧ����ҵ����
strSql = " UPDATE t1 set t1.ί�з���ҵ����=CASE charindex('@',m.SHIP_TO_AD) WHEN 0 THEN m.SHIP_TO_AD  ELSE LEFT (m.SHIP_TO_AD,charindex('@',m.SHIP_TO_AD)-1) END   " & _
" FROM erpdata..zh_tblvt_new t1  " & _
" inner JOIN erptemp..tblstockdb_temp d ON d.remark2=t1.�������  " & _
" inner join erptemp..customer_information m on t1.�ͻ����� =m.CUSTOMER and m.SHIP_TO=isnull(d.remark1,'') " & _
" where  t1.����ʱ��>='" & maxtime1 & "'"
  AddSql2 (strSql)
  
End Sub




