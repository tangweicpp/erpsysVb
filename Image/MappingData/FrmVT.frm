VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmVT 
   Caption         =   "VT���������ϴ�"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   13410
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "ѡ����ϴ����ļ�"
      Height          =   2535
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   7095
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   495
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   4935
      End
      Begin VB.CommandButton Command6 
         Caption         =   ".."
         Height          =   495
         Left            =   6120
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "�ϴ�DB"
         Height          =   480
         Left            =   1200
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "��������"
         Height          =   480
         Left            =   3720
         TabIndex        =   1
         Top             =   1680
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   3000
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ����ϴ���xlsx��"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   1620
      End
   End
   Begin MSDataListLib.DataCombo CboCustomer 
      Height          =   330
      Left            =   1440
      TabIndex        =   7
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ���"
      Height          =   180
      Left            =   840
      TabIndex        =   6
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "FrmVT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vtDataTemp As VTData

Dim mainItemRS As New ADODB.Recordset

Private Sub Command6_Click()

On Error Resume Next
Dim FName
    '˧ѡ�ļ�
    CommonDialog2.Filter = "EXCEL�ļ�(*.xlsx)|*.xlsx"
    
    CommonDialog2.ShowOpen
    '�õ��ļ���
    FName = CommonDialog2.FileName
    If FName <> "" Then
       Text3.Text = FName
    End If



End Sub

Private Sub Command7_Click()
Dim strCust As String


If Trim(CboCustomer.Text) <> "" Then
strCust = UCase(Trim(CboCustomer.Text))
UploadVTData (strCust)

Else
    MsgBox "����ѡ��ͻ�����"
    Exit Sub

End If






End Sub



Private Sub UploadVTData(customerTemp As String)

'�ϴ�����

Dim source_batch_id_Temp As String
'�ϴ�OI��CSV
'�����ļ���
If Text3.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'��ȡ�ļ���
'    If InStrRev(Trim(Text2.Text), "\") > 0 Then
'        strFileName = Mid(Trim(Text2.Text), InStrRev(Trim(Text2.Text), "\") + 1)
'        dirName = Mid$(Trim(Text2.Text), 1, InStrRev(Trim(Text2.Text), "\"))
'    End If
    

'2012-06-27 jiayunzhang �޸Ķ�Excel�ķ�ʽ


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�

    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�

    '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 10 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If


Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
   


SumCount = 0
BCResultFlag = False

 vtDataTemp.Created_ByTemp = gUserName

 For i = 5 To xlSheet.Range("A1").CurrentRegion.Rows.Count
 
    temp = ""
    source_batch_id_Temp = ""
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
        
        If j = 1 Then
        
            vtDataTemp.SHIPDATETemp = Trim(tempVal)
            
        ElseIf j = 2 Then
            vtDataTemp.DeliveryNoTemp = Trim(tempVal)
            
       ElseIf j = 3 Then
            vtDataTemp.CustDeviceTemp = Trim(tempVal)
            
       ElseIf j = 4 Then
            vtDataTemp.CUSTLOTTemp = Trim(tempVal)
            
       ElseIf j = 5 Then
            vtDataTemp.goodDieQtyTemp = Trim(tempVal)
            
       ElseIf j = 6 Then
            vtDataTemp.ngDieQtyTemp = Trim(tempVal)
            
       ElseIf j = 7 Then
            vtDataTemp.TTLTemp = Trim(tempVal)
            
       ElseIf j = 8 Then
       
            vtDataTemp.NetWeightTemp = Trim(tempVal)
            
       ElseIf j = 9 Then
            
            vtDataTemp.GrossWeightTemp = Trim(tempVal)
            
       ElseIf j = 10 Then
            vtDataTemp.remarkTemp = Trim(tempVal)
            
        End If
        

    Next j

  

    '�ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
    If (JudgeFlagVTData(vtDataTemp.DeliveryNoTemp, vtDataTemp.CUSTLOTTemp)) Then
       MsgBox "����Ѵ��ڣ������ϴ�!", vbInformation, "������ʾ"
       GoTo NextRecord2

    End If


    Call AddVTCustomer(vtDataTemp, customerTemp)
    SumCount = SumCount + 1

    '�ϴ���DB
NextRecord2:

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





''��ȡCSV
'Dim source_batch_id_Temp As String
'Dim customerTemp As String
'
'customerTemp = "GC"
'
''�ϴ�OI��CSV
''�����ļ���
'If Text3.Text = "" Then
'    MsgBox "��ѡ����ϴ����ļ�"
'    Exit Sub
'End If
'Dim dirName As String
'Dim FileName As String
'
''��ȡ�ļ���
'    If InStrRev(Trim(Text3.Text), "\") > 0 Then
'        strfilename = Mid(Trim(Text3.Text), InStrRev(Trim(Text3.Text), "\") + 1)
'        dirName = Mid$(Trim(Text3.Text), 1, InStrRev(Trim(Text3.Text), "\"))
'    End If
'
'Dim con As New ADODB.Connection
'Dim Rs As New ADODB.Recordset
'
'
'        con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
'        Rs.open "Select * From " & "[" & strfilename & "]", con, adOpenStatic, adLockReadOnly, adCmdText
'
'        Dim i As Integer
'        Dim j As Integer
'        Dim id As Long
'        Dim temp As String
'        Dim SumCount As Integer
'        Dim GCHeaderFlag As Boolean
'        SumCount = 0
'        Rs.MoveFirst
'
'        GCHeaderFlag = False
'
'        For i = 0 To Rs.RecordCount - 1
'            temp = ""
'            id = 0
'
'            vtDataTemp.SHIPDATETemp = Rs.fields(0).Value
'            vtDataTemp.StockNoTemp = Rs.fields(1).Value
'            vtDataTemp.DeliveryNoTemp = Rs.fields(2).Value
'            vtDataTemp.CustDeviceTemp = Rs.fields(3).Value
'            vtDataTemp.CUSTLOTTemp = Rs.fields(4).Value
'            vtDataTemp.WaferIdTemp = Rs.fields(5).Value
'            vtDataTemp.WLCSPDeviceTemp = Rs.fields(6).Value
'            vtDataTemp.WLCSPLOTTemp = Rs.fields(7).Value
'            vtDataTemp.goodDieQtyTemp = CLng(Rs.fields(8).Value)
'            vtDataTemp.NGDIEQTYTemp = CLng(Rs.fields(9).Value)
'            vtDataTemp.PackingLOTNoTemp = Rs.fields(10).Value
'            vtDataTemp.TTLTemp = IIf(IsNull(Rs.fields(11).Value), "", Rs.fields(11).Value)
'            vtDataTemp.WaferQtyInTemp = IIf(IsNull(Rs.fields(12).Value), "", Rs.fields(12).Value)
'            vtDataTemp.BatchTemp = Rs.fields(13).Value
'            vtDataTemp.SAPCodeTemp = Rs.fields(14).Value
'            vtDataTemp.WorkWeekTemp = IIf(IsNull(Rs.fields(15).Value), "", Rs.fields(15).Value)
'            vtDataTemp.CartonNoTemp = IIf(IsNull(Rs.fields(16).Value), "", Rs.fields(16).Value)
'            vtDataTemp.NetWeightTemp = IIf(IsNull(Rs.fields(17).Value), "", Rs.fields(17).Value)
'            vtDataTemp.GrossWeightTemp = IIf(IsNull(Rs.fields(18).Value), "", Rs.fields(18).Value)
'            vtDataTemp.RemarkTemp = IIf(IsNull(Rs.fields(19).Value), "", Rs.fields(19).Value)
'            vtDataTemp.Created_ByTemp = gUserName
'
'
'
'
'
''                '2013-12-05 jiayun add
''                '�ж�wo�Ƿ�Ϊ��
''
''                If Trim(gcHeaderTemp.WO_NO) = "" Then
''
''                    MsgBox "WO_NO�п�ֵ����ȷ�ϣ�"
''                    Exit Sub
''
''                End If
''
''                '2012-11-07 jiayun �޸�Good_Die_Qty �����г�������
''
''            gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
''
''            '2013-12-27 jiayun add
''
''            If gcDetailTemp.Good_Die_Qty <= 0 Then
''                    MsgBox "��ȷ�Ͽͻ����ֶ�Ӧ��Die���Ƿ���ά���ã�"
''                    Exit Sub
''            End If
''
''
''            '2012-11-05 jiayun �޸� GC
''
''            '�ж�lotID��Header�����Ƿ��Ѵ���
''
''            If (JudgeGCHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)) Then
''
''                If GCHeaderFlag = False Then
''        '            MsgBox "GC ��ʣ�" & gcHeaderTemp.Lot_ID & "�Ѵ��ڣ������ϴ�!"
''                End If
''
''                '2013-12-05 jiayun add ���lotid,wo_no �Ѵ��ڣ����ѯ��id
''                '��lotid�и���ʱ�����ѯ�ϴε�id
''
''                id = GetGCLotIDWOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)
''
''            Else
''            '�ϴ���Header����
''                'ȡĿǰDB����ID��
''                id = GetMaxID()
''                '2013-01-11 jiayun add �ͻ����
''
''                If id = 0 Then
''                    MsgBox "DB����ID����ʧ��1������ϵ��Ѷ��"
''                    Exit Sub
''
''                Else
''
''
''                    Call AddGCHeader(gcHeaderTemp, id, customerTemp)
''                    GCHeaderFlag = True
''
''                End If
''
''            End If
''
''
''            '�ж�lotID��Detail�����Ƿ��Ѵ���
''
''            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_ID)) Then
''               MsgBox "GC ��ʣ�" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_ID & "�Ѵ��ڣ������ϴ�!"
''
''            Else
''            '�ϴ���Detail����
''
''                   '2012-11-05 jiayun �޸� GCT
''
''
''                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
''
''
''                If id = 0 Then
''                    MsgBox "DB����ID����ʧ��2������ϵ��Ѷ��"
''                    Exit Sub
''
''                Else
''                    Call AddGCDetail(gcDetailTemp, customerTemp, id)
''                    SumCount = SumCount + 1
''
''                End If
''
''
''            End If
''
'
'            Rs.MoveNext
'
'        Next i
'
'
'        If SumCount > 0 Then
'            MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�"
'        End If


End Sub

Private Sub Command8_Click()
Dim customerStr As String


If Trim(CboCustomer.Text) = "" Then
    MsgBox "����ѡ��ͻ����ٵ�������", vbInformation, "������ʾ"
    Exit Sub
Else


customerStr = UCase(Trim(CboCustomer.Text))

 ExporToExcel ("  select id, SHIPDATE,DELIVERYNO,CUSTDEVICE,CUSTLOT,GOODDIEQTY,NGDIEQTY,TTL,NETWEIGHT,GROSSWEIGHT,REMARK" & _
               "  Flag, Created_By, created_date " & _
               " From TSV_VT_History where customershortname='" & customerStr & "' order by id  ")
 End If


End Sub

Private Sub Form_Load()
IniCustomerName

End Sub


Private Sub IniCustomerName()
Set mainItemRS = GetJDCustomerName()
Set CboCustomer.RowSource = mainItemRS
CboCustomer.ListField = mainItemRS("productname").Name
CboCustomer.BoundColumn = mainItemRS("PID").Name

End Sub



