VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmShipping 
   Caption         =   "�ͻ�Shipping���ϣ��ͻ�������Ϣ�ϴ�"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   13935
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   12091
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "EQ"
      TabPicture(0)   =   "FrmShipping.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SemTech"
      TabPicture(1)   =   "FrmShipping.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame1 
         Caption         =   "PickingList DN�ϴ�"
         Height          =   2535
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   7095
         Begin VB.TextBox Txtsemtech 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   840
            Width           =   4935
         End
         Begin VB.CommandButton Command3 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   10
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton Command2 
            Caption         =   "�ϴ�DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   9
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "��������"
            Height          =   480
            Left            =   3720
            TabIndex        =   8
            Top             =   1680
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
            Height          =   180
            Index           =   0
            Left            =   840
            TabIndex        =   12
            Top             =   480
            Width           =   1530
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "ѡ����ϴ����ļ�"
         Height          =   2535
         Left            =   -74520
         TabIndex        =   1
         Top             =   840
         Width           =   7095
         Begin VB.CommandButton Command8 
            Caption         =   "��������"
            Height          =   480
            Left            =   3720
            TabIndex        =   5
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
            Caption         =   "�ϴ�DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   4
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   ".."
            Height          =   495
            Left            =   6120
            TabIndex        =   3
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   840
            Width           =   4935
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
            TabIndex        =   6
            Top             =   480
            Width           =   1620
         End
      End
   End
End
Attribute VB_Name = "FrmShipping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vtDataTemp As ShippingData

Private Sub Command1_Click()
'����

  
    ExporToExcel (" select ID,Delivery ,ItemNo ,DeliveryCreationDate ,Plant    , SalesDocument  , " & _
                 "  SOItemNo , Material  ,MarketingPN ,MaterialDescription ,PlannedGIDate ,CustomerPartNumber  ,ShipToName  , " & _
                 "   ShipToCustomer ,PurchasingDocNo ,DateCodeRestrictions  ,LabelRequirement ,ReLabelInstructions ,ShipToStreet1 ,ShipToStreet2 ,ShipToStreet3 ," & _
                 " City  ,State ,PostalCode ,CountryKey , ContactName ,Phone  ,Fax   , FreightForwarder  , " & _
                 " ShippingInstruction ,AdditionalComments ,StorageLocation , BatchNumber ,Quantity ,VolumeWeight ,GrossWeight ,Netweight , " & _
                 "  UoMForWeight ,NoOfCartons ,VendorLotNumber ,ShelfLocation ,BOLOrAirwayBillNo ,ActualShippingDate ,PackagingDetails ,PackingStatus  , " & _
                 " PickingStatus , CustomerCalendar ,customershortname  ,FLAG   ,CREATEDBY ,CREATEDDATE  from CUSTOMERSHIPPINGUPTBL     order by id desc ")
    




End Sub

Private Sub Command2_Click()
'�ϴ�semtech


SumCount = 0

Dim source_batch_id_Temp As String
Dim dntemp As String
Dim dnitemTemp As String



'�����ļ���
If Txtsemtech.Text = "" Then
    MsgBox "��ѡ����ϴ����ļ�", vbInformation, "��ʾ"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'��ȡ�ļ���
    If InStrRev(Trim(Txtsemtech.Text), "\") > 0 Then
        strFileName = Mid(Trim(Txtsemtech.Text), InStrRev(Trim(Txtsemtech.Text), "\") + 1)
        dirName = Mid$(Trim(Txtsemtech.Text), 1, InStrRev(Trim(Txtsemtech.Text), "\"))
    End If
    

Dim con As New adodb.Connection
Dim Rs As New adodb.Recordset

'con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
'Rs.open "Select * From " & strfilename, con, adOpenStatic, adLockReadOnly, adCmdText



  '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Txtsemtech.Text)    '���ļ�

    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�


  '�ж������Excel�еĺ��趨���Ƿ���ͬ
  '2012-10-08 jiayunzhang �г���Ҫ������һ�� comp_code

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 46 Then

        MsgBox "Excel�е��������趨��������һ�£���ȷ��Excel�Ƿ���ȷ��", vbInformation, "��ʾ"
        Exit Sub

    End If







Dim i As Integer
Dim j As Integer
Dim id As Long
Dim temp As String
Dim temp2 As String
Dim tempVal As String
Dim WV_inspect As String
Dim Comp_codeTemp As String
Dim dn_job As String
Dim dn_job1 As String
Dim dn_job_qty  As Long
Dim dn_job_qty1  As Long


dn_job = ""
dn_job1 = ""

SumCount = 0
'Rs.MoveFirst
'For i = 0 To Rs.RecordCount - 1

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count


    temp = ""
    source_batch_id_Temp = ""
'    For j = 0 To Rs.fields.Count - 1

'2012-07-03 ��ͻ�OI����ֶΣ����ݿ����������һ�У����Գ���Ҫ���⴦�� ��������xlSheet.Range("A1").CurrentRegion.Columns.Count ��Ϊ 71
      For j = 1 To 46
      
            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)
            End If

      
'        strChar = Chr(96 + j)
        
        
   
        
        'tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
        
        
        '�滻����Ϊ�ո񣬲�Ȼ��ǩ�����mwl
        tempVal = Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "��", " ") '��ʱ����ֵ
        
        
        
        'temp = temp & "," & newStr("" & tempVal)
                 
        If j = 34 Or j = 35 Or j = 36 Or j = 37 Then
        
           If tempVal = "" Then
           temp = temp & "," & "0"
           
           Else
           
            temp = temp & "," & newStr("" & tempVal)
           
           End If
        
        ElseIf j = 19 Then
        
        temp = temp & "," & Replace(newStr("" & tempVal), ",", " ")
        
        Else
        
        
        temp = temp & "," & newStr("" & tempVal)
        
        End If
        
        
        If j = 1 Then
            dntemp = tempVal
        End If
        
         If j = 2 Then
            dnitemTemp = tempVal
        End If
        
        If j = 32 Then
        dn_job = tempVal
        End If
        
         If j = 33 Then
        dn_job_qty = tempVal
        End If
        
        
        
 
        
    
    Next j
'    If dn_job <> dn_job1 Then
    
    'ȡĿǰDB����ID��
    id = GetshippingMaxID()
    temp = id & temp
    temp2 = temp & ",'37' ,'Y','" & gUserName & "',GETDATE(),'',''"
    temp = temp & ",'37','Y','" & gUserName & "',sysdate,'',''"
    
'    Debug.Print temp

'             '�ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
    If (JudgeFlagStautsShipingUp(dntemp, dnitemTemp)) Then
       MsgBox "���DN��" & dntemp & "�Ѵ��ڣ������ϴ�!"
       GoTo NextRecord2

    End If
    
    If (JudgeFlagStautsShipingUpjob(dntemp, dn_job)) Then
    
     Call AddShippingUPDATE(dntemp, dn_job, dn_job_qty)
    
    Else
    Call AddShippingUP(temp, temp2)
    
    End If
    
'     Else
'
'     dn_job_qty = dn_job_qty + dn_job_qty1
'
'     Call AddShippingUPDATE(dntemp, dn_job, dn_job_qty)
'
'     dn_job_qty = 0
'
'     End If
     
'    dn_job1 = dn_job
'    dn_job_qty1 = dn_job_qty
    '�ϴ���DB
    
NextRecord2:
'    Rs.MoveNext

Next i


If SumCount > 0 Then
    MsgBox "�ѳɹ��ϴ�" & SumCount & "�ʣ�", vbInformation, "��ʾ"
End If





End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim FName
    '˧ѡ�ļ�
    CommonDialog1.Filter = "EXCEL�ļ�(*.xls)|*.xls"
    
    CommonDialog1.ShowOpen
    '�õ��ļ���
    FName = CommonDialog1.FileName
    If FName <> "" Then
       Txtsemtech.Text = FName
    End If
End Sub

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



UploadVTData



End Sub



Private Sub UploadVTData()

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
    

'2015-04-27 jiayunzhang  add EQ shipping request


'Excel�ļ�����

    '1)��Excel

    Set VBExcel = CreateObject("excel.application")     '����Excle����

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '���ļ�

    Set xlSheet = xlBook.Worksheets(1)        '��sheet�еı�

    '�ж������Excel�еĺ��趨���Ƿ���ͬ

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 23 Then

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

 vtDataTemp.CreatedByTemp = gUserName

 For i = 2 To xlSheet.Range("A1").CurrentRegion.Rows.Count
 
    temp = ""
    source_batch_id_Temp = ""
    For j = 1 To xlSheet.Range("A1").CurrentRegion.Columns.Count
        strChar = Chr(96 + j)
        tempVal = xlSheet.Range(strChar & i).Value   '��ʱ����ֵ
        
        If j = 1 Then
        
            vtDataTemp.notemp = Trim(tempVal)
            
        ElseIf j = 2 Then
            vtDataTemp.SubConPOTemp = Trim(tempVal)
            
       ElseIf j = 3 Then
            vtDataTemp.itemTemp = Trim(tempVal)
            
       ElseIf j = 4 Then
            vtDataTemp.QuantityTemp = Trim(tempVal)
            
       ElseIf j = 5 Then
            vtDataTemp.DEVICETemp = Trim(tempVal)
            '-------------
            
       ElseIf j = 6 Then
            vtDataTemp.SPATemp = Trim(tempVal)
            
       ElseIf j = 7 Then
            vtDataTemp.CSDTemp = Trim(tempVal)
            
       ElseIf j = 8 Then
       
            vtDataTemp.lotTemp = Trim(tempVal)
            
       ElseIf j = 9 Then
            
            vtDataTemp.DateCode1Temp = Trim(tempVal)
            
       ElseIf j = 10 Then
            vtDataTemp.DeliveryNameTemp = Trim(tempVal)
            
       ElseIf j = 11 Then
            vtDataTemp.DeliveryAddressTemp = Trim(tempVal)
            
       ElseIf j = 12 Then
            vtDataTemp.WarehouseTemp = Trim(tempVal)
            
       ElseIf j = 13 Then
            vtDataTemp.LocationTemp = Trim(tempVal)
            
       ElseIf j = 14 Then
            vtDataTemp.ModeOfDeliveryTemp = Trim(tempVal)
            
       ElseIf j = 15 Then
            vtDataTemp.dateCodeTemp = Trim(tempVal)
            
            '-------------
            
       ElseIf j = 16 Then
            vtDataTemp.soTemp = Trim(tempVal)
            
       ElseIf j = 17 Then
       
            vtDataTemp.CarrierNotesTemp = Trim(tempVal)
            
       ElseIf j = 18 Then
            
            vtDataTemp.lineTemp = Trim(tempVal)
            
       ElseIf j = 19 Then
            vtDataTemp.ScheduleLineTemp = Trim(tempVal)
            
       ElseIf j = 20 Then
            vtDataTemp.CustPNTemp = Trim(tempVal)
            
       ElseIf j = 21 Then
            vtDataTemp.CountryDistributorTemp = Trim(tempVal)
            
       ElseIf j = 22 Then
            vtDataTemp.customerTemp = Trim(tempVal)
            
       ElseIf j = 23 Then
            vtDataTemp.customerPoTemp = Trim(tempVal)
            
            
            
        End If
        

    Next j

  

'    '�ж����SubstrateId�Ƿ��Ѵ��ڣ�������ڣ����˳���ѭ����һ��
'    If (JudgeFlagVTData(vtDataTemp.DeliveryNoTemp, vtDataTemp.CUSTLOTTemp)) Then
'       MsgBox "����Ѵ��ڣ������ϴ�!", vbInformation, "������ʾ"
'       GoTo NextRecord2
'
'    End If
    vtDataTemp.idTemp = GetEQShippingMaxID()
    Call AddEQShipping(vtDataTemp)
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

 ExporToExcel ("  select  ID ,NO ,SUBCONPO ,ITEM ,QUANTITY ,DEVICE ,SPA ,CSD ,LOT  ,DATECODE1 ,DELIVERYNAME ,DELIVERYADDRESS ,WAREHOUSE  ," & _
               "   Location , MODEOFDELIVERY, DateCode, SO, CARRIERNOTES, Line, SCHEDULELINE, CUSTPN, COUNTRYDISTRIBUTOR, Customer, CUSTOMERPO ,FLAG, CREATEDBY, CREATEDDATE " & _
               " From customershippingtbl order by id  ")
               
               

 


End Sub

Private Function newStr(temp As String)
If temp <> "" Then
   If InStr(temp, "'") > 0 Then
   newStr = "'" & Replace(temp, "'", "") & "'"
   
   Else
   newStr = "'" & temp & "'"
   
   End If
   



Else


newStr = "''"

End If

End Function
