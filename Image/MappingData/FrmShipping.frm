VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmShipping 
   Caption         =   "客户Shipping资料，客户挑料信息上传"
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
         Caption         =   "PickingList DN上传"
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
            Caption         =   "上传DB"
            Height          =   480
            Left            =   1200
            TabIndex        =   9
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "导出报表"
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
            Caption         =   "选择待上传的xls："
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
         Caption         =   "选择待上传的文件"
         Height          =   2535
         Left            =   -74520
         TabIndex        =   1
         Top             =   840
         Width           =   7095
         Begin VB.CommandButton Command8 
            Caption         =   "导出报表"
            Height          =   480
            Left            =   3720
            TabIndex        =   5
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
            Caption         =   "上传DB"
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
            Caption         =   "选择待上传的xlsx："
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
'导出

  
    ExporToExcel (" select ID,Delivery ,ItemNo ,DeliveryCreationDate ,Plant    , SalesDocument  , " & _
                 "  SOItemNo , Material  ,MarketingPN ,MaterialDescription ,PlannedGIDate ,CustomerPartNumber  ,ShipToName  , " & _
                 "   ShipToCustomer ,PurchasingDocNo ,DateCodeRestrictions  ,LabelRequirement ,ReLabelInstructions ,ShipToStreet1 ,ShipToStreet2 ,ShipToStreet3 ," & _
                 " City  ,State ,PostalCode ,CountryKey , ContactName ,Phone  ,Fax   , FreightForwarder  , " & _
                 " ShippingInstruction ,AdditionalComments ,StorageLocation , BatchNumber ,Quantity ,VolumeWeight ,GrossWeight ,Netweight , " & _
                 "  UoMForWeight ,NoOfCartons ,VendorLotNumber ,ShelfLocation ,BOLOrAirwayBillNo ,ActualShippingDate ,PackagingDetails ,PackingStatus  , " & _
                 " PickingStatus , CustomerCalendar ,customershortname  ,FLAG   ,CREATEDBY ,CREATEDDATE  from CUSTOMERSHIPPINGUPTBL     order by id desc ")
    




End Sub

Private Sub Command2_Click()
'上传semtech


SumCount = 0

Dim source_batch_id_Temp As String
Dim dntemp As String
Dim dnitemTemp As String



'处理文件名
If Txtsemtech.Text = "" Then
    MsgBox "先选择待上传的文件", vbInformation, "提示"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'获取文件名
    If InStrRev(Trim(Txtsemtech.Text), "\") > 0 Then
        strFileName = Mid(Trim(Txtsemtech.Text), InStrRev(Trim(Txtsemtech.Text), "\") + 1)
        dirName = Mid$(Trim(Txtsemtech.Text), 1, InStrRev(Trim(Txtsemtech.Text), "\"))
    End If
    

Dim con As New adodb.Connection
Dim Rs As New adodb.Recordset

'con.open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & dirName & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
'Rs.open "Select * From " & strfilename, con, adOpenStatic, adLockReadOnly, adCmdText



  '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Txtsemtech.Text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表


  '判定最大列Excel中的和设定列是否相同
  '2012-10-08 jiayunzhang 市场部要求新增一列 comp_code

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 46 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
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

'2012-07-03 因客户OI添加字段，数据库新增在最后一列，所以程序要特殊处理。 把列数，xlSheet.Range("A1").CurrentRegion.Columns.Count 改为 71
      For j = 1 To 46
      
            If j > 26 Then
                strChar = Chr(96 + Int(j / 26 - 0.001)) & IIf(j Mod 26 = 0, "Z", Chr(96 + (j Mod 26)))
            Else
                strChar = Chr(96 + j)
            End If

      
'        strChar = Chr(96 + j)
        
        
   
        
        'tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
        
        
        '替换逗号为空格，不然标签会出错mwl
        tempVal = Replace(Replace(xlSheet.Range(strChar & i).Value, ",", " "), "，", " ") '临时保存值
        
        
        
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
    
    '取目前DB最大的ID号
    id = GetshippingMaxID()
    temp = id & temp
    temp2 = temp & ",'37' ,'Y','" & gUserName & "',GETDATE(),'',''"
    temp = temp & ",'37','Y','" & gUserName & "',sysdate,'',''"
    
'    Debug.Print temp

'             '判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
    If (JudgeFlagStautsShipingUp(dntemp, dnitemTemp)) Then
       MsgBox "这笔DN：" & dntemp & "已存在，无需上传!"
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
    '上传到DB
    
NextRecord2:
'    Rs.MoveNext

Next i


If SumCount > 0 Then
    MsgBox "已成功上传" & SumCount & "笔！", vbInformation, "提示"
End If





End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim FName
    '帅选文件
    CommonDialog1.Filter = "EXCEL文件(*.xls)|*.xls"
    
    CommonDialog1.ShowOpen
    '得到文件名
    FName = CommonDialog1.FileName
    If FName <> "" Then
       Txtsemtech.Text = FName
    End If
End Sub

Private Sub Command6_Click()

On Error Resume Next
Dim FName
    '帅选文件
    CommonDialog2.Filter = "EXCEL文件(*.xlsx)|*.xlsx"
    
    CommonDialog2.ShowOpen
    '得到文件名
    FName = CommonDialog2.FileName
    If FName <> "" Then
       Text3.Text = FName
    End If



End Sub

Private Sub Command7_Click()



UploadVTData



End Sub



Private Sub UploadVTData()

'上传资料

Dim source_batch_id_Temp As String
'上传OI的CSV
'处理文件名
If Text3.Text = "" Then
    MsgBox "先选择待上传的文件"
    Exit Sub
End If
Dim dirName As String
Dim FileName As String

'获取文件名
'    If InStrRev(Trim(Text2.Text), "\") > 0 Then
'        strFileName = Mid(Trim(Text2.Text), InStrRev(Trim(Text2.Text), "\") + 1)
'        dirName = Mid$(Trim(Text2.Text), 1, InStrRev(Trim(Text2.Text), "\"))
'    End If
    

'2015-04-27 jiayunzhang  add EQ shipping request


'Excel文件处理

    '1)打开Excel

    Set VBExcel = CreateObject("excel.application")     '创建Excle对象

    VBExcel.Visible = False

    Set xlBook = VBExcel.Workbooks.open(Text3.Text)    '打开文件

    Set xlSheet = xlBook.Worksheets(1)        '打开sheet中的表

    '判定最大列Excel中的和设定列是否相同

    If xlSheet.Range("A1").CurrentRegion.Columns.Count <> 23 Then

        MsgBox "Excel中的列数和设定的列数不一致，请确认Excel是否正确！", vbInformation, "提示"
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
        tempVal = xlSheet.Range(strChar & i).Value   '临时保存值
        
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

  

'    '判断这笔SubstrateId是否已存在，如果存在，则退出，循环下一笔
'    If (JudgeFlagVTData(vtDataTemp.DeliveryNoTemp, vtDataTemp.CUSTLOTTemp)) Then
'       MsgBox "这笔已存在，无需上传!", vbInformation, "友情提示"
'       GoTo NextRecord2
'
'    End If
    vtDataTemp.idTemp = GetEQShippingMaxID()
    Call AddEQShipping(vtDataTemp)
    SumCount = SumCount + 1

    '上传到DB
NextRecord2:

Next i


     
     xlBook.Close      '总是提示是否保存   结束Excel

    Set xlSheet = Nothing

    Set xlBook = Nothing

    Set VBExcel = Nothing

'    VBExcel.Quit




If SumCount > 0 Then
    MsgBox "已成功上传" & SumCount & "笔！", , "友情提醒"
    
    Else
        If BCResultFlag = True Then
            MsgBox "上传失败，请确认资料格式！", , "友情提醒"
            Exit Sub
        End If
    
End If





''读取CSV
'Dim source_batch_id_Temp As String
'Dim customerTemp As String
'
'customerTemp = "GC"
'
''上传OI的CSV
''处理文件名
'If Text3.Text = "" Then
'    MsgBox "先选择待上传的文件"
'    Exit Sub
'End If
'Dim dirName As String
'Dim FileName As String
'
''获取文件名
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
''                '判断wo是否为空
''
''                If Trim(gcHeaderTemp.WO_NO) = "" Then
''
''                    MsgBox "WO_NO有空值，请确认！"
''                    Exit Sub
''
''                End If
''
''                '2012-11-07 jiayun 修改Good_Die_Qty 根据市场部规则
''
''            gcDetailTemp.Good_Die_Qty = GetGCGoodDieQty(Trim(gcHeaderTemp.Customer_Device), gcDetailTemp.Good_Die_Qty)
''
''            '2013-12-27 jiayun add
''
''            If gcDetailTemp.Good_Die_Qty <= 0 Then
''                    MsgBox "请确认客户机种对应的Die数是否有维护好！"
''                    Exit Sub
''            End If
''
''
''            '2012-11-05 jiayun 修改 GC
''
''            '判断lotID在Header表中是否已存在
''
''            If (JudgeGCHeaderId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)) Then
''
''                If GCHeaderFlag = False Then
''        '            MsgBox "GC 这笔：" & gcHeaderTemp.Lot_ID & "已存在，无需上传!"
''                End If
''
''                '2013-12-05 jiayun add 如果lotid,wo_no 已存在，则查询出id
''                '当lotid有隔行时，则查询上次的id
''
''                id = GetGCLotIDWOId(gcHeaderTemp.Lot_ID, gcHeaderTemp.WO_NO)
''
''            Else
''            '上传到Header表中
''                '取目前DB最大的ID号
''                id = GetMaxID()
''                '2013-01-11 jiayun add 客户简称
''
''                If id = 0 Then
''                    MsgBox "DB主表ID生成失败1，请联系资讯！"
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
''            '判断lotID在Detail表中是否已存在
''
''            If (JudgeGCDetailId(gcDetailTemp.Lot_ID, gcDetailTemp.Wafer_ID)) Then
''               MsgBox "GC 这笔：" & gcDetailTemp.Lot_ID & "; WaferId:" & gcDetailTemp.Wafer_ID & "已存在，无需上传!"
''
''            Else
''            '上传到Detail表中
''
''                   '2012-11-05 jiayun 修改 GCT
''
''
''                   gcDetailTemp.item = gcDetailTemp.Lot_ID & Right(("0" & gcDetailTemp.Wafer_ID), 2)
''
''
''                If id = 0 Then
''                    MsgBox "DB主表ID生成失败2，请联系资讯！"
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
'            MsgBox "已成功上传" & SumCount & "笔！"
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
