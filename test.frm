VERSION 5.00
Object = "{0D300FC0-B2EA-11D1-8D3B-444553540000}#1.31#0"; "QRmaker.ocx"
Begin VB.Form test 
   Caption         =   "Test"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16650
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
   ScaleHeight     =   8040
   ScaleWidth      =   16650
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "NPI"
      Height          =   360
      Left            =   8520
      TabIndex        =   22
      Top             =   600
      Width           =   990
   End
   Begin VB.CommandButton Command6 
      Caption         =   "流水"
      Height          =   360
      Left            =   2880
      TabIndex        =   21
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton Command5 
      Caption         =   "37TESTDC更新"
      Height          =   360
      Left            =   5880
      TabIndex        =   20
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtDN 
      Height          =   375
      Left            =   2520
      TabIndex        =   19
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GD108箱号更新"
      Height          =   360
      Left            =   720
      TabIndex        =   18
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtMarkingCode 
      Height          =   405
      Left            =   10800
      TabIndex        =   17
      Top             =   3855
      Width           =   3495
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改打标码"
      Height          =   360
      Left            =   11160
      TabIndex        =   15
      Top             =   2520
      Width           =   1935
   End
   Begin QRMAKERLib.QRmaker QRmaker1 
      Height          =   975
      Left            =   2760
      TabIndex        =   14
      Top             =   5880
      Width           =   1455
      _Version        =   65567
      _ExtentX        =   2566
      _ExtentY        =   1720
      _StockProps     =   1
      CellPitch       =   12
      Unit            =   600
      Picture         =   "test.frx":0000
   End
   Begin VB.TextBox tWaferID 
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Top             =   3870
      Width           =   2055
   End
   Begin VB.TextBox tLotID 
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox tWOID 
      Height          =   285
      Left            =   7320
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox tCusCode 
      Height          =   285
      Left            =   7320
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmd 
      Caption         =   "打标码补打"
      Height          =   1200
      Left            =   7320
      TabIndex        =   5
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "开始3"
      Height          =   360
      Left            =   3000
      TabIndex        =   4
      Top             =   4920
      Width           =   990
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "开始2"
      Height          =   360
      Left            =   840
      TabIndex        =   1
      Top             =   4080
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   360
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "打标码"
      Height          =   195
      Index           =   2
      Left            =   10200
      TabIndex        =   16
      Top             =   3960
      Width           =   540
   End
   Begin VB.Label lblWaferID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WaferID"
      Height          =   195
      Left            =   6720
      TabIndex        =   12
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblLOtID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOtID"
      Height          =   195
      Left            =   6720
      TabIndex        =   10
      Top             =   3120
      Width           =   420
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工单号"
      Height          =   195
      Index           =   1
      Left            =   6720
      TabIndex        =   8
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户代码"
      Height          =   195
      Index           =   0
      Left            =   6600
      TabIndex        =   7
      Top             =   1320
      Width           =   720
   End
End
Attribute VB_Name = "test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmd_Click()

Dim strCusCode As String

If tCusCode.Text = "" Then
    MsgBox "输入客户代码", vbInformation, "提示"
    Exit Sub
End If

strCusCode = Trim$(tCusCode.Text)

Select Case strCusCode

    Case "SG005"
        InsertMark1
    Case "US026"
        InsertMark1
    Case ""

End Select


End Sub

Private Sub cmdMod_Click()
If tWaferID.Text = "" Or txtMarkingCode.Text = "" Then
    MsgBox "请输入WaferID或打标码", vbInformation, Me.Caption
    Exit Sub
End If

Dim strSql As String

strSql = "update mappingdatatest set productid = '" & Trim(txtMarkingCode.Text) & "' where substrateid = '" & Trim(tWaferID.Text) & "'"
AddSql (strSql)

strSql = "update ib_waferlist set markingcode = '" & Trim(txtMarkingCode.Text) & "' where waferid = '" & Trim(tWaferID.Text) & "'"
AddSql (strSql)

strSql = "update shop_order_detail set mark_code = '" & Trim(txtMarkingCode.Text) & "' where wafer_id = '" & Trim(tWaferID.Text) & "' "
AddSql (strSql)

strSql = "update shop_order_property set propertyvalue = '" & Trim(txtMarkingCode.Text) & "' where  wafer_id = '" & Trim(tWaferID.Text) & "'  and propertyname = 'MARKING_CODE' "
AddSql (strSql)

strSql = "update [erpdata].[dbo].[tblTSVwaferlist] set markingcode = '" & Trim(txtMarkingCode.Text) & "' where waferid = '" & Trim(tWaferID.Text) & "'  "
AddSql2 (strSql)

strSql = "update [ERPBASE].[dbo].[tblmappingData] set productid = '" & Trim(txtMarkingCode.Text) & "' where substrateid = '" & Trim(tWaferID.Text) & "' "
AddSql2 (strSql)

MsgBox "成功更新打标码", vbInformation, "提示"

End Sub

Private Sub InsertMark1()

Dim rs As New ADODB.Recordset

' If customerTemp = "US026" Or customerTemp = "SG005" Or customerTemp = "TW079" Then
'            ' 增加REMARK
'            gcDetailTemp.Marking_Lot_ID = Mid$(gcHeaderTemp.CUSTOMER_DEVICE, InStr(gcHeaderTemp.CUSTOMER_DEVICE, "-") + 2, 1)
'            gcDetailTemp.Marking_Lot_ID = gcDetailTemp.Marking_Lot_ID & Right(Year(Now), 1)
'            gcDetailTemp.Marking_Lot_ID = gcDetailTemp.Marking_Lot_ID & Hex(Month(Now))
'            gcDetailTemp.Marking_Lot_ID = gcDetailTemp.Marking_Lot_ID & Mid$(waferseq, gcDetailTemp.WAFER_ID, 1)
'
'            Dim spos As String
'
'            spos = InStr(gcHeaderTemp.lot_id, ".")
'
'            If spos > 0 Then
'                gcDetailTemp.Marking_Lot_ID = gcDetailTemp.Marking_Lot_ID & Mid$(gcHeaderTemp.lot_id, spos - 4, 4)
'            Else
'
'                gcDetailTemp.Marking_Lot_ID = gcDetailTemp.Marking_Lot_ID & Right$(gcHeaderTemp.lot_id, 4)
'
'            End If
'
'        End If

' 工单号
Dim strWOID As String
Dim strWaferID As String
Dim strLotID As String
Dim strMark As String
Dim strMPN As String
Dim strWaferSeq As String
Dim waferseq As String

 waferseq = "123456789ABCDEFGHIJKLMNOP"

If tWOID.Text <> "" Then
    strWOID = Trim$(UCase$(tWOID.Text))
End If

Set rs.ActiveConnection = OraConnect
rs.Source = "select distinct waferid,waferlot,markingcode from ib_waferlist where ordername = '" & strWOID & "' "
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

If rs.RecordCount > 0 Then
    rs.MoveFirst
    
    For i = 1 To rs.RecordCount
           strWaferID = Trim("" & rs(0))
           strLotID = Trim$("" & rs(1))
           strMark = Trim$("" & rs(2))
           
           If 1 Then
                strMPN = Trim(Get_OracleStr("select customerpn from ib_wohistory where ordername = '" & strWOID & "' "))
                
                strMark = Mid$(strMPN, InStr(strMPN, "-") + 2, 1) & Right(Year(Now), 1) & Hex(Month(Now)) & Mid$(waferseq, Right(strWaferID, 2), 1)
                
                
                Dim spos As String

                spos = InStr(strLotID, ".")
                
                
            If spos > 0 Then
                strMark = strMark & Mid$(strLotID, spos - 4, 4)
            Else

                strMark = strMark & Right$(strLotID, 4)

            End If
                
                Dim sUpdate As String
                
                sInsert = "update ib_waferlist set markingcode = '" & strMark & "' where ordername = '" & strWOID & "' and waferid = '" & strWaferID & "'"
                sInsert = "update shop_order_detail set MARK_CODE  = '" & strMark & "' where shop_order = '" & strWOID & "' and wafer_id = '" & strWaferID & "'"
                
                
                AddSql (sInsert)
                
           End If
           
           
           
           rs.MoveNext
    Next i
    
    
    
    
    
    
Else
    MsgBox "找不到工单明细", vbCritical, "警告"
    Exit Sub
End If














End Sub










































Private Sub Command1_Click()

    Dim rs   As New ADODB.Recordset

    Dim sWO  As String

    Dim sQty As Long

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    rs.Source = "select distinct ordername from [erpdata].[dbo].[tblTSVworkorder] where qty = 0 and ordername like '%1807%'"
    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount
            sWO = rs("ordername")
            sQty = Get_SqlserverNo("select sum(CONVERT(int,dieqty)) as qty from [erpdata].[dbo].[tblTSVwaferlist] where ordername = '" & sWO & "'")
            Call updateQty(sWO, sQty)
            
            rs.MoveNext
        Next i

    End If

End Sub

Private Sub updateQty(sWO As String, sQty As Long)

    Dim strSql As String

    strSql = "update [erpdata].[dbo].[tblTSVworkorder] set qty = '" & sQty & "' where ordername = '" & sWO & "'"

    AddSql2 (strSql)

End Sub

Private Sub Command2_Click()

'    Dim strSql As String
'
'    strSql = "insert into erpdata..TblQBOXNUMBER_TSV(CONTAINERID,NDPW,WAFERNUMBER,WAFERSCRIBENUMBER,WORKORDERNAME,FIRSTNAME,QBOXNUMBER,CONTAINERNAME,FLAG,CUSTOMERNAME,PDATA1,PRODUCTNAME,SPECNAME,EndStatus) " & "select CONTAINERID,NDPW,WAFERNUMBER,WAFERSCRIBENUMBER,WORKORDERNAME,FIRSTNAME,QBOXNUMBER,CONTAINERNAME,FLAG,CUSTOMERNAME,PDATA1,PRODUCTNAME,SPECNAME,EndStatus from " & "(select * from OPENQUERY(ORACLEDB, 'SELECT * from TSV_QBOXNUMBER_DETAILS' )) AA where AA.QBOXNUMBER = '" & Trim(Text2.Text) & "' and AA.WAFERSCRIBENUMBER = '" & Trim(Text1.Text) & "' "
'
'    Dim i As Integer
'
'    i = AddSql2(strSql)
'
'    If i > 0 Then
'        MsgBox "成功插入" & i & "笔"
'    End If

    Dim rs As New ADODB.Recordset

    Set rs.ActiveConnection = OraConnect

    rs.Source = " select CONTAINERID,NDPW,WAFERNUMBER,WAFERSCRIBENUMBER,WORKORDERNAME,FIRSTNAME,QBOXNUMBER,CONTAINERNAME,FLAG,CUSTOMERNAME,PDATA1,PRODUCTNAME,SPECNAME,EndStatus from TSV_QBOXNUMBER_DETAILS where WAFERSCRIBENUMBER = '" & Trim(Text1.Text) & "'"

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        
         rs.MoveFirst

        For i = 1 To rs.RecordCount - 1
            If rs.BOF = True Then
                Exit For
            End If
            s1 = rs(0)
            s2 = rs(1)
            s3 = rs(2)
            s4 = rs(3)
            s5 = rs(4)
            s6 = rs(5)
            s7 = rs(6)
            s8 = rs(7)
            s9 = rs(8)
            s10 = rs(9)
            s11 = rs(10)
            s12 = rs(11)
            s13 = rs(12)
            s14 = rs(13)
            
            
            Dim sSql As String
            
            sSql = "insert into erpdata..TblQBOXNUMBER_TSV(CONTAINERID,NDPW,WAFERNUMBER,WAFERSCRIBENUMBER,WORKORDERNAME,FIRSTNAME,QBOXNUMBER,CONTAINERNAME,FLAG,CUSTOMERNAME,PDATA1,PRODUCTNAME,SPECNAME,EndStatus) " & _
            " values('" & s1 & "', '" & s2 & "', '" & s3 & "','" & s4 & "', '" & s5 & "', '" & s6 & "','" & s7 & "', '" & s8 & "', '" & s9 & "','" & s10 & "', '" & s11 & "', '" & s12 & "','" & s13 & "','" & s14 & "') "
        
            i = AddSql2(sSql)
            If i > 0 Then
                MsgBox "成功插入" & i & "笔"
            End If
            

        
            rs.MoveNext
        Next i
        
    End If

End Sub

Private Sub Command3_Click()
 
  Dim rs   As New ADODB.Recordset

    Dim sWO  As String

    Dim sQty As Long

 
 Set rs = New ADODB.Recordset
Set rs.ActiveConnection = SqlConnect


    rs.Source = "select distinct  采购单编号, 采购金额 from tblcpurdata order by 采购单编号"
    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount
            cgdh = Trim(rs(0))
            cgje = Trim(rs(1))
            
            
            sQty = Get_SqlserverNo("select SUM(采购数量) from tblCPurDataSub where 采购单编号 = '" & cgdh & "'  ")
            
            sdj = cgje / sQty
            
            AddSql2 ("update tblCPurDataSub set 单价 = '" & sdj & "'  where 采购单编号 = '" & cgdh & "'")
            
            rs.MoveNext
            
            Sleep (10)
        Next i

    End If

End Sub

'Private Sub Command4_Click()
'Dim strSql   As String
'Dim strsql2  As String
'Dim rs       As ADODB.Recordset
'Dim strID    As String
'Dim strBoxID As String
'Dim strSp()  As String
'Dim i        As Integer
'Dim strDC As String
'Dim strTray As String
'
'strSql = "select trayid from packing_detailed where dn_num = '" & Trim(txtDN.Text) & "'  order by seq desc"
'Set rs = Get_OracleRs(strSql)
'If Not rs.EOF Then
'
'    Do While Not rs.EOF
'        strTray = rs!TRAYID
'
'        strDC = Get_SqlStr("SELECT  right(datename(yy,t1.ERPCREATEDATE),2) + datename(ww,t1.ERPCREATEDATE) from [erpdata].[dbo].[tblTSVworkorder] t1" & _
'" inner join erpdata..tblStockNumSub t2 on t2.大工单 = t1.ORDERNAME " & _
'" where  t2.箱号 = '" & strTray & "' ")
'
'        AddSql ("update packing_detailed set datecode = '" & strDC & "' where dn_num = '" & Trim(txtDN.Text) & "' and trayid = '" & strTray & "'")
''
''        rs.MoveNext
''    Loop
''
''End If
''
''MsgBox "更新完成", vbInformation, "提示"
''
''End Sub
'
Private Sub Command4_Click()
Dim strSql   As String
Dim strSql2  As String
Dim rs       As ADODB.Recordset
Dim strID    As String
Dim strBoxID As String
Dim strSp()  As String
Dim i        As Integer
Dim strDC As String
Dim strTray As String
Dim strJob As String
Dim strOldDC As String
Dim strCZDC As String
Dim strRealDC As String
Dim strSql1 As String
Dim strSql3 As String
Dim strSql4 As String

'strSql = "select distinct job_id,datecode from packing_detailed where datecode is not null"
'Set rs = Get_OracleRs(strSql)
'If Not rs.EOF Then
'
'    Do While Not rs.EOF
'        strJob = rs!JOB_ID
'        strOldDC = rs!DATECODE
'
'        strJobNew = Replace$(strJob, "M", "")
'strSql1 = "select distinct case when create_date >= to_date(to_char(create_date, 'yyyy') || '-12-31', 'yyyy-mm-dd') - mod(to_char(create_date, 'YYYY'), 7) - 5  then to_char(create_date, 'yyww') " & "else to_char(create_date + mod(mod(to_char(create_date, 'YYYY'), 7) + 5, 7),'yyww') end as PODATECODE " & "from customeroitbl_test a ,mappingdatatest b ,weight37 c where a.test_mtrl_desc = '" & strJobNew & "' and b.filename = to_char(a.id) and b.lotid = a.source_batch_id " & "and c.waferid = replace(b.substrateid,'+','') "
'strCZDC = Get_OracleStr(strSql1)
'
'        strSql2 = "select TT.TEST_DATECODE from (  " & _
'" select t4.TEST_MTRL_DESC as JOBID,min(DATE_CODE_CONVERT.DC_CONVERT(to_char(t2.ERPCREATEDATE, 'YYYY-MM-DD'),1)) as TEST_DATECODE " & _
'" from ib_waferlist t1 " & _
'" inner join ib_workorder t2 on t1.ORDERNAME = t2.ORDERNAME " & _
'" inner join mappingdatatest t3 on t3.SUBSTRATEID = t1.WAFERID and t3.LOTID = t1.WAFERLOT " & _
'" inner join  customeroitbl_test t4 on t4.SOURCE_BATCH_ID = t3.LOTID  and to_char(t4.ID) = t3.FILENAME " & _
'" where t4.TEST_MTRL_DESC in ('" & strJobNew & "')  group by t4.TEST_MTRL_DESC  ) TT where TT.TEST_DATECODE >=1929 "
'
'strTestDateCode = Get_OracleStr(strSql2)
'
'        If strTestDateCode <> "" Then
'
'            strRealDC = strTestDateCode
'        Else
'            strRealDC = strCZDC
'
'        End If
'
'        '如果和包装不一致,则更新packing_detail并作标记
'        If strRealDC <> strOldDC Then
'
'            strSql3 = "update packing_detailed set datecode = '" & strRealDC & "' , CREATE_BY = 'WRONG' where job_id = '" & strJob & "' "
'            AddSql (strSql3)
'
'
'        End If
'
'        '存清单
'        strSql4 = "insert into TBL37TESTDC(JOBID,DC) values('" & strJob & "','" & strRealDC & "')"
'
'        AddSql (strSql4)
'
'        rs.MoveNext
'    Loop
'
'End If
'
'MsgBox "更新完成", vbInformation, "提示"
'





strSql = "select distinct a.test_mtrl_desc from customeroitbl_test a where  a.test_mtrl_desc <> a.source_mtrl_sloc and not exists (select 1 from TBL37TESTDC b where b.jobid = a.test_mtrl_desc )"

Set rs = Get_OracleRs(strSql)
If Not rs.EOF Then

    Do While Not rs.EOF
        strJob = rs!test_mtrl_desc

        strJobNew = Replace$(strJob, "M", "")
strSql1 = "select distinct case when create_date >= to_date(to_char(create_date, 'yyyy') || '-12-31', 'yyyy-mm-dd') - mod(to_char(create_date, 'YYYY'), 7) - 5  then to_char(create_date, 'yyww') " & "else to_char(create_date + mod(mod(to_char(create_date, 'YYYY'), 7) + 5, 7),'yyww') end as PODATECODE " & "from customeroitbl_test a ,mappingdatatest b ,weight37 c where a.test_mtrl_desc = '" & strJobNew & "' and b.filename = to_char(a.id) and b.lotid = a.source_batch_id " & "and c.waferid = replace(b.substrateid,'+','') "
strCZDC = Get_OracleStr(strSql1)

        strSql2 = "select TT.TEST_DATECODE from (  " & _
" select t4.TEST_MTRL_DESC as JOBID,min(DATE_CODE_CONVERT.DC_CONVERT(to_char(t2.ERPCREATEDATE, 'YYYY-MM-DD'),1)) as TEST_DATECODE " & _
" from ib_waferlist t1 " & _
" inner join ib_workorder t2 on t1.ORDERNAME = t2.ORDERNAME " & _
" inner join mappingdatatest t3 on t3.SUBSTRATEID = t1.WAFERID and t3.LOTID = t1.WAFERLOT " & _
" inner join  customeroitbl_test t4 on t4.SOURCE_BATCH_ID = t3.LOTID  and to_char(t4.ID) = t3.FILENAME " & _
" where t4.TEST_MTRL_DESC in ('" & strJobNew & "')  group by t4.TEST_MTRL_DESC  ) TT where TT.TEST_DATECODE >=1929 "

strTestDateCode = Get_OracleStr(strSql2)

        If strTestDateCode <> "" Then

            strRealDC = strTestDateCode
        ElseIf strCZDC <> "" Then
            strRealDC = strCZDC

        End If

        '如果和包装不一致,则更新packing_detail并作标记

        '存清单
        
        If strRealDC <> "" Then
            strSql4 = "insert into TBL37TESTDC(JOBID,DC) values('" & strJob & "','" & strRealDC & "')"
            AddSql (strSql4)
        End If
  
        rs.MoveNext
    Loop

End If

MsgBox "更新完成", vbInformation, "提示"

End Sub


'Private Sub Command4_Click()
'Dim strSql   As String
'Dim strSql2  As String
'Dim rs       As ADODB.Recordset
'Dim strID    As String
'Dim strBoxID As String
'Dim strSp()  As String
'Dim i        As Integer
'Dim strDC As String
'Dim strTray As String
'
'
'Dim strJob As String
'Dim strOldDC As String
'Dim strCZDC As String
'Dim strRealDC As String
'Dim strSql1 As String
'Dim strSql3 As String
'Dim strSql4 As String
'
'strSql = "select distinct job_id,datecode from packing_detailed where length(datecode) > 4 "
'Set rs = Get_OracleRs(strSql)
'If Not rs.EOF Then
'
'    Do While Not rs.EOF
'        strJob = rs!JOB_ID
'        strOldDC = rs!DATECODE
'
'        strJobNew = Replace$(strJob, "M", "")
'strSql1 = "select distinct case when create_date >= to_date(to_char(create_date, 'yyyy') || '-12-31', 'yyyy-mm-dd') - mod(to_char(create_date, 'YYYY'), 7) - 5  then to_char(create_date, 'yyww') " & "else to_char(create_date + mod(mod(to_char(create_date, 'YYYY'), 7) + 5, 7),'yyww') end as PODATECODE " & "from customeroitbl_test a ,mappingdatatest b ,weight37 c where a.test_mtrl_desc = '" & strJobNew & "' and b.filename = to_char(a.id) and b.lotid = a.source_batch_id " & "and c.waferid = replace(b.substrateid,'+','') "
'strCZDC = Get_OracleStr(strSql1)
'
'        strSql2 = "select TT.TEST_DATECODE from (  " & _
'" select t4.TEST_MTRL_DESC as JOBID,min(DATE_CODE_CONVERT.DC_CONVERT(to_char(t2.ERPCREATEDATE, 'YYYY-MM-DD'),1)) as TEST_DATECODE " & _
'" from ib_waferlist t1 " & _
'" inner join ib_workorder t2 on t1.ORDERNAME = t2.ORDERNAME " & _
'" inner join mappingdatatest t3 on t3.SUBSTRATEID = t1.WAFERID and t3.LOTID = t1.WAFERLOT " & _
'" inner join  customeroitbl_test t4 on t4.SOURCE_BATCH_ID = t3.LOTID  and to_char(t4.ID) = t3.FILENAME " & _
'" where t4.TEST_MTRL_DESC in ('" & strJobNew & "')  group by t4.TEST_MTRL_DESC  ) TT where TT.TEST_DATECODE >=1929 "
'
'strTestDateCode = Get_OracleStr(strSql2)
'
'        If strTestDateCode <> "" Then
'
'            strRealDC = strTestDateCode
'        Else
'            strRealDC = strCZDC
'
'        End If
'
'        '如果和包装不一致,则更新packing_detail并作标记
'        If strRealDC <> strOldDC Then
'
'            strSql3 = "update packing_detailed set datecode = '" & strRealDC & "' , CREATE_BY = 'WRONG' where job_id = '" & strJob & "' "
'            AddSql (strSql3)
'
'
'        End If
'
'        '存清单
'        strSql4 = "update TBL37TESTDC set DC = '" & strRealDC & "' where JOBID =  '" & strJob & "'"
'
'        AddSql (strSql4)
'
'        rs.MoveNext
'    Loop
'
'End If
'
'MsgBox "更新完成", vbInformation, "提示"
'
'End Sub

Private Sub Command5_Click()
Dim strSql    As String
Dim strJobID  As String
Dim strDC1    As String
Dim strDC2    As String
Dim strTestDC As String
Dim i         As Integer
Dim rs        As New ADODB.Recordset

strSql = "select distinct a.test_mtrl_desc from customeroitbl_test a where a.test_mtrl_desc <> a.source_mtrl_sloc and a.customershortname = '37' "
Set rs = Get_OracleRs(strSql)
If rs.RecordCount > 0 Then

    For i = 1 To rs.RecordCount
        strJobID = rs!test_mtrl_desc
        If Right$(strJobID, 1) <> "R" Then
            strSql = "select distinct t3.ordername from mappingdatatest t1 " & "inner join customeroitbl_test t2 on t1.lotid = t2.source_batch_id and to_char(t2.id) = t1.filename " & "inner join ib_waferlist t3 on t3.waferid = t1.substrateid and t3.waferlot = t1.lotid " & "where t2.test_mtrl_desc = '" & strJobID & "' "
            If Get_OracleCnt(strSql) <> 0 Then
                strSql = " select TT.TEST_DATECODE from ( select t4.TEST_MTRL_DESC as JOBID,min(DATE_CODE_CONVERT.DC_CONVERT(to_char(t2.ERPCREATEDATE, 'YYYY-MM-DD'),1)) as TEST_DATECODE " & " from ib_waferlist t1 " & " inner join ib_wohistory t2 on t1.ORDERNAME = t2.ORDERNAME " & " inner join mappingdatatest t3 on t3.SUBSTRATEID = t1.WAFERID and t3.LOTID = t1.WAFERLOT " & " inner join  customeroitbl_test t4 on t4.SOURCE_BATCH_ID = t3.LOTID  and to_char(t4.ID) = t3.FILENAME " & " where t4.TEST_MTRL_DESC in ('" & strJobID & "')  group by t4.TEST_MTRL_DESC  ) TT where TT.TEST_DATECODE >=1929 "
                strDC1 = Get_OracleStr(strSql)
                If strDC1 <> "" Then
                    strTestDC = strDC1
                    AddSql ("insert into tbl37testdc_bak(JOBID,DC) values('" & strJobID & "','" & strTestDC & "')")
                Else
                    strSql = "select distinct case when create_date >= to_date(to_char(create_date, 'yyyy') || '-12-31', 'yyyy-mm-dd') - mod(to_char(create_date, 'YYYY'), 7) - 5  then to_char(create_date, 'yyww') " & "else to_char(create_date + mod(mod(to_char(create_date, 'YYYY'), 7) + 5, 7),'yyww') end as PODATECODE " & "from customeroitbl_test a ,mappingdatatest b ,weight37 c where a.test_mtrl_desc = '" & strJobID & "' and b.filename = to_char(a.id) and b.lotid = a.source_batch_id " & "and c.waferid = replace(b.substrateid,'+','') "
                    strDC2 = Get_OracleStr(strSql)
                    If strDC2 <> "" Then
                        strTestDC = strDC2
                        AddSql ("insert into tbl37testdc_bak(JOBID,DC) values('" & strJobID & "','" & strTestDC & "')")

                    End If

                End If

            End If

        End If

        rs.MoveNext
    Next i

End If

End Sub

Private Sub Command6_Click()
Dim strSql    As String
Dim rs        As New ADODB.Recordset
Dim i         As Integer
Dim strLotID  As String
Dim strMaxID1 As String
Dim strMaxID2 As String
Dim strMaxID  As String

strSql = "select distinct LOTID from ERPBASE..tblmappingData where CUSTOMERSHORTNAME in ('GD108','HK080') "
Set rs = Get_SqlserveRs(strSql)
If rs.RecordCount > 0 Then

    For i = 1 To rs.RecordCount
        strLotID = rs!LOTID
        strMaxID1 = Get_SqlStr("select isnull(MAX(right(Content2,2)),0) as maxSeq from erpdata..tblME_PrintInfo where charindex('" & strLotID & "' + '-C',Content2) > 0 ")
        strMaxID2 = Get_SqlStr("select isnull(MAX(right(EVENT_ID,2)),0) as maxSeq from erpdata..tblME_PrintInfo where charindex('" & strLotID & "' + '-C',EVENT_ID) > 0 ")
        If CLng(strMaxID1) >= CLng(strMaxID2) Then
            strMaxID = strMaxID1
        Else
            strMaxID = strMaxID2

        End If

        AddSql2 ("insert into erptemp..tblGD108lotboxidseq(LOTID,BOXTYPE,MAXSEQ) values('" & strLotID & "','-C','" & strMaxID & "') ")
        
        rs.MoveNext
    Next i

End If
MsgBox "更新完成", vbInformation, "提示"
End Sub

Private Sub Command7_Click()
Dim strCustCode As String
Dim strCustPN As String
Dim strHTPN As String
Dim strProduct As String
Dim strSql As String
Dim strSql2 As String
Dim rs As New ADODB.Recordset

strSql = "select customershortname,customerptno1,qtechptno,qtechptno2 from tbltsvnpiproduct_bak where customershortname = 'HK075' "
Set rs = Get_OracleRs(strSql)
rs.MoveFirst
Do While Not rs.EOF
    strCustCode = rs!CUSTOMERSHORTNAME
    strCustPN = rs!CustomerPTNo1
    strHTPN = rs!qtechPTNo
    strProduct = rs!QtechPTNo2
    
    strSql2 = "select * from tbltsvnpiproduct where customershortname = '" & strCustCode & "' and customerptno1 = '" & strCustPN & "' and  qtechptno = '" & strHTPN & "' and  qtechptno2 = '" & strProduct & "'  "
    If Get_OracleCnt(strSql2) = 0 Then
        AddSql ("insert into tbltsvnpiproduct select * from tbltsvnpiproduct_bak where customershortname = '" & strCustCode & "' and customerptno1 = '" & strCustPN & "' and  qtechptno = '" & strHTPN & "' and  qtechptno2 = '" & strProduct & "'")
    End If
    rs.MoveNext
Loop


MsgBox "全部插入完成"





























End Sub
