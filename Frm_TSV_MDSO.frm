VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Begin VB.Form Frm_TSV_MDSO 
   Caption         =   "�г��� �������ܱ����ѯ"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0FF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9885
   ScaleWidth      =   17115
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   360
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "��ѯ"
      Height          =   360
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   9015
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   16695
      _Version        =   196613
      _ExtentX        =   29448
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
      SpreadDesigner  =   "Frm_TSV_MDSO.frx":0000
      TextTip         =   2
   End
End
Attribute VB_Name = "Frm_TSV_MDSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail�֭�
    E_ID = 0                'id��
    E_Factory            '��ҵ��
    E_Customer           '�ͻ�����
    E_ToDayQty            '����Ƭ��
    E_Mon1                'һ��
    E_Mon2                '����
    E_Mon3                '����
    E_Mon4                '����
    E_Mon5                '����
    E_Mon6                '����
    E_Mon7                '7��
    E_Mon8                '8��
    E_Mon9                '9��
    E_Mon10                '10��
    E_Mon11                '11��
    E_Mon12                '12��
    E_sum                  '����
    

    E_End
    
    
End Enum



Dim Rs As New ADODB.Recordset
Dim mainItemRS As New ADODB.Recordset
Dim bomRS2        As New ADODB.Recordset



Private Sub CmdOut_Click()
Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String


Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(TxtWo.Text))
productTemp = UCase(Trim(TxtProduct.Text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as Ƭ�� , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
          
sql3 = " order by a.ordername,b.waferid  "

  
  If productTemp <> "" Then
  
  sql2 = " and a.product='" + productTemp + "'"
  
  End If
  
  
 If woTemp <> "" Then
  
  sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
  End If
  
  If Trim(sql2) <> "" Then
  
  sqlTemp = sql1 & sql2 & sql3
  
  Else
  
  sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<=to_date('" + endTime + "','YYYY/MM/DD')"
  
   sqlTemp = sql1 & sql2 & sql3
  
  End If
  
  
  
  
  
  ExporToExcel (sqlTemp)



End Sub

Private Sub Command1_Click()


Dim waferidTemp As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""
waferidTemp = UCase(Trim(TxtWaferID.Text))


'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as Ƭ�� , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
'  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
'"  CASE ORDERTYPE  WHEN '1' THEN 'һ�㹤��'  WHEN '5' THEN '�ټӹ�����'   WHEN '7' THEN 'ί�⹤��'   WHEN '8' THEN '�ع�ί�⹤��' " & _
'" WHEN '11' THEN '���ʽ����'    WHEN '13' THEN 'Ԥ�⹤��'   WHEN '15' THEN '�Բ�����' Else '����' END as ORDERTYPE ," & _
'"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
'" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
'" from erpintegration2.ib_wohistory a where  modifyflag='1' "


  sql1 = " select b.wafersequence, b.ordername,b.waferid,b.completeflag,b.dieqty, b.FGDieQty , b.WaferLot, b.MarkingCode " & _
"from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername and b.waferid='" + waferidTemp + "' "

          
          
sql3 = " order by a.ordername ,b.WaferLot,b.waferid "

  
 
  
  If Trim(sql2) <> "" Then
  
  sqlTemp = sql1 & sql2 & sql3
  
  Else
  
  sql2 = ""
  
   sqlTemp = sql1 & sql2 & sql3
  
  End If
  
  
  
  



Set mainItemRS = GetPMCWOLine(sqlTemp)

With fps(1)
        .MaxRows = 0
        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
       
        End If
End With





End Sub

Private Sub CmdQuery_Click()
Dim strSql As String
Dim m As Integer
Dim i As Integer
Dim factoryIdTemp As Integer
Dim custIDTemp As Integer
Dim custNameTemp As String
Dim customerTemp As String
Dim ptFirstTemp As String
Dim ptLastTemp As String
Dim currentRowIDTemp As Integer




Dim j As Integer
Dim qtyMonTemp As Long
Dim monThTemp As String





' ----8��TSV--------------
'--1 AA   2- 18P  3- 7M  4- TSV
ptFirstTemp = "18"
ptLastTemp = "TSV"

i = 0
With fps(0)

        Set Rs = GetMDSOCustomer("18", "TSV")
        
        If Rs.RecordCount > 0 Then
              .MaxRows = Rs.RecordCount + 1
        
              For m = 1 To Rs.RecordCount
        
                    customerTemp = Rs("custid").Value
                    
                         i = i + 1
                        .Row = i
                        .Col = 1
                        .Text = "8��TSV"
                        
                        .Row = i
                        .Col = 2
                        .Text = customerTemp
                        
                        .Row = i
                        .Col = 3
                        .Text = GetQtyMDSOMonthDay(customerTemp, ptFirstTemp, ptLastTemp)
                         
                         For j = 1 To 12
                         
                         monThTemp = Right(0 & j, 2)
                         
                         qtyMonTemp = GetQtyMDSOMonth(customerTemp, ptFirstTemp, monThTemp, ptLastTemp)
                         
                         
                        .Row = i
                        .Col = 3 + j
                        .Text = qtyMonTemp
                         
                         Next j
                 
                    Rs.MoveNext
              Next m
        
        End If
        
        currentRowIDTemp = .MaxRows - 1

 End With

Rs.Close

Set Rs = Nothing


'---------------------01


' ----8��Bumping--------------
'--1 AA   2- 18P  3- 7M  4- TSV
ptFirstTemp = "18"
ptLastTemp = "Bumping"

i = 0
i = currentRowIDTemp

With fps(0)

        Set Rs = GetMDSOCustomer("18", "Bumping")
        
        If Rs.RecordCount > 0 Then
              .MaxRows = currentRowIDTemp + Rs.RecordCount
        
              For m = 1 To Rs.RecordCount
        
                    customerTemp = Rs("custid").Value
                    
                         i = i + 1
                        .Row = i
                        .Col = 1
                        .Text = "8��Bumping"
                        
                        .Row = i
                        .Col = 2
                        .Text = customerTemp
                        
                        .Row = i
                        .Col = 3
                        .Text = GetQtyMDSOMonthDay(customerTemp, ptFirstTemp, ptLastTemp)
                         
                         For j = 1 To 12
                         
                         monThTemp = Right(0 & j, 2)
                         
                         qtyMonTemp = GetQtyMDSOMonth(customerTemp, ptFirstTemp, monThTemp, ptLastTemp)
                         
                         
                        .Row = i
                        .Col = 3 + j
                        .Text = qtyMonTemp
                         
                         Next j
                 
                    Rs.MoveNext
              Next m
        
        End If
        
        currentRowIDTemp = .MaxRows - 1

 End With

Rs.Close

Set Rs = Nothing


'----------------------


currentRowIDTemp = currentRowIDTemp + 1

' ----12��TSV--------------
'--1 AA   2- 18P  3- 7M  4- TSV
ptFirstTemp = "19"
ptLastTemp = "TSV"

i = 0
i = currentRowIDTemp

With fps(0)

        Set Rs = GetMDSOCustomer("19", "TSV")
        
        If Rs.RecordCount > 0 Then
              .MaxRows = currentRowIDTemp + Rs.RecordCount
        
              For m = 1 To Rs.RecordCount
        
                    customerTemp = Rs("custid").Value
                    
                         i = i + 1
                        .Row = i
                        .Col = 1
                        .Text = "12��TSV"
                        
                        .Row = i
                        .Col = 2
                        .Text = customerTemp
                        
                        .Row = i
                        .Col = 3
                        .Text = GetQtyMDSOMonthDay(customerTemp, ptFirstTemp, ptLastTemp)
                         
                         For j = 1 To 12
                         
                         monThTemp = Right(0 & j, 2)
                         
                         qtyMonTemp = GetQtyMDSOMonth(customerTemp, ptFirstTemp, monThTemp, ptLastTemp)
                         
                         
                        .Row = i
                        .Col = 3 + j
                        .Text = qtyMonTemp
                         
                         Next j
                 
                    Rs.MoveNext
              Next m
        
        End If
        
        currentRowIDTemp = .MaxRows - 1

 End With

Rs.Close

Set Rs = Nothing


'-------------------------03

' ----12��Bumping--------------
'--1 AA   2- 18P  3- 7M  4- TSV

currentRowIDTemp = currentRowIDTemp + 1


ptFirstTemp = "19"
ptLastTemp = "Bumping"

i = 0
i = currentRowIDTemp

With fps(0)

        Set Rs = GetMDSOCustomer("19", "Bumping")
        
        If Rs.RecordCount > 0 Then
              .MaxRows = currentRowIDTemp + Rs.RecordCount
        
              For m = 1 To Rs.RecordCount
        
                    customerTemp = Rs("custid").Value
                    
                         i = i + 1
                        .Row = i
                        .Col = 1
                        .Text = "12��Bumping"
                        
                        .Row = i
                        .Col = 2
                        .Text = customerTemp
                        
                        .Row = i
                        .Col = 3
                        .Text = GetQtyMDSOMonthDay(customerTemp, ptFirstTemp, ptLastTemp)
                         
                         For j = 1 To 12
                         
                         monThTemp = Right(0 & j, 2)
                         
                         qtyMonTemp = GetQtyMDSOMonth(customerTemp, ptFirstTemp, monThTemp, ptLastTemp)
                         
                         
                        .Row = i
                        .Col = 3 + j
                        .Text = qtyMonTemp
                         
                         Next j
                 
                    Rs.MoveNext
              Next m
        
        End If
        
        currentRowIDTemp = .MaxRows

 End With

Rs.Close

Set Rs = Nothing





'      i = 0
'      With fps(0)
'            .MaxRows = Rs.RecordCount + 1
'
'              For m = 1 To Rs.RecordCount
'                factoryIdTemp = CInt(Rs("id").Value)
'                custIDTemp = CInt(Rs("cusID").Value)
'                custNameTemp = Rs("customerid").Value
'
'
'                i = i + 1
'                .Row = i
'                .Col = 1
'                .Text = Rs("workname").Value
'
'                .Row = i
'                .Col = 2
'                .Text = custNameTemp
'
'                .Row = i
'                .Col = 3
'                .Text = GetQtyMDDay(factoryIdTemp, custIDTemp, j, custNameTemp)
'
'
'
'                 For j = 1 To 12
'
'                 qtyMonTemp = GetQtyMDMonth(factoryIdTemp, custIDTemp, j, custNameTemp)
'
'
'                .Row = i
'                .Col = 3 + j
'                .Text = qtyMonTemp
'
'                 Next j
'
'
'
'               Rs.MoveNext
'
'              Next m
'
'         End With
         
'����

'�������һ�У�

'Dim iclo As Integer
'Dim irow As Integer
'Dim qtyAll As Long
'Dim qty01Temp As Long
'
'
'
'
'With fps(0)
' .Row = .MaxRows
'.Col = 1
'.Text = "Total��"
'
'
'For iclo = 3 To 15
'    qtyAll = 0
'
'     For irow = 1 To .MaxRows - 1
'        .Row = irow
'        .Col = iclo
'        qty01Temp = CLng(.Text)
'
'        qtyAll = qtyAll + qty01Temp
'
'     Next irow
'
'      .Row = .MaxRows
'      .Col = iclo
'      .Text = qtyAll
'
'Next iclo
'
'End With


'---------------------------------------temp

'����

'�������һ��:

Dim iclo As Integer
Dim irow As Integer
Dim qtyAll As Long
Dim qty01Temp As Long

Dim TSVQtyALL As Long
Dim BumpingQtyALL As Long
Dim T12QtyALL As Long
Dim T12BumpingQtyALL As Long

Dim TSVQtyRID As Integer
Dim BumpingQtyRID As Integer
Dim T12QtyRID As Integer
Dim T12BumpingQtyRID As Integer



TSVQtyALL = 0
BumpingQtyALL = 0
T12QtyALL = 0
T12BumpingQtyALL = 0


TSVQtyRID = 0
BumpingQtyRID = 0
T12QtyRID = 0
T12BumpingQtyRID = 0


With fps(0)

.MaxRows = .MaxRows + 4



  For irow = 1 To .MaxRows - 5
        .Row = irow
        .Col = 1

        If .Text = "8��TSV" Then
            TSVQtyRID = irow
        End If

        If .Text = "8��Bumping" Then
            BumpingQtyRID = irow
        End If

        If .Text = "12��TSV" Then
            T12QtyRID = irow
        End If

        If .Text = "12��Bumping" Then
            T12BumpingQtyRID = irow
        End If




  Next irow

End With



With fps(0)

 .MaxRows = .MaxRows + 1
 
 
 .Row = .MaxRows - 4
.Col = 1
.Text = "8��TSV���ܣ�"

 .Row = .MaxRows - 3
.Col = 1
.Text = "8��Bumping���ܣ�"

 .Row = .MaxRows - 2
.Col = 1
.Text = "12��TSV���ܣ�"

 .Row = .MaxRows - 1
.Col = 1
.Text = "12��Bumping���ܣ�"

 .Row = .MaxRows
.Col = 1
.Text = "���л��ܣ�"



For iclo = 3 To 15
    qtyAll = 0

    TSVQtyALL = 0
    BumpingQtyALL = 0
    T12QtyALL = 0
    T12BumpingQtyALL = 0
    


     For irow = 1 To TSVQtyRID
        .Row = irow
        .Col = iclo
        qty01Temp = CLng(.Text)

        TSVQtyALL = TSVQtyALL + qty01Temp

     Next irow


       For irow = TSVQtyRID + 1 To BumpingQtyRID
        .Row = irow
        .Col = iclo
        qty01Temp = CLng(.Text)

        BumpingQtyALL = BumpingQtyALL + qty01Temp

      Next irow
      
      
    For irow = BumpingQtyRID + 1 To T12QtyRID
        .Row = irow
        .Col = iclo
        qty01Temp = CLng(.Text)

        T12QtyALL = T12QtyALL + qty01Temp

      Next irow
      
      

    For irow = T12QtyRID + 1 To .MaxRows - 5

        .Row = irow
        .Col = iclo
        qty01Temp = CLng(.Text)

        T12BumpingQtyALL = T12BumpingQtyALL + qty01Temp

      Next irow
      
      
      

      .Row = .MaxRows - 4
      .Col = iclo
      .Text = TSVQtyALL

      .Row = .MaxRows - 3
      .Col = iclo
      .Text = BumpingQtyALL

      .Row = .MaxRows - 2
      .Col = iclo
      .Text = T12QtyALL
      
       .Row = .MaxRows - 1
      .Col = iclo
      .Text = T12BumpingQtyALL


       .Row = .MaxRows
      .Col = iclo
      .Text = TSVQtyALL + BumpingQtyALL + T12QtyALL + T12BumpingQtyALL


Next iclo

End With




''�������һ�У�



With fps(0)

For iclo = 1 To .MaxRows
    qtyAll = 0

     For irow = 4 To 15
        .Row = iclo
        .Col = irow
        qty01Temp = CLng(.Text)

        qtyAll = qtyAll + qty01Temp

     Next irow

      .Row = iclo
      .Col = 16
      .Text = qtyAll


Next iclo

End With

 
 MsgBox "��ѯ�ɹ�", vbInformation, "������ʾ"
       
End Sub

Private Sub Command2_Click()




  Call ExportFpspreadToExcel(fps(0))






End Sub

Private Sub ComOutLine_Click()

Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(TxtWo.Text))
productTemp = UCase(Trim(TxtProduct.Text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as Ƭ�� , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
"  CASE ORDERTYPE  WHEN '1' THEN 'һ�㹤��'  WHEN '5' THEN '�ټӹ�����'   WHEN '7' THEN 'ί�⹤��'   WHEN '8' THEN '�ع�ί�⹤��' " & _
" WHEN '11' THEN '���ʽ����'    WHEN '13' THEN 'Ԥ�⹤��'   WHEN '15' THEN '�Բ�����' Else '����' END as ORDERTYPE ," & _
"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
" from erpintegration2.ib_wohistory a where  modifyflag='1' "
          
          
sql3 = " order by a.ordername  "

  
  If productTemp <> "" Then
  
  sql2 = " and a.product='" + productTemp + "'"
  
  End If
  
  
 If woTemp <> "" Then
  
  sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
  End If
  
  If Trim(sql2) <> "" Then
  
  sqlTemp = sql1 & sql2 & sql3
  
  Else
  
  sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
   sqlTemp = sql1 & sql2 & sql3
  
  End If
  
  
  
   ExporToExcel (sqlTemp)










End Sub

Private Sub ComQueryHead_Click()
'HEAD

Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(TxtWo.Text))
productTemp = UCase(Trim(TxtProduct.Text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as Ƭ�� , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
"  CASE ORDERTYPE  WHEN '1' THEN 'һ�㹤��'  WHEN '5' THEN '�ټӹ�����'   WHEN '7' THEN 'ί�⹤��'   WHEN '8' THEN '�ع�ί�⹤��' " & _
" WHEN '11' THEN '���ʽ����'    WHEN '13' THEN 'Ԥ�⹤��'   WHEN '15' THEN '�Բ�����' Else '����' END as ORDERTYPE ," & _
"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
" from erpintegration2.ib_wohistory a where  modifyflag='1' "
          
          
sql3 = " order by a.ordername  "

  
  If productTemp <> "" Then
  
  sql2 = " and a.product='" + productTemp + "'"
  
  End If
  
  
 If woTemp <> "" Then
  
  sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
  End If
  
  If Trim(sql2) <> "" Then
  
  sqlTemp = sql1 & sql2 & sql3
  
  Else
  
  sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
   sqlTemp = sql1 & sql2 & sql3
  
  End If
  
  
  
  



Set reportRS = GetPMCWOHeader(sqlTemp)

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With





End Sub

Private Sub ComQueryLine_Click()
'Line


Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim sql1  As String

Dim sql2 As String

Dim sql3 As String

sql1 = ""
sql2 = ""
sql3 = ""


woTemp = UCase(Trim(TxtWo.Text))
productTemp = UCase(Trim(TxtProduct.Text))
beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")

'sql1 = " select a.customer, a.ordername,a.ordertype,a.product,a.para1 as Ƭ�� , a.qty,a.planstartdate,a.planenddate,a.erpuser,a.erpcreatedate ,b.WaferId , b.DieQty " & _
'          " from  erpintegration2.wlo_ib_workorder a, erpintegration2.WLO_IB_WAFERLIST b where  a.OrderName = b.OrderName "
'
          
'  sql1 = "  select seq_ibwo,CUSTOMER ,ORDERNAME , " & _
'"  CASE ORDERTYPE  WHEN '1' THEN 'һ�㹤��'  WHEN '5' THEN '�ټӹ�����'   WHEN '7' THEN 'ί�⹤��'   WHEN '8' THEN '�ع�ί�⹤��' " & _
'" WHEN '11' THEN '���ʽ����'    WHEN '13' THEN 'Ԥ�⹤��'   WHEN '15' THEN '�Բ�����' Else '����' END as ORDERTYPE ," & _
'"  PRODUCT ,QTY,ERPCREATEDATE,PLANSTARTDATE ,PLANENDDATE ,SALESORDER, PARA5,  CUSTOMERPN ,FABFACILITY,IMAGERREV,DESIGNID,MLEVEL235,MLEVEL260 ,NGFLAG,PARA1," & _
'" PARA2,PARA3,PARA4, LOT_STATUS,MPN,PROTECTIVE_FILM_APLD,PARA7,PARA6,CUSTOMER ,to_char(ERPCREATEDATE,'YYYY')||to_char(ERPCREATEDATE,'ww') as datecode " & _
'" from erpintegration2.ib_wohistory a where  modifyflag='1' "


  sql1 = " select b.wafersequence, b.ordername,b.waferid,b.completeflag,b.dieqty, b.FGDieQty , b.WaferLot, b.MarkingCode " & _
"from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername"

          
          
sql3 = " order by a.ordername ,b.WaferLot,b.waferid "

  
  If productTemp <> "" Then
  
  sql2 = " and a.product='" + productTemp + "'"
  
  End If
  
  
 If woTemp <> "" Then
  
  sql2 = sql2 & " and a.ORDERNAME='" + woTemp + "'"
  
  End If
  
  If Trim(sql2) <> "" Then
  
  sqlTemp = sql1 & sql2 & sql3
  
  Else
  
  sql2 = " and a.erpcreatedate>=to_date('" + beginTime + "','YYYY/MM/DD') and a.erpcreatedate<to_date('" + endTime + "','YYYY/MM/DD')+1"
  
   sqlTemp = sql1 & sql2 & sql3
  
  End If
  
  
  
  



Set mainItemRS = GetPMCWOLine(sqlTemp)

With fps(1)
        .MaxRows = 0
        If mainItemRS.RecordCount > 0 Then
            Set .DataSource = mainItemRS
       
        End If
End With


End Sub

Private Sub Form_Activate()
'CmbLine.Text = "TSV"

IniFpsHeader1
'IniFpsHeader2

End Sub
'Private Sub IniProduct()
'Set mainItemRS = GetProduct()
'Set Text3.RowSource = mainItemRS
'Text3.ListField = mainItemRS("productname").Name
'Text3.BoundColumn = mainItemRS("PID").Name
'
'End Sub


Private Sub TabStrip1_Click()

End Sub

Private Sub IniFpsHeader1()
    With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        '�]�m�榡
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        
.SetText E_FPS0.E_ID, 0, "���"
.SetText E_FPS0.E_Factory, 0, "��ҵ��"
.SetText E_FPS0.E_Customer, 0, "�ͻ�����"
.SetText E_FPS0.E_ToDayQty, 0, "�������Ƭ��"
.SetText E_FPS0.E_Mon1, 0, "1�·�"
.SetText E_FPS0.E_Mon2, 0, "2�·�"
.SetText E_FPS0.E_Mon3, 0, "3�·�"
.SetText E_FPS0.E_Mon4, 0, "4�·�"
.SetText E_FPS0.E_Mon5, 0, "5�·�"
.SetText E_FPS0.E_Mon6, 0, "6�·�"

.SetText E_FPS0.E_Mon7, 0, "7�·�"
.SetText E_FPS0.E_Mon8, 0, "8�·�"
.SetText E_FPS0.E_Mon9, 0, "9�·�"
.SetText E_FPS0.E_Mon10, 0, "10�·�"
.SetText E_FPS0.E_Mon11, 0, "11�·�"
.SetText E_FPS0.E_Mon12, 0, "12�·�"
.SetText E_FPS0.E_sum, 0, "Total"


        .ColWidth(E_FPS0.E_ID) = 5
        .ColWidth(E_FPS0.E_Factory) = 12
        .ColWidth(E_FPS0.E_Customer) = 8
        .ColWidth(E_FPS0.E_ToDayQty) = 8
        
        .ColWidth(E_FPS0.E_Mon1) = 8
        .ColWidth(E_FPS0.E_Mon2) = 8
        .ColWidth(E_FPS0.E_Mon3) = 8
        .ColWidth(E_FPS0.E_Mon4) = 8
        .ColWidth(E_FPS0.E_Mon5) = 8
        .ColWidth(E_FPS0.E_Mon6) = 8
        .ColWidth(E_FPS0.E_Mon7) = 8
        .ColWidth(E_FPS0.E_Mon8) = 8
        .ColWidth(E_FPS0.E_Mon9) = 8
        .ColWidth(E_FPS0.E_Mon10) = 8
        .ColWidth(E_FPS0.E_Mon11) = 8
        .ColWidth(E_FPS0.E_Mon12) = 8
        .ColWidth(E_FPS0.E_sum) = 8
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub

Private Sub IniFpsHeader2()
    With fps(1)
        .ReDraw = False
        .MaxCols = E_FPS1.E_End - 1
        .MaxRows = 0
        
        '�]�m�榡
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
          
        .SetText E_FPS1.E_ID, 0, "���"
        .SetText E_FPS1.E_Wo, 0, "������"
        .SetText E_FPS1.E_WaferID, 0, "WaferId"
        .SetText E_FPS1.E_CompleteFlag, 0, "��ɱ�־"
        .SetText E_FPS1.E_TotalDie, 0, "TotalDie����"
        .SetText E_FPS1.E_GoodDie, 0, "GoodDie����"
        .SetText E_FPS1.E_WaferLot, 0, "WaferLot"
        .SetText E_FPS1.E_MarkingCode, 0, "MarkingCode"
        
        
        .ColWidth(E_FPS1.E_ID) = 10
        .ColWidth(E_FPS1.E_Wo) = 10
        .ColWidth(E_FPS1.E_WaferID) = 10
        .ColWidth(E_FPS1.E_CompleteFlag) = 10
        .ColWidth(E_FPS1.E_TotalDie) = 10
        .ColWidth(E_FPS1.E_GoodDie) = 10
        .ColWidth(E_FPS1.E_WaferLot) = 10
        .ColWidth(E_FPS1.E_MarkingCode) = 10
        
     

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub



Private Sub Form_Load()
'IniProduct
End Sub

Private Sub Label9_Click()

End Sub

