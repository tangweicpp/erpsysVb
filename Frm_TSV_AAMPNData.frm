VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_TSV_AAMPNData 
   Caption         =   "�����Ϣ��ѯ"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17535
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
   ScaleHeight     =   9975
   ScaleWidth      =   17535
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdStockExport 
      Caption         =   "�ֿ⵼��ר��"
      Height          =   480
      Left            =   14520
      TabIndex        =   24
      Top             =   3240
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTP4 
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111214594
      CurrentDate     =   43825.5
   End
   Begin MSComCtl2.DTPicker DTP3 
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111214594
      CurrentDate     =   43825.625
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0080FFFF&
      Caption         =   "�˳�"
      Height          =   480
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "����"
      Height          =   480
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton CmdQuery 
      BackColor       =   &H00FFFF00&
      Caption         =   "��ѯ"
      Height          =   480
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "�ͻ�����"
      Height          =   1215
      Left            =   480
      TabIndex        =   11
      Top             =   1920
      Width           =   14055
      Begin VB.CheckBox chkStockOnly 
         Caption         =   "ֻ����"
         Height          =   255
         Left            =   8760
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton OptCIS 
         Caption         =   "CIS"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton OptBUMPING 
         Caption         =   "BUMPING"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OptAA 
         Caption         =   "AA�ͻ�"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblLabel4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����"
         Height          =   195
         Left            =   5280
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   720
      End
      Begin MSForms.ComboBox cbCusCode 
         Height          =   375
         Left            =   6360
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3836;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblAddtion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Addtion: ��ⵥ��ѯ����ѡ�ÿͻ�"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.TextBox TxtInBill 
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   1440
      Width           =   3855
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      ItemData        =   "Frm_TSV_AAMPNData.frx":0000
      Left            =   7320
      List            =   "Frm_TSV_AAMPNData.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox TxtBillNo 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111214593
      CurrentDate     =   41424
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111214593
      CurrentDate     =   41424
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   6015
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   3960
      Width           =   21495
      _Version        =   524288
      _ExtentX        =   37915
      _ExtentY        =   10610
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
      SpreadDesigner  =   "Frm_TSV_AAMPNData.frx":0022
      TextTip         =   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ⵥ�ݱ�ţ�"
      Height          =   195
      Left            =   5880
      TabIndex        =   10
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ͣ�"
      Height          =   195
      Left            =   6720
      TabIndex        =   7
      Top             =   360
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������ݱ�ţ�"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼʱ�䣺"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ�䣺"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   900
   End
End
Attribute VB_Name = "Frm_TSV_AAMPNData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum E_FPS0          'Detail�֭�
    e_ID = 0                'id��
    E_ContainerName
    E_QboxName           '���
    E_LOTID              'lot
    E_qty                '����
    E_MPNSeq             'seq
    E_NewLotid           'new lotid
    E_PRODUCT            'product
    E_CustomerProuduct   'CustomerProuduct
    E_BigQbox             'BigBox
    E_QtyInStock          '�������
    E_END
    
    
End Enum

Private Enum E_FPS1
    e_NO = 0
    e_date
    e_OrderName
    E_SmallBoxID
    E_name
    E_partno
    E_BatchPlant
    E_LOTID
    E_WAFERID
    E_GrossDie
    E_GoodDie
    E_Pieces
    E_INsiteNgDie
    E_BigBoxID
    E_SecondCode
    E_QtyInStock          '�������
    E_END
End Enum
        
Private Enum E_FPS_GC
    e_NO = 0
    e_date '���ʱ��
    E_Bond '��˰��Ǳ�
    E_Mpndesc '�ͻ�����
    E_name 'Ʒ��
    E_partno '�Ϻ�
    E_BatchPlant 'LOT�ź�׺
    E_DieByPcs '��Ƭ����=�ͻ����GoodDie/����Ƭ��
    E_Pieces '����Ƭ��
    E_GrossDie '�ͻ����GoodDie
    E_SecondCode '��������
    E_ProductType '��ʽ
    e_ShipNo '��������
    E_ShipTo '������
    E_BigBoxID '�����
    E_SmallBoxID 'С���
    e_OrderName '������
    E_LOTID 'LOT��
    E_WAFERID 'WAFERLIST
    E_GoodDie 'GOODDIE����
    E_INsiteNgDie '����NG
    E_QtyInStock          '�������
    E_END
End Enum


Dim reportRS As New ADODB.Recordset
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


woTemp = UCase(Trim(txtWO.text))
productTemp = UCase(Trim(TxtProduct.text))
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


Dim waferIdTemp As String
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
waferIdTemp = UCase(Trim(txtWaferID.text))


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
"from erpintegration2.ib_wohistory a, ib_waferlist b  where a.modifyflag='1' and b.ordername=a.ordername and b.waferid='" + waferIdTemp + "' "

          
          
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



Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdquery_Click()

    Dim beginTime   As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim sqlTemp     As String

    Dim sql1        As String

    Dim billNoTemp  As String

    Dim lotIdTemp   As String

    Dim bigQboxTemp As String

    Dim productTemp As String

    Dim date1Temp   As String

    Dim date2Temp   As String

    Dim sql2        As String

    Dim sql3        As String

    Dim sql4        As String

    sql1 = ""
    sql2 = ""
    sql3 = ""
    sqlTemp = ""
    sql4 = ""

    '1. ����
    If Trim(TxtBillNo.text) <> "" Then
        Call IniFpsHeader1
        billNoTemp = UCase(Trim(TxtBillNo.text))
           
        sqlTemp = " select X.��� , X.qboxnumber, X.������, Sum(X.����),  substring(X.���,2,9)  as MPN_SEQ, X.newlotid, X.�Ϻ�, X.MPN, X.���2 from ( " & " SELECT distinct b.���,   C.qboxnumber, B.������,B.���̿����,B.����, C.MPN_SEQ,C.newlotid ,B.�Ϻ�,d.mpn,f.��� as ���2 " & " FROM   tblStockMove A ,tblStockMovesub B ,TblQBOXNUMBER_TSVMPN C,dbo.TblQBOXNUMBER_TSV d  ,tblPackTreeInf e ,tblPackTreeInf f " & " Where A.���߱��=1 AND A.�ͻ�����='AA' AND A.��������=1 and A.���ݱ��='" + billNoTemp + "' " & " and b.���ݱ��=a.���ݱ�� and b.������=a.������ and C.qboxnumber=B.���  and d.QBOXNUMBER=C.qboxnumber and e.���=C.qboxnumber and f.���=e.�ϼ����) X " & " group by X.���,  X.qboxnumber, X.������,X.newlotid ,X.�Ϻ�,X.mpn,X.���2 order by X.���2, X.���,X.������ "
           
        Set mainItemRS = GetAAMPNDataSQL(sqlTemp)
        
        With fps(0)
            .MaxRows = 0

            If mainItemRS.RecordCount > 0 Then
                Set .DataSource = mainItemRS
       
            End If

        End With
        
        Exit Sub

    End If

    ' 2. ���
  '  If Trim(TxtInBill.Text) <> "" Or OptBUMPING.Value = True Then
        date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
        date2Temp = Format(DTP2.Value, "YYYY-MM-DD")
        
        billNoTemp = UCase(Trim(TxtInBill.text))
        Dim s1 As String
        s1 = """;""CUSTOMER_PN"""

        If OptAA.Value = True Then  ' AA
            If Trim(TxtInBill.text) = "" Then
                MsgBox "��������ⵥ��!", vbInformation, "��ʾ"
                Exit Sub
            End If
            Call IniFpsHeader1

             sqlTemp = "select X.��ⵥ��� , X.���, X.WAFERNUMBER, Sum(X.NDPW),  substring(X.���,2,9) as MPN_SEQ, X.newlotid, X.productName, X.MPN, X.���2  FROM ( " & _
            " SELECT distinct a.��ⵥ���,c.��� ,d.WAFERNUMBER,d.WAFERSCRIBENUMBER,d.NDPW,e.MPN_SEQ,e.newlotid,d.PRODUCTNAME,d.MPN ,b.��� as ���2 " & _
            " FROM dbo.tblPackToHouseSub a ,tblPackTreeInf b ,tblPackTreeInf c ,dbo.TblQBOXNUMBER_TSV d,TblQBOXNUMBER_TSVMPN e " & _
            " WHERE a.��ⵥ���='" & billNoTemp & "' and b.���=a.��� and c.�ϼ����=b.��� and d.QBOXNUMBER=c.��� and e.qboxnumber=c.��� ) X " & _
            " group by X.��ⵥ���,X.��� ,X.WAFERNUMBER,X.newlotid,X.PRODUCTNAME,X.MPN ,X.���2 " & _
            " UNION  " & _
            " SELECT xx.��ⵥ���,xx.��� ,xx.������ ,xx.qty ,xx.AA_Q ,SUBSTRING(yy.Content,CHARINDEX('CUSTOMER_LOT_COMPLEX"",""',yy.Content) + 23,10 ),xx.PRODUCT,xx.MPN  ,xx.BBOX FROM ( " & _
            " SELECT  x.��ⵥ���,x.���,x.������,x.qty,x.AA_Q,x.KEYID ,X.PRODUCT,  X.MPN,X.BBOX, MAX(y.Createdate) AS createdate FROM ( " & _
            "  SELECT a.��ⵥ���,b.���,b.������,SUM(b.����) AS qty , SUBSTRING(b.���,2,9)  AS AA_Q  ,c.KEYID ,e.PRODUCT,  e.MPN,aa.��� AS BBOX " & _
            "  FROM erpdata..tblPackTreeInf a  ,erpdata..tblPackMainInfSub b ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblTSVworkorder e ,erpdata..tblPackTreeInf aa " & _
            "   WHERE a.��ⵥ��� = '" & billNoTemp & "'  AND b.��� = a.��� AND c.KEY_VALUE = b.���  AND e.ORDERNAME = b.�󹤵�  AND aa.��� = a.�ϼ���� " & _
            "  GROUP BY a.��ⵥ���,b.���,b.������,c.KEYID,e.PRODUCT,  e.MPN,aa.��� ) X  left JOIN   erpdata..tblME_PrintInfo y ON  y.EVENT_ID = x.KEYID  AND y.LABEL_ID = 'AAMPN4' " & _
            "  GROUP BY x.��ⵥ���,x.���,x.������,x.qty,x.AA_Q,x.KEYID,X.PRODUCT,  X.MPN,X.BBOX ) xx " & _
            "  left JOIN erpdata..tblME_PrintInfo yy ON  yy.EVENT_ID = xx.KEYID  AND yy.LABEL_ID = 'AAMPN4' AND yy.Createdate = xx.createdate  ORDER BY X.PRODUCTNAME "
             
 
            Set mainItemRS = GetAAMPNDataSQL(sqlTemp)
            
            With fps(0)
                .MaxRows = 0

                If mainItemRS.RecordCount > 0 Then
                    Set .DataSource = mainItemRS
       
                End If

            End With
            
            Exit Sub

        ElseIf OptCIS.Value = True Then ' CIS
            date1Temp = Format(DTP1.Value, "YYYY-MM-DD") & " " & Format(DTP3.Value, "hh:mm:ss")
            date2Temp = Format(DTP2.Value, "YYYY-MM-DD") & " " & Format(DTP4.Value, "hh:mm:ss")

            Call delQboxNoTw
            Call InsertQboxTemp(billNoTemp)
            If UCase(Trim(cbCusCode.text)) = "GC" Then ' GC
                Call IniFpsHeader_GC
                sqlTemp = " SELECT x.����ʱ�� AS ���ʱ��,CASE LEFT(x.�󹤵�,1) WHEN 'A'THEN '��˰' ELSE '�Ǳ�˰'END AS ��˰��Ǳ�,x.Mpn_desc AS �ͻ�����,x.QTECHPTNO AS Ʒ�� ,x.�Ϻ� AS �Ϻ�," & _
                        " case CHARINDEX(',', x.WaferId)  when  0 then Rtrim(x.������)+'.' +x.WaferId  else Rtrim(x.������) +'.' + left(x.WaferId,CHARINDEX(',', x.WaferId)-1) End AS LOT�ź�׺, " & _
                         " SUM(x.gross) /sum(x.price) AS ��Ƭ���� , sum(x.price) AS ����Ƭ��,SUM(x.gross) as '�ͻ����GoodDie', x.IMAGER_CUSTOMER_REV AS �������� , isnull(z.�Ƴ�,'') AS ��ʽ,'' AS ��������,'' AS ������,x.���� AS �����,x.��� AS С���, x.�󹤵� AS ������ ,x.������ AS LOT��, x.WaferId AS WAFERLIST  FROM ( " & _
                        " SELECT CONVERT(VARCHAR(20), c.����ʱ��, 23) AS ����ʱ��, b.�󹤵�,f.Mpn_desc,  b.���, g.QTECHPTNO, b.�Ϻ�, SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12) AS SFC, b.������, " & _
                        "  WaferId = (STUFF((SELECT ',' + SUBSTRING(REPLACE(b1.���̿����, '+', ''), len(REPLACE(b1.���̿����, '+', '')) - 1, 2)  FROM erpdata .. tblPackMainInfSub b1 " & _
                        "  WHERE b.��� = b1.��� and  b.�󹤵� = b1.�󹤵� order by b1.���̿���� FOR XML PATH('')), 1,  1, '')),  e.PASSBINCOUNT + e.FAILBINCOUNT AS gross,  SUM(CONVERT(INT, h1.KEY_VALUE)) AS good_die, " & _
                        "  COUNT(DISTINCT b.���̿����) AS price,  SUM(CONVERT(INT, h2.KEY_VALUE)) AS ng_die, a1.��� AS ����,  f.IMAGER_CUSTOMER_REV, b.���̿���� ,isnull(k.����,'') as ��� FROM erpdata .. tblPackTreeInf a " & _
                        " INNER JOIN erpdata .. tblPackMainInfSub b  ON b.��� = a.��� " & _
                        " INNER JOIN ERPBASE .. tblmappingData e ON e.SUBSTRATEID = b.���̿���� " & _
                        " INNER JOIN ERPBASE .. tblCustomerOI f  ON  convert(varchar(50),convert(int,f.ID)) = e.FILENAME and f.SOURCE_BATCH_ID = e.LOTID " & _
                        " INNER JOIN erptemp .. tbltsvnpiproduct g ON g.QTECHPTNO2 = b.�Ϻ�  AND g.CUSTOMERPTNO1 = f.MPN_DESC  AND g.CUSTOMERSHORTNAME = f.CUSTOMERSHORTNAME " & _
                        " INNER JOIN erpdata .. tblErpInStockDetailInfo h  ON h.KEY_VALUE = a.���  " & _
                        " INNER JOIN erpdata .. tblErpInStockRelation j ON j.BOX_ID = h.BOX_ID  AND SUBSTRING(REPLACE(j.WAFER_ID, j.SFC_ID, ''),  2, CHARINDEX('::', REPLACE(j.WAFER_ID, j.SFC_ID, '')) - 2) =  b.���̿���� " & _
                        " LEFT join erpdata .. tblErpInStockDetailInfo h1  ON h1.BOX_ID = h.BOX_ID  AND h1.KEY_NAME = 'GOOD_DIE' AND h1.KEY_TYPE = 'WAFER' AND H1.KEYID = J.WAFER_ID AND h1.KEY_VALUE <> '0' " & _
                        " LEFT join erpdata .. tblPackTreeInf a1  ON a1.��� = a.�ϼ���� " & _
                        " INNER JOIN erpdata .. tblPackMainInf c  ON c.��� = a1.��� " & _
                        " LEFT join erpdata .. tblErpInStockDetailInfo h2   ON h2.BOX_ID = h.BOX_ID AND h2.KEY_NAME IN ('BAD1_DIE', 'BAD2_DIE') AND h2.KEY_TYPE = 'WAFER' AND H2.KEYID = J.WAFER_ID  AND h2.KEY_VALUE <> '0' "
                        
                        If chkStockOnly.Value = 1 Then '��ֻ����
                            sqlTemp = sqlTemp & "  inner JOIN erpdata..Tblstocknumsub k ON b.���̿���� =k.���̿����   where k.�ⷿ���<>72 "
                            If Trim(TxtInBill.text) <> "" Then
                                sqlTemp = sqlTemp & "  and  f.CUSTOMERSHORTNAME='GC'  and  a.��ⵥ��� = '" & billNoTemp & " ' "
                            Else
                                sqlTemp = sqlTemp & " and  f.CUSTOMERSHORTNAME='GC'  and  c.����ʱ�� >= '" & date1Temp & "' and c.����ʱ�� <= '" & date2Temp & " '  AND isnull(a.��ⵥ���,'')<>'' "
                            End If
                        Else '���е�
                            sqlTemp = sqlTemp & "  left JOIN erpdata..Tblstocknumsub k ON b.���̿���� =k.���̿���� and b.��� =k.��� "
                            If Trim(TxtInBill.text) <> "" Then
                                sqlTemp = sqlTemp & "  where  f.CUSTOMERSHORTNAME='GC'  and  a.��ⵥ��� = '" & billNoTemp & " ' "
                            Else
                                sqlTemp = sqlTemp & " where  f.CUSTOMERSHORTNAME='GC'  and  c.����ʱ�� >= '" & date1Temp & "' and c.����ʱ�� <= '" & date2Temp & " '  AND isnull(a.��ⵥ���,'')<>'' "
                            End If
                        End If

                        
                        sqlTemp = sqlTemp & " GROUP BY CONVERT(VARCHAR(20), c.����ʱ��, 23), b.�󹤵�, b. ���, g.QTECHPTNO, b.�Ϻ�,b.���̿����,  SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12), b.������, e.PASSBINCOUNT + e.FAILBINCOUNT,  a1.���,  f.IMAGER_CUSTOMER_REV,f.Mpn_desc,k.���� ) x " & _
                        " LEFT JOIN  erptemp..GcCode_Reference  z ON  x.Mpn_desc=z.�ͻ������� AND x.�Ϻ�=z.��Ʒ�Ϻ�  AND  (SUBSTRING(x.IMAGER_CUSTOMER_REV,3,1)=z.�������� or SUBSTRING(x.IMAGER_CUSTOMER_REV,3,1)=z.��bin�������� ) " & _
                        "  GROUP BY  x.����ʱ��,x.�󹤵�,x.���,x.QTECHPTNO,x.�Ϻ�,x.SFC,x.������,x.WaferId ,x.����,x.IMAGER_CUSTOMER_REV,x.Mpn_desc,z.�Ƴ� ORDER BY x.QTECHPTNO "
', SUM(x.good_die) AS GOODDIE���� ,SUM(x.ng_die) AS ����NG, x.���
',x.���

            Else
                Call IniFpsHeader2
               ' sqlTemp = "select waredate,workordername,qboxnumber,alternatename,productname,firstname,wafernumber,TO_CHAR(wmsys.wm_concat(substr(waferscribenumber, -2, 2))) As WaferList," & "sum(gross_dies),sum(gross_dies) - sum(ng_dies),count(waferscribenumber),sum(ng_dies),outpack,imager_customer_rev from cis_in_report group by waredate,workordername," & "qboxnumber,alternatename,productname,firstname,wafernumber,outpack,imager_customer_rev order by wafernumber"
                  
               sqlTemp = " SELECT x.����ʱ��,x.�󹤵�,x.���,x.QTECHPTNO,x.�Ϻ�," & _
                        " case CHARINDEX(',', x.WaferId)  when  0 then Rtrim(x.������)+'.' +x.WaferId  else Rtrim(x.������) +'.' + left(x.WaferId,CHARINDEX(',', x.WaferId)-1) End ,x.������, " & _
                        " x.WaferId,SUM(x.gross), SUM(x.good_die), sum(x.price),SUM(x.ng_die),x.����,x.IMAGER_CUSTOMER_REV  FROM ( " & _
                        " SELECT CONVERT(VARCHAR(20), c.����ʱ��, 23) AS ����ʱ��, b.�󹤵�,  b.���, g.QTECHPTNO, b.�Ϻ�, SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12) AS SFC, b.������, " & _
                        "  WaferId = (STUFF((SELECT ',' + SUBSTRING(REPLACE(b1.���̿����, '+', ''), len(REPLACE(b1.���̿����, '+', '')) - 1, 2)  FROM erpdata .. tblPackMainInfSub b1 " & _
                        "  WHERE b.��� = b1.��� and  b.�󹤵� = b1.�󹤵� order by b1.���̿���� FOR XML PATH('')), 1,  1, '')),  e.PASSBINCOUNT + e.FAILBINCOUNT AS gross,  SUM(CONVERT(INT, h1.KEY_VALUE)) AS good_die, " & _
                        "  COUNT(DISTINCT b.���̿����) AS price,  SUM(CONVERT(INT, h2.KEY_VALUE)) AS ng_die, a1.��� AS ����,  f.IMAGER_CUSTOMER_REV, b.���̿����  , k.���� as ���  FROM erpdata .. tblPackTreeInf a " & _
                        " INNER JOIN erpdata .. tblPackMainInfSub b  ON b.��� = a.��� INNER JOIN erpdata .. tblPackMainInf c  ON c.��� = b.��� INNER JOIN ERPBASE .. tblmappingData e ON e.SUBSTRATEID = b.���̿���� " & _
                        " INNER JOIN ERPBASE .. tblCustomerOI f  ON  convert(varchar(50),convert(int,f.ID)) = e.FILENAME and f.SOURCE_BATCH_ID = e.LOTID INNER JOIN erptemp .. tbltsvnpiproduct g ON g.QTECHPTNO2 = b.�Ϻ� " & _
                        "   AND g.CUSTOMERPTNO1 = f.MPN_DESC  AND g.CUSTOMERSHORTNAME = f.CUSTOMERSHORTNAME INNER JOIN erpdata .. tblErpInStockDetailInfo h  ON h.KEY_VALUE = a.��� INNER JOIN erpdata .. tblErpInStockRelation j " & _
                        "  ON j.BOX_ID = h.BOX_ID  AND SUBSTRING(REPLACE(j.WAFER_ID, j.SFC_ID, ''),  2, CHARINDEX('::', REPLACE(j.WAFER_ID, j.SFC_ID, '')) - 2) =  b.���̿����  LEFT join erpdata .. tblErpInStockDetailInfo h1 " & _
                        "  ON h1.BOX_ID = h.BOX_ID  AND h1.KEY_NAME = 'GOOD_DIE' AND h1.KEY_TYPE = 'WAFER' AND H1.KEYID = J.WAFER_ID AND h1.KEY_VALUE <> '0'  LEFT join erpdata .. tblPackTreeInf a1  ON a1.��� = a.�ϼ���� " & _
                        "  LEFT join erpdata .. tblErpInStockDetailInfo h2   ON h2.BOX_ID = h.BOX_ID AND h2.KEY_NAME IN ('BAD1_DIE', 'BAD2_DIE') AND h2.KEY_TYPE = 'WAFER' AND H2.KEYID = J.WAFER_ID  AND h2.KEY_VALUE <> '0' "
                If chkStockOnly.Value = 1 Then '��ֻ����
                    sqlTemp = sqlTemp & "  inner JOIN erpdata..Tblstocknumsub k ON b.���̿���� =k.���̿����  where k.�ⷿ���<>72 "
                    If Trim(TxtInBill.text) <> "" Then
                        sqlTemp = sqlTemp & "  and   a.��ⵥ��� = '" & billNoTemp & " ' "
                    Else
                        sqlTemp = sqlTemp & "  and  c.����ʱ�� >= '" & date1Temp & "' and c.����ʱ�� <= '" & date2Temp & " '  AND isnull(a.��ⵥ���,'')<>'' "
                    End If
                Else '���е�
                    sqlTemp = sqlTemp & "  left JOIN erpdata..Tblstocknumsub k ON b.���̿���� =k.���̿���� and  b.��� =k.��� "
                    If Trim(TxtInBill.text) <> "" Then
                        sqlTemp = sqlTemp & "  Where  a.��ⵥ��� = '" & billNoTemp & " ' "
                    Else
                        sqlTemp = sqlTemp & "  Where  c.����ʱ�� >= '" & date1Temp & "' and c.����ʱ�� <= '" & date2Temp & " '  AND isnull(a.��ⵥ���,'')<>'' "
                    End If
                                        
                End If

                If Trim(cbCusCode.text) <> "" Then
                    sqlTemp = sqlTemp & " and f.CUSTOMERSHORTNAME='" & Trim(cbCusCode.text) & "' "
                End If
                sqlTemp = sqlTemp & " GROUP BY CONVERT(VARCHAR(20), c.����ʱ��, 23), b.�󹤵�, b. ���, g.QTECHPTNO, b.�Ϻ�,b.���̿����,  SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12), b.������, e.PASSBINCOUNT + e.FAILBINCOUNT,  a1.���,  f.IMAGER_CUSTOMER_REV , k.����) x " & _
                        "  GROUP BY  x.����ʱ��,x.�󹤵�,x.���,x.QTECHPTNO,x.�Ϻ�,x.SFC,x.������,x.WaferId ,x.����,x.IMAGER_CUSTOMER_REV  order by x.QTECHPTNO"
            End If
                                      
            Set mainItemRS = getSqlStr(sqlTemp)
            
            With fps(0)
                .MaxRows = 0

                If mainItemRS.RecordCount > 0 Then
                    Set .DataSource = mainItemRS
       
                End If

            End With
            
            Exit Sub

        ElseIf OptBUMPING.Value = True Then ' BUMPING
            date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
            date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")

            If billNoTemp = "" Then
                sqlTemp = "select X.���ʱ�� ,X.��ⵥ���,X.ORDERNAME as ������,X.���,X.SALESORDER as ����,X.PRODUCT as �Ϻ�,X.SFC as ��������,X.������ as �ͻ�LOT,X.���̿���� as WAFER_ID,CONVERT(INT, X.DIEQTY)  - CONVERT(INT, Z.KEY_VALUE) as ������Ʒ,X.KEY_VALUE as �Ƴ���Ʒ,Y.KEY_VALUE as �Ƴ̲���Ʒ,Z.KEY_VALUE as ���ϲ���" & _
                   ",X.MARKINGCODE as �����   from (select CONVERT(varchar(100),aa.���ʱ��,23) as ���ʱ��,a.��ⵥ���,g.ORDERNAME,a.���,g.SALESORDER,g.PRODUCT,replace(d.SFC_ID,'SFCBO:1020,','') as SFC,b.������ " & _
                   ",b.���̿����,e.KEY_VALUE,f.DIEQTY,F.MARKINGCODE,e.KEYID from erpdata..tblPackTreeInf a ,erpdata..tblPackToHouse aa ,erpdata..tblPackMainInfSub b  ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblErpInStockRelation d " & _
                   ",erpdata..tblErpInStockDetailInfo e,erpdata..tblTSVwaferlist f,erpdata..tblTSVworkorder g where aa.���ʱ�� >= '" & date1Temp & "' and aa.���ʱ�� <= '" & date2Temp & "' and a.�ϼ���� <> '0' and aa.��ⵥ��� = a.��ⵥ��� and b.��� = a.��� " & _
                   "and c.KEY_VALUE = b.���  and d.BOX_ID = c.BOX_ID and SUBSTRING(replace(d.WAFER_ID,'SFCBO:1020,',''), CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))+1,CHARINDEX('::',replace(d.WAFER_ID,'SFCBO:1020,',''))-CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))-1) = b.���̿���� " & _
                   "and e.KEYID = d.WAFER_ID and e.KEY_NAME in ('GOOD_DIE' ) and f.WAFERID = b.���̿���� and g.ORDERNAME = f.ORDERNAME and g.ORDERNAME = aa.�󹤵� ) x,erpdata..tblErpInStockDetailInfo y,erpdata..tblErpInStockDetailInfo z where y.KEYID = x.KEYID " & _
                   "and y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' order by x.���ʱ��,x.��ⵥ���, x.���̿����"
            Else
                Dim rkdID As String
                rkdID = billNoTemp
                 sqlTemp = " SELECT X.���ʱ�� ,X.��ⵥ���,X.ORDERNAME AS ������,X.���, X.HT_DEVICE AS Ʒ��,X.PRODUCT AS �Ϻ�,X.SFC AS ��������,X.������ AS LOT_ID, " & _
                        " X.���̿���� AS WAFER_ID,CONVERT(INT, X.DIEQTY) AS �ͻ����GOODDIE,X.KEY_VALUE AS GOODDIE����,1 AS ��Ʒ��վƬ��, Y.KEY_VALUE AS ����NGDIE����, " & _
                        " Z.KEY_VALUE AS �ͻ�NGDIE����,X.MARKINGCODE AS ����������,X.CUSTOMER AS �ͻ����� FROM (SELECT CONVERT(VARCHAR(100), aa.���ʱ��, 23) AS ���ʱ��, " & _
                        " a.��ⵥ���, g.ORDERNAME, a.���,h.HT_DEVICE,g.PRODUCT,REPLACE(d.SFC_ID, 'SFCBO:1020,', '') AS SFC,b.������, b.���̿����,e.KEY_VALUE,f.DIEQTY, " & _
                        " F.MARKINGCODE,e.KEYID,g.CUSTOMER FROM erpdata .. tblPackTreeInf a,erpdata .. tblPackToHouse aa,erpdata .. tblPackMainInfSub b, erpdata .. tblErpInStockDetailInfo c, " & _
                        " erpdata .. tblErpInStockRelation d,erpdata .. tblErpInStockDetailInfo e,erpdata .. tblTSVwaferlist f,erpdata .. tblTSVworkorder g,erpdata .. SHOP_ORDER h " & _
                        " WHERE a.�ϼ���� <> '0'  AND aa.��ⵥ��� = a.��ⵥ��� AND b.��� = a.��� and a.��ⵥ��� = '" & billNoTemp & "' AND c.KEY_VALUE = b.��� AND h.SHOP_ORDER = g.ORDERNAME " & _
                        " AND d.BOX_ID = c.BOX_ID AND SUBSTRING(REPLACE(d.WAFER_ID, 'SFCBO:1020,', ''), CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) + 1, CHARINDEX('::',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - " & _
                        " CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - 1) =b.���̿���� AND e.KEYID = d.WAFER_ID AND e.KEY_NAME IN ('GOOD_DIE') AND f.WAFERID = b.���̿���� AND g.ORDERNAME = f.ORDERNAME) x, " & _
                        " erpdata .. tblErpInStockDetailInfo y, erpdata .. tblErpInStockDetailInfo z WHERE y.KEYID = x.KEYID AND y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' "

            End If

            Set mainItemRS = GetAAMPNDataSQL(sqlTemp)
            
            With fps(0)
                .MaxRows = 0

                If mainItemRS.RecordCount > 0 Then
                    Set .DataSource = mainItemRS
       
                End If

            End With
            
            Exit Sub

        Else
            Call IniFpsHeader1
            Exit Sub

        End If
    
   ' End If
    
    ' 3.�����LOTID����
    If Trim(TxtInBill.text) = "" And Trim(TxtBillNo.text) = "" Then
        Call IniFpsHeader1
        productTemp = UCase(Trim(CmbType.text))
     
        date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
        date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
     
        If productTemp = "���LOTID����" Then
     
            sqlTemp = "select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW), substr(X.QBOXNUMBER,2,9)  as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn ,''  from ( " & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID ,a.productname, a.mpn ,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' and a.QBOXNUMBER like 'QM%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"
     
        Else
     
            sqlTemp = " select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW),substr(X.QBOXNUMBER,2,9) as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn ,''  from (" & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID, a.productname, a.mpn,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"

        End If

        Set mainItemRS = GetAAMPNData(sqlTemp)
        
        With fps(0)
            .MaxRows = 0

            If mainItemRS.RecordCount > 0 Then
                Set .DataSource = mainItemRS
       
            End If

        End With

    End If

End Sub



Private Sub InitCuscode()
Dim rs As New ADODB.Recordset, i As Integer

Set rs.ActiveConnection = SqlConnect
rs.Source = "select distinct �ͻ����� from tblxcustomer"
rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
cbCusCode.Clear
If rs.RecordCount > 0 Then
    rs.MoveFirst

    For i = 1 To rs.RecordCount
        cbCusCode.AddItem Trim(rs("�ͻ�����"))
        rs.MoveNext
    Next i

End If

End Sub

Private Sub cmdStockExport_Click()
'    SqlServer2ExporToExcel ("select AA.�󹤵�,CC.QTY as ������,BB.�����, AA.������ from(select �󹤵�,SUM(����) as ������ from ERPDATA.. TBLSTOCKMOVESUB where �󹤵� is not null group by �󹤵�)AA, " & _
'" (select �󹤵�,SUM(�����) as ����� from dbo.tblPackToHouse  group by �󹤵�) BB, " & _
'" dbo.[tblTSVworkorder] CC Where aa.�󹤵� = BB.�󹤵� and AA.�󹤵� = CC.ORDERNAME and CC.CUSTOMER = '95' and AA.�󹤵� not like '%-16%'")

    Dim beginTime   As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim sqlTemp     As String

    Dim sql1        As String

    Dim billNoTemp  As String

    Dim lotIdTemp   As String

    Dim bigQboxTemp As String

    Dim productTemp As String

    Dim date1Temp   As String

    Dim date2Temp   As String

    Dim sql2        As String

    Dim sql3        As String

    sql1 = ""
    sql2 = ""
    sql3 = ""
    sqlTemp = ""

    If Trim(TxtBillNo.text) <> "" Then   '����

        billNoTemp = UCase(Trim(TxtBillNo.text))

        sqlTemp = " select X.��� , X.qboxnumber, X.������, Sum(X.����) as ����, substring(X.���,2,9) as MPN_SEQ, X.newlotid, X.�Ϻ�, X.MPN as OPN, X.���2 as ����� from ( " & " SELECT distinct b.���,   C.qboxnumber, B.������,B.���̿����,B.����, C.MPN_SEQ,C.newlotid ,B.�Ϻ�,d.mpn,f.��� as ���2 " & " FROM   tblStockMove A ,tblStockMovesub B ,TblQBOXNUMBER_TSVMPN C,dbo.TblQBOXNUMBER_TSV d  ,tblPackTreeInf e ,tblPackTreeInf f " & " Where A.���߱��=1 AND A.�ͻ�����='AA' AND A.��������=1 and A.���ݱ��='" + billNoTemp + "' " & " and b.���ݱ��=a.���ݱ�� and b.������=a.������ and C.qboxnumber=B.���  and d.QBOXNUMBER=C.qboxnumber and e.���=C.qboxnumber and f.���=e.�ϼ����) X " & " group by X.���,  X.qboxnumber, X.������,X.newlotid ,X.�Ϻ�,X.mpn,X.���2 order by X.���2, X.���,X.������ "

        SqlServer2ExporToExcel_Trim (sqlTemp)
        Exit Sub

    End If
    If OptCIS.Value = True Then   ' CIS ��Ϊֱ����FPS���EXCEL
        FpsToExcel
        Exit Sub
    End If
    

    If Trim(TxtInBill.text) <> "" Or OptBUMPING.Value = True Then

        billNoTemp = UCase(Trim(TxtInBill.text))

        If OptAA.Value = True Then ' AA
 
             sqlTemp = "select X.��ⵥ��� , X.���, X.WAFERNUMBER, Sum(X.NDPW),  substring(X.���,2,9) as MPN_SEQ, X.newlotid, X.productName, X.MPN,x.���� as  ���� , X.���2,  '' as �����,'' as ���� ,'' as ����  ,x.���� as ë�� ,x.�ߴ� as ���,'' as ���� FROM ( " & _
            " SELECT distinct a.��ⵥ���,c.��� ,d.WAFERNUMBER,d.WAFERSCRIBENUMBER,d.NDPW,e.MPN_SEQ,e.newlotid,d.PRODUCTNAME,d.MPN ,b.��� as ���2 ,f.�ߴ�,f.����,g.���� " & _
            " FROM dbo.tblPackToHouseSub a ,tblPackTreeInf b ,tblPackTreeInf c ,dbo.TblQBOXNUMBER_TSV d,TblQBOXNUMBER_TSVMPN e ,erpdata..tblStockNumTree f,erpdata..tblWeight_AA g " & _
            " WHERE a.��ⵥ���='" & billNoTemp & "' and b.���=a.��� and c.�ϼ����=b.��� and d.QBOXNUMBER=c.��� and e.qboxnumber=c.��� and  f.���=b.��� and g.�Ϻ�=d.productName ) X " & _
            " group by X.��ⵥ���,X.��� ,X.WAFERNUMBER,X.newlotid,X.PRODUCTNAME,X.MPN ,X.���2  ,x.�ߴ�,x.����,x.����  " & _
            " UNION  " & _
            " SELECT xx.��ⵥ���,xx.��� ,xx.������ ,xx.qty ,xx.AA_Q ,SUBSTRING(yy.Content,CHARINDEX('CUSTOMER_LOT_COMPLEX"",""',yy.Content) + 23,10),xx.PRODUCT,xx.MPN,mm.���� as  ����   ,xx.BBOX , '' as �����,'' as ���� ,'' as ����  ,zz.���� as ë��,  zz.�ߴ�  as ���,'' as ���� FROM ( " & _
            " SELECT  x.��ⵥ���,x.���,x.������,x.qty,x.AA_Q,x.KEYID ,X.PRODUCT,  X.MPN,X.BBOX, MAX(y.Createdate) AS createdate FROM ( " & _
            "  SELECT a.��ⵥ���,b.���,b.������,SUM(b.����) AS qty , SUBSTRING(b.���,2,9)  AS AA_Q  ,c.KEYID ,e.PRODUCT,  e.MPN,aa.��� AS BBOX " & _
            "  FROM erpdata..tblPackTreeInf a  ,erpdata..tblPackMainInfSub b ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblTSVworkorder e ,erpdata..tblPackTreeInf aa " & _
            "   WHERE a.��ⵥ��� = '" & billNoTemp & "'  AND b.��� = a.��� AND c.KEY_VALUE = b.���  AND e.ORDERNAME = b.�󹤵�  AND aa.��� = a.�ϼ���� " & _
            "  GROUP BY a.��ⵥ���,b.���,b.������,c.KEYID,e.PRODUCT,  e.MPN,aa.��� ) X  left JOIN   erpdata..tblME_PrintInfo y ON  y.EVENT_ID = x.KEYID  AND y.LABEL_ID = 'AAMPN4' " & _
            "  GROUP BY x.��ⵥ���,x.���,x.������,x.qty,x.AA_Q,x.KEYID,X.PRODUCT,  X.MPN,X.BBOX ) xx " & _
            "  LEFT JOIN  erpdata..tblStockNumTree zz ON xx.BBOX=zz.��� " & _
            "  LEFT JOIN  erpdata..tblWeight_AA mm ON xx.PRODUCT=mm.�Ϻ� " & _
            "  left JOIN erpdata..tblME_PrintInfo yy ON  yy.EVENT_ID = xx.KEYID  AND yy.LABEL_ID = 'AAMPN4' AND yy.Createdate = xx.createdate  ORDER BY X.PRODUCTNAME , X.���2"
             

 
            SqlServer2ExporToExcel_Trim_AA (sqlTemp)

        ' ElseIf OptCIS.Value = True Then   ' CIS
                   
            ' sqlTemp = " SELECT x.����ʱ��,x.�󹤵�,x.���,x.QTECHPTNO,x.�Ϻ�, " & _
                    ' " case CHARINDEX(',', x.WaferId)  when  0 then Rtrim(x.������)+'.' +x.WaferId  else Rtrim(x.������) +'.' + left(x.WaferId,CHARINDEX(',', x.WaferId)-1) End ,x.������, " & _
                    ' "x.WaferId,SUM(x.gross), SUM(x.good_die), sum(x.price),SUM(x.ng_die),x.����,x.IMAGER_CUSTOMER_REV FROM ( " & _
                    ' " SELECT CONVERT(VARCHAR(20), c.����ʱ��, 23) AS ����ʱ��, b.�󹤵�,  b.���, g.QTECHPTNO, b.�Ϻ�, SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12) AS SFC, b.������, " & _
                    ' "  WaferId = (STUFF((SELECT ',' + SUBSTRING(REPLACE(b1.���̿����, '+', ''), len(REPLACE(b1.���̿����, '+', '')) - 1, 2)  FROM erpdata .. tblPackMainInfSub b1 " & _
                    ' "  WHERE b.��� = b1.��� and  b.�󹤵� = b1.�󹤵� order by b1.���̿���� FOR XML PATH('')), 1,  1, '')),  e.PASSBINCOUNT + e.FAILBINCOUNT AS gross,  SUM(CONVERT(INT, h1.KEY_VALUE)) AS good_die, " & _
                    ' "  COUNT(DISTINCT b.���̿����) AS price,  SUM(CONVERT(INT, h2.KEY_VALUE)) AS ng_die, a1.��� AS ����,  f.IMAGER_CUSTOMER_REV, b.���̿����  FROM erpdata .. tblPackTreeInf a " & _
                    ' " INNER JOIN erpdata .. tblPackMainInfSub b  ON b.��� = a.��� INNER JOIN erpdata .. tblPackMainInf c  ON c.��� = b.��� INNER JOIN ERPBASE .. tblmappingData e ON e.SUBSTRATEID = b.���̿���� " & _
                    ' " INNER JOIN ERPBASE .. tblCustomerOI f  ON  convert(varchar(50),convert(int,f.ID)) = e.FILENAME and f.SOURCE_BATCH_ID = e.LOTID INNER JOIN erptemp .. tbltsvnpiproduct g ON g.QTECHPTNO2 = b.�Ϻ� " & _
                    ' "   AND g.CUSTOMERPTNO1 = f.MPN_DESC  AND g.CUSTOMERSHORTNAME = f.CUSTOMERSHORTNAME INNER JOIN erpdata .. tblErpInStockDetailInfo h  ON h.KEY_VALUE = a.��� INNER JOIN erpdata .. tblErpInStockRelation j " & _
                    ' "  ON j.BOX_ID = h.BOX_ID  AND SUBSTRING(REPLACE(j.WAFER_ID, j.SFC_ID, ''),  2, CHARINDEX('::', REPLACE(j.WAFER_ID, j.SFC_ID, '')) - 2) =  b.���̿����  LEFT join erpdata .. tblErpInStockDetailInfo h1 " & _
                    ' "  ON h1.BOX_ID = h.BOX_ID  AND h1.KEY_NAME = 'GOOD_DIE' AND h1.KEY_TYPE = 'WAFER' AND H1.KEYID = J.WAFER_ID AND h1.KEY_VALUE <> '0'  LEFT join erpdata .. tblPackTreeInf a1  ON a1.��� = a.�ϼ���� " & _
                    ' "  LEFT join erpdata .. tblErpInStockDetailInfo h2   ON h2.BOX_ID = h.BOX_ID AND h2.KEY_NAME IN ('BAD1_DIE', 'BAD2_DIE') AND h2.KEY_TYPE = 'WAFER' AND H2.KEYID = J.WAFER_ID  AND h2.KEY_VALUE <> '0' " & _
                    ' "  WHERE a.��ⵥ��� = '" & billNoTemp & " ' " & _
                    ' " GROUP BY CONVERT(VARCHAR(20), c.����ʱ��, 23), b.�󹤵�, b. ���, g.QTECHPTNO, b.�Ϻ�,b.���̿����,  SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12), b.������, e.PASSBINCOUNT + e.FAILBINCOUNT,  a1.���,  f.IMAGER_CUSTOMER_REV ) x " & _
                    ' "  GROUP BY  x.����ʱ��,x.�󹤵�,x.���,x.QTECHPTNO,x.�Ϻ�,x.SFC,x.������,x.WaferId ,x.����,x.IMAGER_CUSTOMER_REV"
                                  
           ' ' ExporToExcel (sqlTemp)
           ' SqlServer2ExporToExcel_Trim (sqlTemp)

        
        ElseIf OptBUMPING.Value = True Then ' BUMPING
    
            date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
            date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")

            If TxtInBill.text = "" Then
                sqlTemp = "select X.���ʱ�� ,X.��ⵥ���,X.ORDERNAME as ������,X.���,X.SALESORDER as ����,X.PRODUCT as �Ϻ�,X.SFC as ��������,X.������ as �ͻ�LOT,X.���̿���� as WAFER_ID,CONVERT(INT, X.DIEQTY)  - CONVERT(INT, Z.KEY_VALUE) as ������Ʒ,X.KEY_VALUE as �Ƴ���Ʒ,Y.KEY_VALUE as �Ƴ̲���Ʒ,Z.KEY_VALUE as ���ϲ���" & _
                   ",X.MARKINGCODE as �����   from (select CONVERT(varchar(100),aa.���ʱ��,23) as ���ʱ��,a.��ⵥ���,g.ORDERNAME,a.���,g.SALESORDER,g.PRODUCT,replace(d.SFC_ID,'SFCBO:1020,','') as SFC,b.������ " & _
                   ",b.���̿����,e.KEY_VALUE,f.DIEQTY,F.MARKINGCODE,e.KEYID from erpdata..tblPackTreeInf a ,erpdata..tblPackToHouse aa ,erpdata..tblPackMainInfSub b  ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblErpInStockRelation d " & _
                   ",erpdata..tblErpInStockDetailInfo e,erpdata..tblTSVwaferlist f,erpdata..tblTSVworkorder g where aa.���ʱ�� >= '" & date1Temp & "' and aa.���ʱ�� <= '" & date2Temp & "' and a.�ϼ���� <> '0' and aa.��ⵥ��� = a.��ⵥ��� and b.��� = a.��� " & _
                   "and c.KEY_VALUE = b.���  and d.BOX_ID = c.BOX_ID and SUBSTRING(replace(d.WAFER_ID,'SFCBO:1020,',''), CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))+1,CHARINDEX('::',replace(d.WAFER_ID,'SFCBO:1020,',''))-CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))-1) = b.���̿���� " & _
                   "and e.KEYID = d.WAFER_ID and e.KEY_NAME in ('GOOD_DIE' ) and f.WAFERID = b.���̿���� and g.ORDERNAME = f.ORDERNAME and g.ORDERNAME = aa.�󹤵� ) x,erpdata..tblErpInStockDetailInfo y,erpdata..tblErpInStockDetailInfo z where y.KEYID = x.KEYID " & _
                   "and y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' order by x.���ʱ��,x.��ⵥ���, x.���̿����"
            Else
                Dim rkdID As String
                    
                rkdID = Trim$(TxtInBill.text)
            
                sqlTemp = " SELECT X.���ʱ�� ,X.��ⵥ���,X.ORDERNAME AS ������,X.���, X.HT_DEVICE AS Ʒ��,X.PRODUCT AS �Ϻ�,X.SFC AS ��������,X.������ AS LOT_ID, " & _
                        " X.���̿���� AS WAFER_ID,CONVERT(INT, X.DIEQTY) AS �ͻ����GOODDIE,X.KEY_VALUE AS GOODDIE����,1 AS ��Ʒ��վƬ��, Y.KEY_VALUE AS ����NGDIE����, " & _
                        " Z.KEY_VALUE AS �ͻ�NGDIE����,X.MARKINGCODE AS ����������,X.CUSTOMER AS �ͻ����� FROM (SELECT CONVERT(VARCHAR(100), aa.���ʱ��, 23) AS ���ʱ��, " & _
                        " a.��ⵥ���, g.ORDERNAME, a.���,h.HT_DEVICE,g.PRODUCT,REPLACE(d.SFC_ID, 'SFCBO:1020,', '') AS SFC,b.������, b.���̿����,e.KEY_VALUE,f.DIEQTY, " & _
                        " F.MARKINGCODE,e.KEYID,g.CUSTOMER FROM erpdata .. tblPackTreeInf a,erpdata .. tblPackToHouse aa,erpdata .. tblPackMainInfSub b, erpdata .. tblErpInStockDetailInfo c, " & _
                        " erpdata .. tblErpInStockRelation d,erpdata .. tblErpInStockDetailInfo e,erpdata .. tblTSVwaferlist f,erpdata .. tblTSVworkorder g,erpdata .. SHOP_ORDER h " & _
                        " WHERE a.�ϼ���� <> '0'  AND aa.��ⵥ��� = a.��ⵥ��� AND b.��� = a.��� and a.��ⵥ��� = '" & billNoTemp & "' AND c.KEY_VALUE = b.��� AND h.SHOP_ORDER = g.ORDERNAME " & _
                        " AND d.BOX_ID = c.BOX_ID AND SUBSTRING(REPLACE(d.WAFER_ID, 'SFCBO:1020,', ''), CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) + 1, CHARINDEX('::',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - " & _
                        " CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - 1) =b.���̿���� AND e.KEYID = d.WAFER_ID AND e.KEY_NAME IN ('GOOD_DIE') AND f.WAFERID = b.���̿���� AND g.ORDERNAME = f.ORDERNAME) x, " & _
                        " erpdata .. tblErpInStockDetailInfo y, erpdata .. tblErpInStockDetailInfo z WHERE y.KEYID = x.KEYID AND y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' "

            End If

            SqlServer2ExporToExcel_Trim (sqlTemp)
        
        Else
        
        End If
  
    End If
    
    If Trim(TxtBillNo.text) = "" And Trim(TxtInBill.text) = "" Then

        productTemp = UCase(Trim(CmbType.text))
     
        date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
        date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
     
        If productTemp = "���LOTID����" Then
     
            sqlTemp = "select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW) as ����,substr(X.QBOXNUMBER,2,9) as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn  as OPN ,'' as �����  from ( " & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID ,a.productname, a.mpn ,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' and a.QBOXNUMBER like 'QM%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"
     
        Else

            sqlTemp = " select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW) as ����,substr(X.QBOXNUMBER,2,9) as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn as OPN ,'' as �����  from (" & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID, a.productname, a.mpn,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"
          
        End If
           
        ExporToExcel (sqlTemp)

    End If

End Sub

Private Sub Command2_Click()

'    SqlServer2ExporToExcel ("select AA.�󹤵�,CC.QTY as ������,BB.�����, AA.������ from(select �󹤵�,SUM(����) as ������ from ERPDATA.. TBLSTOCKMOVESUB where �󹤵� is not null group by �󹤵�)AA, " & _
'" (select �󹤵�,SUM(�����) as ����� from dbo.tblPackToHouse  group by �󹤵�) BB, " & _
'" dbo.[tblTSVworkorder] CC Where aa.�󹤵� = BB.�󹤵� and AA.�󹤵� = CC.ORDERNAME and CC.CUSTOMER = '95' and AA.�󹤵� not like '%-16%'")

    Dim beginTime   As String

    Dim endTime     As String

    Dim woTemp      As String

    Dim sqlTemp     As String

    Dim sql1        As String

    Dim billNoTemp  As String

    Dim lotIdTemp   As String

    Dim bigQboxTemp As String

    Dim productTemp As String

    Dim date1Temp   As String

    Dim date2Temp   As String

    Dim sql2        As String

    Dim sql3        As String

    sql1 = ""
    sql2 = ""
    sql3 = ""
    sqlTemp = ""

    If Trim(TxtBillNo.text) <> "" Then   '����

        billNoTemp = UCase(Trim(TxtBillNo.text))

        sqlTemp = " select X.��� , X.qboxnumber, X.������, Sum(X.����) as ����, substring(X.���,2,9) as MPN_SEQ, X.newlotid, X.�Ϻ�, X.MPN as OPN, X.���2 as ����� from ( " & " SELECT distinct b.���,   C.qboxnumber, B.������,B.���̿����,B.����, C.MPN_SEQ,C.newlotid ,B.�Ϻ�,d.mpn,f.��� as ���2 " & " FROM   tblStockMove A ,tblStockMovesub B ,TblQBOXNUMBER_TSVMPN C,dbo.TblQBOXNUMBER_TSV d  ,tblPackTreeInf e ,tblPackTreeInf f " & " Where A.���߱��=1 AND A.�ͻ�����='AA' AND A.��������=1 and A.���ݱ��='" + billNoTemp + "' " & " and b.���ݱ��=a.���ݱ�� and b.������=a.������ and C.qboxnumber=B.���  and d.QBOXNUMBER=C.qboxnumber and e.���=C.qboxnumber and f.���=e.�ϼ����) X " & " group by X.���,  X.qboxnumber, X.������,X.newlotid ,X.�Ϻ�,X.mpn,X.���2 order by X.���2, X.���,X.������ "

        SqlServer2ExporToExcel_Trim (sqlTemp)
        Exit Sub

    End If
    If OptCIS.Value = True Then   ' CIS ��Ϊֱ����FPS���EXCEL
        FpsToExcel
        Exit Sub
    End If
    

    If Trim(TxtInBill.text) <> "" Or OptBUMPING.Value = True Then

        billNoTemp = UCase(Trim(TxtInBill.text))

        If OptAA.Value = True Then ' AA
 
             sqlTemp = "select X.��ⵥ��� , X.���, X.WAFERNUMBER, Sum(X.NDPW),  substring(X.���,2,9) as MPN_SEQ, X.newlotid, X.productName, X.MPN, X.���2  FROM ( " & _
            " SELECT distinct a.��ⵥ���,c.��� ,d.WAFERNUMBER,d.WAFERSCRIBENUMBER,d.NDPW,e.MPN_SEQ,e.newlotid,d.PRODUCTNAME,d.MPN ,b.��� as ���2 " & _
            " FROM dbo.tblPackToHouseSub a ,tblPackTreeInf b ,tblPackTreeInf c ,dbo.TblQBOXNUMBER_TSV d,TblQBOXNUMBER_TSVMPN e " & _
            " WHERE a.��ⵥ���='" & billNoTemp & "' and b.���=a.��� and c.�ϼ����=b.��� and d.QBOXNUMBER=c.��� and e.qboxnumber=c.��� ) X " & _
            " group by X.��ⵥ���,X.��� ,X.WAFERNUMBER,X.newlotid,X.PRODUCTNAME,X.MPN ,X.���2 " & _
            " UNION  " & _
            " SELECT xx.��ⵥ���,xx.��� ,xx.������ ,xx.qty ,xx.AA_Q ,SUBSTRING(yy.Content,CHARINDEX('CUSTOMER_LOT_COMPLEX"",""',yy.Content) + 23,10),xx.PRODUCT,xx.MPN  ,xx.BBOX FROM ( " & _
            " SELECT  x.��ⵥ���,x.���,x.������,x.qty,x.AA_Q,x.KEYID ,X.PRODUCT,  X.MPN,X.BBOX, MAX(y.Createdate) AS createdate FROM ( " & _
            "  SELECT a.��ⵥ���,b.���,b.������,SUM(b.����) AS qty , SUBSTRING(b.���,2,9)  AS AA_Q  ,c.KEYID ,e.PRODUCT,  e.MPN,aa.��� AS BBOX " & _
            "  FROM erpdata..tblPackTreeInf a  ,erpdata..tblPackMainInfSub b ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblTSVworkorder e ,erpdata..tblPackTreeInf aa " & _
            "   WHERE a.��ⵥ��� = '" & billNoTemp & "'  AND b.��� = a.��� AND c.KEY_VALUE = b.���  AND e.ORDERNAME = b.�󹤵�  AND aa.��� = a.�ϼ���� " & _
            "  GROUP BY a.��ⵥ���,b.���,b.������,c.KEYID,e.PRODUCT,  e.MPN,aa.��� ) X  left JOIN   erpdata..tblME_PrintInfo y ON  y.EVENT_ID = x.KEYID  AND y.LABEL_ID = 'AAMPN4' " & _
            "  GROUP BY x.��ⵥ���,x.���,x.������,x.qty,x.AA_Q,x.KEYID,X.PRODUCT,  X.MPN,X.BBOX ) xx " & _
            "  left JOIN erpdata..tblME_PrintInfo yy ON  yy.EVENT_ID = xx.KEYID  AND yy.LABEL_ID = 'AAMPN4' AND yy.Createdate = xx.createdate  ORDER BY X.PRODUCTNAME "
             

 
            SqlServer2ExporToExcel_Trim (sqlTemp)

        ' ElseIf OptCIS.Value = True Then   ' CIS
                   
            ' sqlTemp = " SELECT x.����ʱ��,x.�󹤵�,x.���,x.QTECHPTNO,x.�Ϻ�, " & _
                    ' " case CHARINDEX(',', x.WaferId)  when  0 then Rtrim(x.������)+'.' +x.WaferId  else Rtrim(x.������) +'.' + left(x.WaferId,CHARINDEX(',', x.WaferId)-1) End ,x.������, " & _
                    ' "x.WaferId,SUM(x.gross), SUM(x.good_die), sum(x.price),SUM(x.ng_die),x.����,x.IMAGER_CUSTOMER_REV FROM ( " & _
                    ' " SELECT CONVERT(VARCHAR(20), c.����ʱ��, 23) AS ����ʱ��, b.�󹤵�,  b.���, g.QTECHPTNO, b.�Ϻ�, SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12) AS SFC, b.������, " & _
                    ' "  WaferId = (STUFF((SELECT ',' + SUBSTRING(REPLACE(b1.���̿����, '+', ''), len(REPLACE(b1.���̿����, '+', '')) - 1, 2)  FROM erpdata .. tblPackMainInfSub b1 " & _
                    ' "  WHERE b.��� = b1.��� and  b.�󹤵� = b1.�󹤵� order by b1.���̿���� FOR XML PATH('')), 1,  1, '')),  e.PASSBINCOUNT + e.FAILBINCOUNT AS gross,  SUM(CONVERT(INT, h1.KEY_VALUE)) AS good_die, " & _
                    ' "  COUNT(DISTINCT b.���̿����) AS price,  SUM(CONVERT(INT, h2.KEY_VALUE)) AS ng_die, a1.��� AS ����,  f.IMAGER_CUSTOMER_REV, b.���̿����  FROM erpdata .. tblPackTreeInf a " & _
                    ' " INNER JOIN erpdata .. tblPackMainInfSub b  ON b.��� = a.��� INNER JOIN erpdata .. tblPackMainInf c  ON c.��� = b.��� INNER JOIN ERPBASE .. tblmappingData e ON e.SUBSTRATEID = b.���̿���� " & _
                    ' " INNER JOIN ERPBASE .. tblCustomerOI f  ON  convert(varchar(50),convert(int,f.ID)) = e.FILENAME and f.SOURCE_BATCH_ID = e.LOTID INNER JOIN erptemp .. tbltsvnpiproduct g ON g.QTECHPTNO2 = b.�Ϻ� " & _
                    ' "   AND g.CUSTOMERPTNO1 = f.MPN_DESC  AND g.CUSTOMERSHORTNAME = f.CUSTOMERSHORTNAME INNER JOIN erpdata .. tblErpInStockDetailInfo h  ON h.KEY_VALUE = a.��� INNER JOIN erpdata .. tblErpInStockRelation j " & _
                    ' "  ON j.BOX_ID = h.BOX_ID  AND SUBSTRING(REPLACE(j.WAFER_ID, j.SFC_ID, ''),  2, CHARINDEX('::', REPLACE(j.WAFER_ID, j.SFC_ID, '')) - 2) =  b.���̿����  LEFT join erpdata .. tblErpInStockDetailInfo h1 " & _
                    ' "  ON h1.BOX_ID = h.BOX_ID  AND h1.KEY_NAME = 'GOOD_DIE' AND h1.KEY_TYPE = 'WAFER' AND H1.KEYID = J.WAFER_ID AND h1.KEY_VALUE <> '0'  LEFT join erpdata .. tblPackTreeInf a1  ON a1.��� = a.�ϼ���� " & _
                    ' "  LEFT join erpdata .. tblErpInStockDetailInfo h2   ON h2.BOX_ID = h.BOX_ID AND h2.KEY_NAME IN ('BAD1_DIE', 'BAD2_DIE') AND h2.KEY_TYPE = 'WAFER' AND H2.KEYID = J.WAFER_ID  AND h2.KEY_VALUE <> '0' " & _
                    ' "  WHERE a.��ⵥ��� = '" & billNoTemp & " ' " & _
                    ' " GROUP BY CONVERT(VARCHAR(20), c.����ʱ��, 23), b.�󹤵�, b. ���, g.QTECHPTNO, b.�Ϻ�,b.���̿����,  SUBSTRING(j.SFC_ID, 12, CHARINDEX('.', j.SFC_ID) - 12), b.������, e.PASSBINCOUNT + e.FAILBINCOUNT,  a1.���,  f.IMAGER_CUSTOMER_REV ) x " & _
                    ' "  GROUP BY  x.����ʱ��,x.�󹤵�,x.���,x.QTECHPTNO,x.�Ϻ�,x.SFC,x.������,x.WaferId ,x.����,x.IMAGER_CUSTOMER_REV"
                                  
           ' ' ExporToExcel (sqlTemp)
           ' SqlServer2ExporToExcel_Trim (sqlTemp)

        
        ElseIf OptBUMPING.Value = True Then ' BUMPING
    
            date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
            date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")

            If TxtInBill.text = "" Then
                sqlTemp = "select X.���ʱ�� ,X.��ⵥ���,X.ORDERNAME as ������,X.���,X.SALESORDER as ����,X.PRODUCT as �Ϻ�,X.SFC as ��������,X.������ as �ͻ�LOT,X.���̿���� as WAFER_ID,CONVERT(INT, X.DIEQTY)  - CONVERT(INT, Z.KEY_VALUE) as ������Ʒ,X.KEY_VALUE as �Ƴ���Ʒ,Y.KEY_VALUE as �Ƴ̲���Ʒ,Z.KEY_VALUE as ���ϲ���" & _
                   ",X.MARKINGCODE as �����   from (select CONVERT(varchar(100),aa.���ʱ��,23) as ���ʱ��,a.��ⵥ���,g.ORDERNAME,a.���,g.SALESORDER,g.PRODUCT,replace(d.SFC_ID,'SFCBO:1020,','') as SFC,b.������ " & _
                   ",b.���̿����,e.KEY_VALUE,f.DIEQTY,F.MARKINGCODE,e.KEYID from erpdata..tblPackTreeInf a ,erpdata..tblPackToHouse aa ,erpdata..tblPackMainInfSub b  ,erpdata..tblErpInStockDetailInfo c ,erpdata..tblErpInStockRelation d " & _
                   ",erpdata..tblErpInStockDetailInfo e,erpdata..tblTSVwaferlist f,erpdata..tblTSVworkorder g where aa.���ʱ�� >= '" & date1Temp & "' and aa.���ʱ�� <= '" & date2Temp & "' and a.�ϼ���� <> '0' and aa.��ⵥ��� = a.��ⵥ��� and b.��� = a.��� " & _
                   "and c.KEY_VALUE = b.���  and d.BOX_ID = c.BOX_ID and SUBSTRING(replace(d.WAFER_ID,'SFCBO:1020,',''), CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))+1,CHARINDEX('::',replace(d.WAFER_ID,'SFCBO:1020,',''))-CHARINDEX(',',replace(d.WAFER_ID,'SFCBO:1020,',''))-1) = b.���̿���� " & _
                   "and e.KEYID = d.WAFER_ID and e.KEY_NAME in ('GOOD_DIE' ) and f.WAFERID = b.���̿���� and g.ORDERNAME = f.ORDERNAME and g.ORDERNAME = aa.�󹤵� ) x,erpdata..tblErpInStockDetailInfo y,erpdata..tblErpInStockDetailInfo z where y.KEYID = x.KEYID " & _
                   "and y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' order by x.���ʱ��,x.��ⵥ���, x.���̿����"
            Else
                Dim rkdID As String
                    
                rkdID = Trim$(TxtInBill.text)
            
                sqlTemp = " SELECT X.���ʱ�� ,X.��ⵥ���,X.ORDERNAME AS ������,X.���, X.HT_DEVICE AS Ʒ��,X.PRODUCT AS �Ϻ�,X.SFC AS ��������,X.������ AS LOT_ID, " & _
                        " X.���̿���� AS WAFER_ID,CONVERT(INT, X.DIEQTY) AS �ͻ����GOODDIE,X.KEY_VALUE AS GOODDIE����,1 AS ��Ʒ��վƬ��, Y.KEY_VALUE AS ����NGDIE����, " & _
                        " Z.KEY_VALUE AS �ͻ�NGDIE����,X.MARKINGCODE AS ����������,X.CUSTOMER AS �ͻ����� FROM (SELECT CONVERT(VARCHAR(100), aa.���ʱ��, 23) AS ���ʱ��, " & _
                        " a.��ⵥ���, g.ORDERNAME, a.���,h.HT_DEVICE,g.PRODUCT,REPLACE(d.SFC_ID, 'SFCBO:1020,', '') AS SFC,b.������, b.���̿����,e.KEY_VALUE,f.DIEQTY, " & _
                        " F.MARKINGCODE,e.KEYID,g.CUSTOMER FROM erpdata .. tblPackTreeInf a,erpdata .. tblPackToHouse aa,erpdata .. tblPackMainInfSub b, erpdata .. tblErpInStockDetailInfo c, " & _
                        " erpdata .. tblErpInStockRelation d,erpdata .. tblErpInStockDetailInfo e,erpdata .. tblTSVwaferlist f,erpdata .. tblTSVworkorder g,erpdata .. SHOP_ORDER h " & _
                        " WHERE a.�ϼ���� <> '0'  AND aa.��ⵥ��� = a.��ⵥ��� AND b.��� = a.��� and a.��ⵥ��� = '" & billNoTemp & "' AND c.KEY_VALUE = b.��� AND h.SHOP_ORDER = g.ORDERNAME " & _
                        " AND d.BOX_ID = c.BOX_ID AND SUBSTRING(REPLACE(d.WAFER_ID, 'SFCBO:1020,', ''), CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) + 1, CHARINDEX('::',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - " & _
                        " CHARINDEX(',',REPLACE(d.WAFER_ID, 'SFCBO:1020,', '')) - 1) =b.���̿���� AND e.KEYID = d.WAFER_ID AND e.KEY_NAME IN ('GOOD_DIE') AND f.WAFERID = b.���̿���� AND g.ORDERNAME = f.ORDERNAME) x, " & _
                        " erpdata .. tblErpInStockDetailInfo y, erpdata .. tblErpInStockDetailInfo z WHERE y.KEYID = x.KEYID AND y.KEY_NAME = 'BAD1_DIE' AND Z.KEYID = Y.KEYID AND Z.KEY_NAME = 'BAD2_DIE' "

            End If

            SqlServer2ExporToExcel_Trim (sqlTemp)
        
        Else
        
        End If
  
    End If
    
    If Trim(TxtBillNo.text) = "" And Trim(TxtInBill.text) = "" Then

        productTemp = UCase(Trim(CmbType.text))
     
        date1Temp = Format(DTP1.Value, "YYYY-MM-DD")
        date2Temp = Format(DTP2.Value + 1, "YYYY-MM-DD")
     
        If productTemp = "���LOTID����" Then
     
            sqlTemp = "select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW) as ����,substr(X.QBOXNUMBER,2,9) as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn  as OPN ,'' as �����  from ( " & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID ,a.productname, a.mpn ,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' and a.QBOXNUMBER like 'QM%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"
     
        Else

            sqlTemp = " select  X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,sum(X.NDPW) as ����,substr(X.QBOXNUMBER,2,9) as MPN_SEQ,X.NEWLOTID ,X.productname, X.mpn as OPN ,'' as �����  from (" & " select  distinct a.CONTAINERNAME,a.QBOXNUMBER,a.WAFERNUMBER,a.WAFERSCRIBENUMBER,a.NDPW,b.MPN_SEQ,b.NEWLOTID, a.productname, a.mpn,'' " & " from tsv_qboxnumber_details a ,tsv_qboxnumber_mpn b " & " where a.customername='AA' and a.containername like '%-A%' " & " and a.create_date>=to_date('" + date1Temp + "','YYYY-MM-DD') and a.create_date<to_date('" + date2Temp + "' ,'YYYY-MM-DD')" & " and b.containername=a.containername ) X " & " group by X.CONTAINERNAME,X.QBOXNUMBER,X.WAFERNUMBER,X.NEWLOTID ,X.productname, X.mpn order by X.containername,X.wafernumber"
          
        End If
           
        ExporToExcel (sqlTemp)

    End If

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


woTemp = UCase(Trim(txtWO.text))
productTemp = UCase(Trim(TxtProduct.text))
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


woTemp = UCase(Trim(txtWO.text))
productTemp = UCase(Trim(TxtProduct.text))
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


woTemp = UCase(Trim(txtWO.text))
productTemp = UCase(Trim(TxtProduct.text))
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

TxtBillNo.SetFocus

OptAA.Value = True
DTP1.Value = Now - 1

DTP2.Value = Now

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
        .TypeMaxEditLen = 500
        .ReDraw = False
        .MaxCols = E_FPS0.E_END - 1
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
        

        .SetText E_FPS0.e_ID, 0, "���"
        .SetText E_FPS0.E_ContainerName, 0, "������"
        .SetText E_FPS0.E_QboxName, 0, "С���"
        .SetText E_FPS0.E_LOTID, 0, "LotID"
        .SetText E_FPS0.E_qty, 0, "����"
        .SetText E_FPS0.E_MPNSeq, 0, "SerialNumber"
        .SetText E_FPS0.E_NewLotid, 0, "��LotID"
        .SetText E_FPS0.E_PRODUCT, 0, "�����Ϻ�"
        .SetText E_FPS0.E_CustomerProuduct, 0, "�ͻ��Ϻ�"
        .SetText E_FPS0.E_BigQbox, 0, "�����"


        .ColWidth(E_FPS0.e_ID) = 5
        .ColWidth(E_FPS0.E_ContainerName) = 10
        .ColWidth(E_FPS0.E_QboxName) = 10
        .ColWidth(E_FPS0.E_LOTID) = 10
        .ColWidth(E_FPS0.E_qty) = 5
        .ColWidth(E_FPS0.E_MPNSeq) = 10
        .ColWidth(E_FPS0.E_NewLotid) = 10
        
        .ColWidth(E_FPS0.E_PRODUCT) = 11
        .ColWidth(E_FPS0.E_CustomerProuduct) = 17
        .ColWidth(E_FPS0.E_BigQbox) = 10
        
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub

Private Sub IniFpsHeader2()
    With fps(0)
        .TypeMaxEditLen = 500
        .ReDraw = False
        .MaxCols = E_FPS1.E_END - 1
        .MaxRows = 0
        
        
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .SetText E_FPS1.e_NO, 0, "���"
        .SetText E_FPS1.e_date, 0, "���ʱ��"
        .SetText E_FPS1.e_OrderName, 0, "������"
        .SetText E_FPS1.E_SmallBoxID, 0, "С���"
        .SetText E_FPS1.E_name, 0, "Ʒ��"
        .SetText E_FPS1.E_partno, 0, "�Ϻ�"
       ' .SetText E_FPS1.E_BatchPlant, 0, "��������"
        .SetText E_FPS1.E_BatchPlant, 0, "LOT�ź�׺"
        .SetText E_FPS1.E_LOTID, 0, "LOT��"
        .SetText E_FPS1.E_WAFERID, 0, "WAFERLIST"
        .SetText E_FPS1.E_GrossDie, 0, "�ͻ����GoodDie"
        .SetText E_FPS1.E_GoodDie, 0, "GOODDIE����"
        .SetText E_FPS1.E_Pieces, 0, "����Ƭ��"
        .SetText E_FPS1.E_INsiteNgDie, 0, "����NG"
        .SetText E_FPS1.E_BigBoxID, 0, "�����"
        .SetText E_FPS1.E_SecondCode, 0, "��������"
        .SetText E_FPS1.E_QtyInStock, 0, "�������"
        
        
        .ColWidth(E_FPS1.e_NO) = 10
        .ColWidth(E_FPS1.e_date) = 10
        .ColWidth(E_FPS1.e_OrderName) = 10
        .ColWidth(E_FPS1.E_SmallBoxID) = 10
        .ColWidth(E_FPS1.E_name) = 10
        .ColWidth(E_FPS1.E_partno) = 10
        .ColWidth(E_FPS1.E_BatchPlant) = 10
        .ColWidth(E_FPS1.E_LOTID) = 10
        .ColWidth(E_FPS1.E_WAFERID) = 50
        .ColWidth(E_FPS1.E_GrossDie) = 10
        .ColWidth(E_FPS1.E_GoodDie) = 10
        .ColWidth(E_FPS1.E_Pieces) = 10
        .ColWidth(E_FPS1.E_INsiteNgDie) = 10
        .ColWidth(E_FPS1.E_BigBoxID) = 10
        .ColWidth(E_FPS1.E_SecondCode) = 10
        .ColWidth(E_FPS1.E_QtyInStock) = 10

        
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
             
        .ReDraw = True
    End With
    
    
    

End Sub

Private Sub IniFpsHeader_GC()
    With fps(0)
        .TypeMaxEditLen = 500
        .ReDraw = False
        .MaxCols = E_FPS_GC.E_END - 1
        .MaxRows = 0
        
        
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .SetText E_FPS_GC.e_NO, 0, "���"
        .SetText E_FPS_GC.e_date, 0, "���ʱ��"
        .SetText E_FPS_GC.e_OrderName, 0, "������"
        .SetText E_FPS_GC.E_SmallBoxID, 0, "С���"
        .SetText E_FPS_GC.E_name, 0, "Ʒ��"
        .SetText E_FPS_GC.E_partno, 0, "�Ϻ�"
       ' .SetText E_FPS1.E_BatchPlant, 0, "��������"
        .SetText E_FPS_GC.E_BatchPlant, 0, "LOT��+��׺"
        .SetText E_FPS_GC.E_LOTID, 0, "LOT��"
        .SetText E_FPS_GC.E_WAFERID, 0, "WAFER��"
        .SetText E_FPS_GC.E_GrossDie, 0, "�ͻ����GoodDie"
        .SetText E_FPS_GC.E_GoodDie, 0, "GOODDIE����"
        .SetText E_FPS_GC.E_Pieces, 0, "����Ƭ��"
        .SetText E_FPS_GC.E_INsiteNgDie, 0, "����NG"
        .SetText E_FPS_GC.E_BigBoxID, 0, "�����"
        .SetText E_FPS_GC.E_SecondCode, 0, "��������"
        .SetText E_FPS_GC.E_Mpndesc, 0, "�ͻ�����"
        .SetText E_FPS_GC.E_Bond, 0, "��˰��Ǳ�"
        .SetText E_FPS_GC.E_ProductType, 0, "��ʽ"
        .SetText E_FPS_GC.E_QtyInStock, 0, "�������"
        .SetText E_FPS_GC.E_DieByPcs, 0, "��Ƭ����"
        .SetText E_FPS_GC.e_ShipNo, 0, "��������"
        .SetText E_FPS_GC.E_ShipTo, 0, "������"
        
        
        
        .ColWidth(E_FPS_GC.e_NO) = 10
        .ColWidth(E_FPS_GC.e_date) = 10
        .ColWidth(E_FPS_GC.e_OrderName) = 10
        .ColWidth(E_FPS_GC.E_SmallBoxID) = 10
        .ColWidth(E_FPS_GC.E_name) = 10
        .ColWidth(E_FPS_GC.E_partno) = 15
        .ColWidth(E_FPS_GC.E_BatchPlant) = 10
        .ColWidth(E_FPS_GC.E_LOTID) = 10
        .ColWidth(E_FPS_GC.E_WAFERID) = 50
        .ColWidth(E_FPS_GC.E_GrossDie) = 10
        .ColWidth(E_FPS_GC.E_GoodDie) = 10
        .ColWidth(E_FPS_GC.E_Pieces) = 10
        .ColWidth(E_FPS_GC.E_INsiteNgDie) = 10
        .ColWidth(E_FPS_GC.E_BigBoxID) = 10
        .ColWidth(E_FPS_GC.E_SecondCode) = 10
        .ColWidth(E_FPS_GC.E_Mpndesc) = 10
        .ColWidth(E_FPS_GC.E_Bond) = 10
        .ColWidth(E_FPS_GC.E_ProductType) = 10
        .ColWidth(E_FPS_GC.E_QtyInStock) = 10
        
        .ColWidth(E_FPS_GC.E_DieByPcs) = 10
        .ColWidth(E_FPS_GC.e_ShipNo) = 10
        .ColWidth(E_FPS_GC.E_ShipTo) = 10
        
        
        
        
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
             
        .ReDraw = True
    End With
    
    
    

End Sub
Private Sub IniFpsHeader3()
    With fps(0)
        .TypeMaxEditLen = 500
        .ReDraw = False
        .MaxCols = E_FPS1.E_END - 1
        .MaxRows = 0
        
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = False
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        
        .RowHeight(0) = 20
        .RowHeight(-1) = 15
             
        .ReDraw = True
    End With
    
End Sub



Private Sub Form_Load()
'IniProduct
OptCIS.Value = True


End Sub


Private Sub OptAA_Click()
lblLabel4.Visible = False
cbCusCode.Visible = False
chkStockOnly.Visible = False

IniFpsHeader1
End Sub

Private Sub OptCIS_Click()
lblLabel4.Visible = True
cbCusCode.Visible = True
chkStockOnly.Visible = True
InitCuscode
'IniFpsHeader2

End Sub


Private Sub OptBUMPING_Click()
lblLabel4.Visible = False
cbCusCode.Visible = False
chkStockOnly.Visible = False
IniFpsHeader3

End Sub

Private Function SqlServer2ExporToExcel_Trim(strOpen As String)
'����ȥǰ��ո�Ĺ���
'*********************************************************
'* �W�١GExporToExcel
'* �\��G�ɥX���u��EXCEL
'* �Ϊk�GExporToExcel(sql�d�ߦr�Ŧ�)
'*********************************************************
' ����SqlServer Excel

    Dim Rs_Data As New ADODB.Recordset
    Dim Irowcount As Long
    Dim Icolcount As Integer
    Dim i As Integer
    Dim j As Integer
    
    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlQuery As Excel.QueryTable
    
'      If Cnn.State = 0 Then
'        ConOracle
'      End If
      
    If INIadoCon2.State = 0 Then
        INIConnectSTART2
    End If


    
    
   ' CONN_TO_ORACLE_DATABASE2
    
    With Rs_Data
        If .State = adStateOpen Then
            .Close
        End If
        .ActiveConnection = INIadoCon2
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = strOpen
        .Open
    End With
    With Rs_Data
        If .RecordCount < 1 Then
            MsgBox ("��ѯ��������!")
            Exit Function
        End If
        '�O���`��
        Irowcount = .RecordCount
        '�r�q�`��
        Icolcount = .Fields.count
    End With
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlApp.Visible = False
    
    '�K�[�d�߻y�y�A�ɤJEXCEL�ƾ��u
    Set xlQuery = xlSheet.QueryTables.Add(Rs_Data, xlSheet.Range("a1"))
    
    With xlQuery
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = True
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
    End With
    
    xlQuery.FieldNames = True '��ܦr�q�W
    xlQuery.Refresh
    
    With xlSheet
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.name = "���^"
        '�]�m���^�r
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = True
        '�]�m�r�^�[��
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
        '�]�m�����ؼ˦�

        For i = 2 To Irowcount + 1
            For j = 1 To Icolcount
                .Cells(i, j).Value = Replace(Trim(.Cells(i, j).Value), Chr(10), "")
            Next
        Next
    End With
    With xlSheet.Range("2:" & Irowcount + 1)
        .horizontalAlignment = xlLeft
    End With
    xlSheet.Range("A1").Select
    xlApp.Columns.AutoFit
    
    
    
    
'    With xlSheet.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""���^_GB2312,�`?""&10���q�W�١G"   ' & Gsmc
'        .CenterHeader = "&""���^_GB2312,�`�W""���ʥN��궵��&""���^,�`�W""" & Chr(10) & "&""���^_GB2312,�`?""&10�� ���G"
'        .RightHeader = "" & Chr(10) & "&""���^_GB2312,�`�W""&10���G���h"
'        .LeftFooter = "&""���^_GB2312,�`�W""&10���H�G"
'        .CenterFooter = "&""���^_GB2312,�`�W""&10������G"
'        .RightFooter = "&""���^_GB2312,�`?""&10��&P���@&N��"
'    End With
    
    xlApp.Application.Visible = True
    Set xlApp = Nothing  '"���ٱ��Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Function


Private Function SqlServer2ExporToExcel_Trim_AA(strOpen As String)
'����ȥǰ��ո�Ĺ���
'*********************************************************
'* �W�١GExporToExcel
'* �\��G�ɥX���u��EXCEL
'* �Ϊk�GExporToExcel(sql�d�ߦr�Ŧ�)
'*********************************************************
' ����SqlServer Excel

    Dim Rs_Data As New ADODB.Recordset
    Dim Irowcount As Long
    Dim Icolcount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Bigboxid As Integer
    Dim QtyinBigbox As Long
    Dim WeightinBigbox As Double
    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlQuery As Excel.QueryTable
    
'      If Cnn.State = 0 Then
'        ConOracle
'      End If
      
    If INIadoCon2.State = 0 Then
        INIConnectSTART2
    End If


    
    
   ' CONN_TO_ORACLE_DATABASE2
    
    With Rs_Data
        If .State = adStateOpen Then
            .Close
        End If
        .ActiveConnection = INIadoCon2
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = strOpen
        .Open
    End With
    With Rs_Data
        If .RecordCount < 1 Then
            MsgBox ("��ѯ��������!")
            Exit Function
        End If
        '�O���`��
        Irowcount = .RecordCount
        '�r�q�`��
        Icolcount = .Fields.count
    End With
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlApp.Visible = False
    
    '�K�[�d�߻y�y�A�ɤJEXCEL�ƾ��u
    Set xlQuery = xlSheet.QueryTables.Add(Rs_Data, xlSheet.Range("a1"))
    
    With xlQuery
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = True
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
    End With
    
    xlQuery.FieldNames = True '��ܦr�q�W
    xlQuery.Refresh
    
    With xlSheet
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.name = "���^"
        '�]�m���^�r
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = True
        '�]�m�r�^�[��
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
        '�]�m�����ؼ˦�

        For i = 2 To Irowcount + 1
            For j = 1 To Icolcount
                .Cells(i, j).Value = Replace(Trim(.Cells(i, j).Value), Chr(10), "")
            Next
        Next
        '�ϲ���Ԫ��
        Bigboxid = 0
        QtyinBigbox = 0
        For i = 2 To Irowcount + 1
           Bigboxid = Bigboxid + 1
           QtyinBigbox = Val(.Cells(i, 4).Value)
           WeightinBigbox = Val(.Cells(i, 4).Value) * Val(.Cells(i, 9).Value)
            For j = i + 1 To Irowcount + 1
               If .Cells(j, 10).Value <> .Cells(i, 10).Value Then
                   Exit For
               End If
               QtyinBigbox = QtyinBigbox + Val(.Cells(j, 4).Value)
               WeightinBigbox = WeightinBigbox + Val(.Cells(j, 4).Value) * Val(.Cells(j, 9).Value)
            Next
            .Cells(i, 11).Value = Bigboxid '�����
            .Cells(i, 12).Value = QtyinBigbox '����
            .Cells(i, 13).Value = WeightinBigbox '����
            .Cells(i, 16).Value = Format(Now(), "yyyy/mm/dd") '����
            xlApp.Application.DisplayAlerts = False '��������ʾ��
            xlSheet.Range(Chr(10 + 64) & i & ":" & Chr(10 + 64) & j - 1).Merge
            xlSheet.Range(Chr(11 + 64) & i & ":" & Chr(11 + 64) & j - 1).Merge
            xlSheet.Range(Chr(12 + 64) & i & ":" & Chr(12 + 64) & j - 1).Merge
            xlSheet.Range(Chr(13 + 64) & i & ":" & Chr(13 + 64) & j - 1).Merge
            xlSheet.Range(Chr(14 + 64) & i & ":" & Chr(14 + 64) & j - 1).Merge
            xlSheet.Range(Chr(15 + 64) & i & ":" & Chr(15 + 64) & j - 1).Merge
            xlSheet.Range(Chr(16 + 64) & i & ":" & Chr(16 + 64) & j - 1).Merge
            i = j - 1
        Next
        .Cells(Irowcount + 2, 4).Value = WorksheetFunction.Sum(Range(Chr(4 + 64) & "2:" & Chr(4 + 64) & Irowcount + 1))
        .Cells(Irowcount + 2, 11).Value = Bigboxid
        .Cells(Irowcount + 2, 12).Value = WorksheetFunction.Sum(Range(Chr(12 + 64) & "2:" & Chr(12 + 64) & Irowcount + 1))
        .Cells(Irowcount + 2, 13).Value = WorksheetFunction.Sum(Range(Chr(13 + 64) & "2:" & Chr(13 + 64) & Irowcount + 1))
        .Cells(Irowcount + 2, 14).Value = WorksheetFunction.Sum(Range(Chr(14 + 64) & "2:" & Chr(14 + 64) & Irowcount + 1))
      
    End With
    With xlSheet.Range("2:" & Irowcount + 2)
        .horizontalAlignment = xlLeft
    End With
    xlSheet.Range("A1").Select
    xlApp.Columns.AutoFit


    
'    With xlSheet.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""���^_GB2312,�`?""&10���q�W�١G"   ' & Gsmc
'        .CenterHeader = "&""���^_GB2312,�`�W""���ʥN��궵��&""���^,�`�W""" & Chr(10) & "&""���^_GB2312,�`?""&10�� ���G"
'        .RightHeader = "" & Chr(10) & "&""���^_GB2312,�`�W""&10���G���h"
'        .LeftFooter = "&""���^_GB2312,�`�W""&10���H�G"
'        .CenterFooter = "&""���^_GB2312,�`�W""&10������G"
'        .RightFooter = "&""���^_GB2312,�`?""&10��&P���@&N��"
'    End With
    
    xlApp.Application.Visible = True
    Set xlApp = Nothing  '"���ٱ��Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing

End Function


Private Sub FpsToExcel()
    If fps(0).MaxRows = 0 Then
        MsgBox "û�����ݿ��Ե���", vbInformation, "��ʾ"
        Exit Sub
    End If

    Dim i As Integer
    Dim j As Integer
    
    Dim xlApp      As Excel.Application
    Dim xlBook     As Excel.Workbook
    Dim xlSheet    As Excel.Worksheet
    

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

    With xlApp
        .Rows(1).Font.Bold = True
    End With
    
 On Error GoTo Ert
    With fps(0)

        For i = 0 To .MaxRows
            For j = 1 To .MaxCols
                .Col = j
                .Row = i
                xlSheet.Cells(i + 1, j) = Trim$(("'" & .text))
            Next j
       
        Next i
        
    End With

    '�����и�ʽ����
    For j = 1 To fps(0).MaxCols
        If Trim(xlSheet.Cells(1, j)) = "�ͻ����GoodDie" Or Trim(xlSheet.Cells(1, j)) = "��Ƭ����" Or Trim(xlSheet.Cells(1, j)) = "GOODDIE����" Or Trim(xlSheet.Cells(1, j)) = "����Ƭ��" Or Trim(xlSheet.Cells(1, j)) = "����NG" Or Trim(xlSheet.Cells(1, j)) = "�������" Then
            For i = 2 To fps(0).MaxRows + 1
                xlSheet.Cells(i, j) = Replace(xlSheet.Cells(i, j), "'", "")
            Next
        End If
    Next
    With xlSheet.Range("2:" & fps(0).MaxRows + 1)
        .horizontalAlignment = xlLeft
    End With
    xlSheet.Range("A1").Select
    xlApp.Columns.AutoFit
    
    xlApp.Application.Visible = True
    
    
    Set xlApp = Nothing  '"���ٱ��Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing
Ert:
    If Not (xlApp Is Nothing) Then
        
    Set xlApp = Nothing  '"���ٱ��Excel
    Set xlBook = Nothing
    Set xlSheet = Nothing
    End If
    
    
End Sub






























