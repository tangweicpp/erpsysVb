VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmSemtech_LablePrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Semtech��ǩ��ӡ"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
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
   ScaleHeight     =   7200
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra 
      Height          =   855
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12375
      Begin VB.TextBox txtText1 
         Height          =   405
         Left            =   8040
         TabIndex        =   34
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         Caption         =   "��  ѯ"
         Height          =   360
         Index           =   0
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmd 
         Caption         =   "������ǰ����"
         Height          =   360
         Index           =   1
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd 
         Caption         =   "�� ��"
         Height          =   360
         Index           =   3
         Left            =   5760
         TabIndex        =   11
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmd 
         Caption         =   "��  ӡ"
         Height          =   360
         Index           =   2
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblLOT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ɢ��LOT"
         Height          =   195
         Left            =   7200
         TabIndex        =   33
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "��ѯ����"
      ForeColor       =   &H00FF0000&
      Height          =   7335
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   3495
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   450
         Index           =   5
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1560
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   765
         Index           =   4
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   4680
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   765
         Index           =   3
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   3840
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   765
         Index           =   2
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   3000
         Width           =   2355
      End
      Begin VB.Frame Fra 
         Height          =   1575
         Index           =   3
         Left            =   0
         TabIndex        =   15
         Top             =   5760
         Width           =   3495
         Begin MSComCtl2.DTPicker DTP 
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   16
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   221511681
            CurrentDate     =   41387
         End
         Begin MSComCtl2.DTPicker DTP 
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   17
            Top             =   840
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   221511681
            CurrentDate     =   41387
         End
         Begin VB.Label lblJobNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   720
         End
         Begin VB.Label lblJobNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʼ����"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   720
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "��ǩ����"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   5520
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   765
         Index           =   1
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2160
         Width           =   2355
      End
      Begin VB.ComboBox cmbDN 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "FrmSemtech_LablePrint.frx":0000
         Left            =   960
         List            =   "FrmSemtech_LablePrint.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   810
         Index           =   0
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label lblJobNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��        ��"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label lblJobNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����·��"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   4920
         Width           =   720
      End
      Begin VB.Label lblJobNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ں�·��"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   4080
         Width           =   720
      End
      Begin VB.Label lblJobNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����·��"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   3240
         Width           =   720
      End
      Begin VB.Label lblJobNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job      No"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lblJobNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "   DN"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   525
      End
      Begin VB.Label lblJobNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǩ����"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   720
      End
   End
   Begin VB.Frame Fra 
      ForeColor       =   &H000000FF&
      Height          =   7335
      Index           =   1
      Left            =   3480
      TabIndex        =   0
      Top             =   840
      Width           =   9615
      Begin VB.OptionButton Opt 
         Caption         =   "ɢ��1"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   6120
         TabIndex        =   35
         Top             =   0
         Width           =   975
      End
      Begin VB.OptionButton Opt 
         Caption         =   "ɢ��2"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   32
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkChoose 
         Caption         =   "ȫѡ"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton Opt 
         Caption         =   "����"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   28
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton Opt 
         Caption         =   "����"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   27
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton Opt 
         Caption         =   "����"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   26
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   3255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6015
         _Version        =   524288
         _ExtentX        =   10610
         _ExtentY        =   5741
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
         SpreadDesigner  =   "FrmSemtech_LablePrint.frx":0016
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
End
Attribute VB_Name = "FrmSemtech_LablePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strdjbh             As String
Dim strShipToCust       As String
Const C_Left = 60
Const C_Top = 120

Private Enum FpsDetail
    e_Choose = 1
End Enum

'Check��ı仯
Private Sub chk_Click()
    If chk.Value = 1 Then
        Fra(3).Visible = True
    Else
        Fra(3).Visible = False
    End If
End Sub
'ȫѡ��ȡ��
Private Sub chkChoose_Click()
Dim i           As Long

    If Fps(0).MaxRows <= 0 Then
        chkChoose.Value = 0
    Else
        With Fps(0)
            For i = 1 To .MaxRows
                If chkChoose.Value = 1 Then 'ȫѡ
                    .SetText FpsDetail.e_Choose, i, "1"
                Else    'ȡ��
                    .SetText FpsDetail.e_Choose, i, "0"
                End If
            Next
        End With
    End If
    
End Sub

'DN�ı仯
Private Sub cmbDN_Click()
Dim i                   As Long
Dim strSql              As String
Dim Rs                  As New adodb.Recordset
    
    '�ȳ�ʼ���ؼ�
    For i = 0 To txt.UBound
        txt(i).Text = ""
    Next
    chk.Value = 0
    Fps(0).MaxRows = 0
    chkChoose.Value = 0
    '��ѯ���ݸ�ֵ���ؼ�
    strSql = "SELECT a.BatchNumber,a.LabelRequirement,b.PARA,b.PARA1,b.PARA2,a.Quantity,a.ShipToCustomer " & _
             " FROM erpbase..tblCustomerShippingUp a " & _
             " LEFT JOIN erpdata..tblSysIncrement b ON a.ShipToCustomer=b.Kind " & _
             " WHERE a.flag='Y' AND a.customershortname='37' " & _
             " AND a.Delivery='" & Trim$(cmbDN.Text) & "'"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If Not Rs.EOF Then
        Do While Not Rs.EOF
            txt(0).Text = txt(0).Text + Trim$("" & Rs!BatchNumber) + ";"
            txt(1).Text = Trim$("" & Rs!LabelRequirement)
            txt(2).Text = Trim$("" & Rs!PARA)
            txt(3).Text = Trim$("" & Rs!Para1)
            txt(4).Text = Trim$("" & Rs!Para2)
            txt(5).Text = Val(txt(5).Text) + Val(Trim$("" & Rs!QUANTITY))
            strShipToCust = Trim$("" & Rs!ShipToCustomer)   '������
            Rs.MoveNext
        Loop
    End If
    Rs.Close

End Sub

Private Sub cmd_Click(Index As Integer) '��ѯ����
Dim i                   As Long
Dim strSql              As String
Dim Rs                  As New adodb.Recordset
Dim strExportName       As String
Dim lotTemp As String
lotTemp = Trim(txtText1.Text)
    '---------------------------------------------
    If Index = 0 Then          '��ѯ
        If cmbDN.Text = "" And Opt(3).Value = False And Opt(4).Value = False Then
            MsgBox "����ѡ��DN�ţ�"
            Exit Sub
        End If
        If Opt(0).Value = True Then '����
            If strShipToCust = "2000561" Then   '��LG�ı�ǩ
                strSql = "SELECT 0 'ѡ��',a.TRAYQBOXNUMBER QBOXNUMBER,b.CustomerPartNumber,b.BatchNumber,a.HTLOTID,a.PoDateCode,a.QTY " & _
                         " FROM erpdata..TblTSV_Tray_details a " & _
                         " INNER JOIN (SELECT DISTINCT Delivery,CustomerPartNumber,BatchNumber,VendorLotNumber " & _
                         " FROM erpbase..tblCustomerShippingUp WHERE flag='Y' AND customershortname='37') b ON a.CUSTOMERLOTID=b.BatchNumber " & _
                         " Inner Join erpdata..tblstocknumsub c on a.TRAYQBOXNUMBER=rtrim(c.���) " & _
                         " WHERE b.Delivery='" & Trim$(cmbDN.Text) & "'"
            Else
                strSql = "SELECT 0 'ѡ��',a.TRAYQBOXNUMBER QBOXNUMBER,b.CustomerPartNumber,'TVS DIODES' Specification" & _
                         ",CASE WHEN CHARINDEX('E2',b.LabelRequirement)>0 THEN 'PO TYPE,:E2' ELSE ',' END Potype" & _
                         ",a.CUSTOMERLOTID,a.QTY,b.MarketingPN,'DPTK' VendorCode " & _
                         " FROM erpdata..TblTSV_Tray_details a " & _
                         " INNER JOIN (SELECT DISTINCT Delivery,CustomerPartNumber,LabelRequirement,MarketingPN,BatchNumber " & _
                         " FROM erpbase..tblCustomerShippingUp WHERE flag='Y' AND customershortname='37') b ON a.CUSTOMERLOTID=b.BatchNumber " & _
                         " Inner Join erpdata..tblstocknumsub c on a.TRAYQBOXNUMBER=rtrim(c.���) " & _
                         " WHERE b.Delivery='" & Trim$(cmbDN.Text) & "'"
            End If
        End If
        If Opt(1).Value = True Then '�ں�
             If strShipToCust = "2000561" Then   '��LG�ı�ǩ
                strSql = "SELECT 0 'ѡ��',a.CONTAINERNAME QBOXNUMBER,b.CustomerPartNumber,b.BatchNumber,a.HTLOTID,a.PoDateCode,SUM(a.QTY) Qty " & _
                         " FROM erpdata..TblTSV_INBOX_DETAILS a " & _
                         " INNER JOIN (SELECT DISTINCT Delivery,CustomerPartNumber,BatchNumber,VendorLotNumber " & _
                         " FROM erpbase..tblCustomerShippingUp WHERE flag='Y' AND customershortname='37') b ON a.CUSTOMERLOTID=b.BatchNumber " & _
                         " Inner Join erpdata..tblstocknumsub c on a.SUBCONTAINERNAME=rtrim(c.���) " & _
                         " WHERE b.Delivery='" & Trim$(cmbDN.Text) & "'"
             Else
                strSql = "SELECT 0 'ѡ��',a.NHBox QBOXNUMBER,b.CustomerPartNumber,'TVS DIODES' Specification" & _
                         ",CASE WHEN CHARINDEX('E2',b.LabelRequirement)>0 THEN 'PO TYPE,:E2' ELSE ',' END Potype" & _
                         ",SUM(a.QTY) Qty,b.MarketingPN,'DPTK' VendorCode " & _
                         " FROM erpdata..TblTSV_INBOX_DETAILS a " & _
                         " INNER JOIN (SELECT DISTINCT Delivery,CustomerPartNumber,LabelRequirement,MarketingPN,BatchNumber " & _
                         " FROM erpbase..tblCustomerShippingUp WHERE flag='Y' AND customershortname='37') b ON a.CUSTOMERLOTID=b.BatchNumber " & _
                         " Inner Join erpdata..tblstocknumsub c on a.SUBCONTAINERNAME=rtrim(c.���) " & _
                         " WHERE b.Delivery='" & Trim$(cmbDN.Text) & "'"
            End If
        End If
        If Opt(2).Value = True Then '����
             strSql = "SELECT 0 'ѡ��',a.CONTAINERNAME QboxNumber,a.Invoice,'I'+a.Invoice Invoice1,left(a.PONO,10),'K'+left(a.PONO,10) PONO1,CASE WHEN CHARINDEX('E2',b.LabelRequirement)>0 THEN 'E2' ELSE '' END Potype" & _
                      ",left(a.customerPT,11),'P'+left(a.customerPT,11) customerPT1,a.MFGPT,'Z'+a.MFGPT MFGPT1,SUM(a.Qty) ����,'Q'+Rtrim(SUM(a.Qty)) ����1,a.Forwarder,a.coo " & _
                      ",left(a.shiptoname,33),a.shiptostreet1,a.shiptostreet2,a.shiptostreet3,a.shiptostreet4,a.countrykey " & _
                      ",'Attn:'+a.contactname+';Tel:'+a.phone ��ϵ��, 'N/A','P' +'N/A','N/A','9D' + 'N/A'" & _
                      " FROM erpdata..TblTSV_OutBOX_DETAILS a " & _
                      " INNER JOIN (SELECT DISTINCT Delivery,LabelRequirement FROM erpbase..tblCustomerShippingUp WHERE flag='Y' AND customershortname='37') b ON a.Invoice=b.Delivery " & _
                      " WHERE a.Invoice='" & Trim$(cmbDN.Text) & "'"
        End If
        '���ͨ��
        If chk.Value = 0 Then
            strSql = strSql & " And a.PrintFlag=0 "
        Else    '�����ǩ
            strSql = strSql & " And a.PrintFlag=1 And a.PrintTime>='" & DTP(0).Value & "' And a.PrintTime<'" & DTP(1).Value + 1 & "'"
        End If
        
        If Opt(1).Value = True Then '�ں�
            If strShipToCust = "2000561" Then   '��LG�ı�ǩ
                strSql = strSql & " GROUP BY a.CONTAINERNAME,b.CustomerPartNumber,b.BatchNumber,a.HTLOTID,a.PoDateCode "
            Else
                strSql = strSql & " GROUP BY  a.NHBox,b.CustomerPartNumber " & _
                        ",CASE WHEN CHARINDEX('E2',b.LabelRequirement)>0 THEN 'PO TYPE,:E2' ELSE ',' END,b.MarketingPN"
            End If
        End If
        If Opt(2).Value = True Then '����
            strSql = strSql & " GROUP BY a.CONTAINERNAME,a.Invoice,'I'+a.Invoice,left(a.PONO,10),'K'+left(a.PONO,10),CASE WHEN CHARINDEX('E2',b.LabelRequirement)>0 THEN 'E2' ELSE '' END" & _
                     ",left(a.customerPT,11),'P'+left(a.customerPT,11),a.MFGPT,'Z'+a.MFGPT,a.forwarder,a.coo,left(a.shiptoname,33),a.shiptostreet1,a.shiptostreet2 " & _
                     ",a.shiptostreet3,a.shiptostreet4,a.countrykey,'Attn:'+a.contactname+';Tel:'+a.phone"
        End If
        
        If Opt(3).Value = True Then
        If Len(lotTemp) < 2 Then
           MsgBox "������LOT�ţ�"
            Exit Sub
        End If
         strSql = "select distinct '', ct.fab_conv_id,ct.mpn_desc, A.Waferscribenumber, to_char(WW.CREATE_DATE + 6, 'YYWW') as DateCode, B.QTY,WW.WEIGHT,ct.test_mtrl_desc, get_37bagid(b.containername) as code1," & _
            " to_char(sysdate, 'mm/dd/yyyy') as Pdate, to_char(sysdate, 'hh24:mi:ss') as Pdate1,ct.fab_conv_id || ';' || ct.mpn_desc || ';' || A.Waferscribenumber || ';' || to_char(WW.CREATE_DATE + 6, 'YYWW') || ';' || B.QTY || ';' || " & _
            " WW.WEIGHT || ';' || ct.test_mtrl_desc || ';' || get_37bagid(b.containername) as code2  from a_lotwafers a, CONTAINER B, a_lotattributes c, PRODUCT  P, customeroitbl_test ct, mappingdatatest  mt, " & _
            "  weight37 ww,ib_wohistory  ibo, mfgorder f  where a.containerid = b.containerid  AND P.PRODUCTID = B.PRODUCTID   and a.waferscribenumber = mt.substrateid   AND A.WAFERSCRIBENUMBER = MT.SUBSTRATEID " & _
            " AND MT.FILENAME = to_char(CT.ID)  AND WW.WAFERID = REPLACE(MT.SUBSTRATEID, '+', '') and b.containerid = c.containerid and f.mfgordername = a.workordername and ibo.ordername = f.mfgordername " & _
            " and mt.filename = ct.id  and ct.source_batch_id = '" & lotTemp & "'  AND MT.CUSTOMERSHORTNAME = '37' AND C.WAFERBIN = 'A' and mt.substrateid not like '%+' "
  
        End If
        
            If Opt(4).Value = True Then
        If Len(lotTemp) < 2 Then
           MsgBox "������LOT�ţ�"
            Exit Sub
        End If
         strSql = "select distinct '', ct.fab_conv_id,ct.mpn_desc,'Production','D' || to_char(WW.CREATE_DATE + 6, 'YYWW') || 'B' ||cc.bline || 'C' || cc.code as DateCode,get_37bagid(b.containername) as code1," & _
                " to_char(sysdate, 'mm/dd/yyyy') as Pdate,to_char(sysdate, 'hh24:mi:ss') as Pdate1,trglabelseq.QTSeq_37(b.containername), DC.NOTES,cc.code  from a_lotwafers  a,  CONTAINER  B, " & _
                "  a_lotattributes c, PRODUCT P,customeroitbl_test ct, mappingdatatest mt, ib_wohistory  ibo, mfgorder f,datecode37 dc, CODE37 cc,WEIGHT37  WW  where a.containerid = b.containerid " & _
                " AND P.PRODUCTID = B.PRODUCTID and f.mfgordername = a.workordername   AND A.WAFERSCRIBENUMBER = MT.SUBSTRATEID  AND MT.FILENAME = CT.ID AND WW.WAFERID = REPLACE(MT.SUBSTRATEID, '+', '') " & _
                " and b.containerid = c.containerid and ibo.ordername = f.mfgordername and cc.device = ct.mpn_desc and dc.datecode = to_char(WW.CREATE_DATE + 6, 'YYWW') and a.waferscribenumber = mt.substrateid " & _
                " and mt.filename = ct.id and ct.source_batch_id =  '" & lotTemp & "'  AND MT.CUSTOMERSHORTNAME = '37'  AND C.WAFERBIN = 'A' "
        End If
      If Opt(3).Value = True Or Opt(4).Value = True Then
      
            If Rs.State = adStateOpen Then Rs.Close
        Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
        Fps(0).MaxRows = 0
        If Not Rs.EOF Then
            With Fps(0)
                .MaxRows = 0
                Set .DataSource = Rs
                .MaxRows = Rs.RecordCount
            End With
        End If
        Rs.Close
      Else
        
        '��ѯ���ݵ�FPS
        If Rs.State = adStateOpen Then Rs.Close
        Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
        Fps(0).MaxRows = 0
        If Not Rs.EOF Then
            With Fps(0)
                .MaxRows = 0
                Set .DataSource = Rs
                .MaxRows = Rs.RecordCount
            End With
        End If
        Rs.Close
     End If
    ElseIf Index = 1 Then   '����
        
        If Opt(0).Value = True Then
            strExportName = Opt(0).Caption + "��ǩ��Ϣ"
        ElseIf Opt(1).Value = True Then
            strExportName = Opt(1).Caption + "��ǩ��Ϣ"
        ElseIf Opt(2).Value = True Then
            strExportName = Opt(2).Caption + "��ǩ��Ϣ"
        End If
        If Not ExportFpspreadToExcel(Fps(0), strExportName, strExportName) Then Exit Sub
    
    ElseIf Index = 2 Then   '��ӡ
      
        'У������
        If Not CheckData Then Exit Sub
        '��ʼ��Fps���ݣ������ǩͷ����ӡ��ǩ
        Call IniLable
        
    ElseIf Index = 3 Then   '�˳�
        Unload Me
    End If
'
End Sub
'��ʼ����ǩ���
Private Sub IniLable()
Dim i               As Long
Dim j               As Integer
Dim strTmp(9)       As String
Dim strLable        As String
Dim strFileName     As String
    
    With Fps(0)

        Set .DataSource = Nothing
        For i = .MaxRows To 1 Step -1
            .Row = i
            .Col = FpsDetail.e_Choose
            If .Value = 1 Then  'ѡ���˴�ӡ����
                strLable = ""
                strFileName = ""
                If Opt(0).Value = True Then     '����
                    If strShipToCust = "2000561" Then   '��LG�ı�ǩ
                        .Col = 2    '���
                        strFileName = Trim$(.Text)
                        .Col = 3    'DVC
                        strTmp(0) = Trim$(.Text)
                        .Col = 4    'Wafer Lot
                        strTmp(1) = Trim$(.Text)
                        .Col = 5    'Assy Lot
                        strTmp(2) = Trim$(.Text)
                        .Col = 6    'Date Code
                        strTmp(3) = Trim$(.Text)
                        .Col = 7    'Qty
                        strTmp(4) = Trim$(.Text)
                        '�����ǩ���
                        strLable = strTmp(0) + "," + strTmp(1) + "," + strTmp(2) + "," + strTmp(3) + "," + strTmp(4)
                    Else
                        .Col = 2    '���
                        strFileName = Trim$(.Text)
                        .Col = 3    'Part No
                        strTmp(0) = Trim$(.Text)
                        .Col = 4    'SPECIFICATION
                        strTmp(1) = Trim$(.Text)
                        .Col = 5    'PO Type
                        strTmp(8) = Trim$(.Text)
                        If InStr(strTmp(8), "E2") > 0 Then
                            strTmp(2) = "E2"
                        Else
                            strTmp(2) = ""
                        End If
                        .Col = 6    'LOT NO
                        strTmp(3) = Trim$(.Text)
                        strTmp(4) = GetLableXH(strTmp(3)) '���
                        .Col = 7    'Qty
                        strTmp(5) = Trim$(.Text)
                        .Col = 8    'VenDor P/N
                        strTmp(6) = Trim$(.Text)
                        .Col = 9    'VenDor Code
                        strTmp(7) = Trim$(.Text)
                        '�����ǩ���
                        strLable = strTmp(0) + strTmp(7) + strTmp(2) + Left$(strTmp(3) + "00000000", 8) + strTmp(4) + Right$("000000" + strTmp(5), 6) + ","
                        strLable = strLable + strTmp(0) + "," + strTmp(1) + "," + strTmp(8) + "," + Left$(strTmp(3) + "00000000", 8) + strTmp(4) + "," + strTmp(5) + ","
                        strLable = strLable + strTmp(6) + "," + strTmp(7)
                    End If
                    '��ʼ��ӡTXT�ļ���ָ��λ��
                    Call PrintLable(strFileName, strLable, Trim(txt(2).Text)) '����
                End If
                If Opt(1).Value = True Then     '�ں�
                    If strShipToCust = "2000561" Then   '��LG�ı�ǩ
                        .Col = 2    '���
                        strFileName = Trim$(.Text)
                        .Col = 3    'DVC
                        strTmp(0) = Trim$(.Text)
                        .Col = 4    'Wafer Lot
                        strTmp(1) = Trim$(.Text)
                        .Col = 5    'Assy Lot
                        strTmp(2) = Trim$(.Text)
                        .Col = 6    'Date Code
                        strTmp(3) = Trim$(.Text)
                        .Col = 7    'Qty
                        strTmp(4) = Trim$(.Text)
                        '�����ǩ���
                        strLable = strTmp(0) + "," + strTmp(1) + "," + strTmp(2) + "," + strTmp(3) + "," + strTmp(4)
                    Else
                        .Col = 2    '���
                        strFileName = Trim$(.Text)
                        .Col = 3    'Part No
                        strTmp(0) = Trim$(.Text)
                        .Col = 4    'SPECIFICATION
                        strTmp(1) = Trim$(.Text)
                        .Col = 5    'PO Type
                        strTmp(6) = Trim$(.Text)
                        If InStr(strTmp(6), "E2") > 0 Then
                            strTmp(2) = "E2"
                        Else
                            strTmp(2) = ""
                        End If
                        .Col = 6    'Qty
                        strTmp(3) = Trim$(.Text)
                        .Col = 7    'VenDor P/N
                        strTmp(4) = Trim$(.Text)
                        .Col = 8    'VenDor Code
                        strTmp(5) = Trim$(.Text)
                        '�����ǩ���
                        strLable = strTmp(0) + strTmp(5) + strTmp(2) + Right$("000000" + strTmp(3), 6) + ","
                        strLable = strLable + strTmp(0) + "," + strTmp(1) + "," + strTmp(6) + "," + strTmp(3) + ","
                        strLable = strLable + strTmp(4) + "," + strTmp(5)
                    End If
                    '��ʼ��ӡTXT�ļ���ָ��λ��
                    Call PrintLable(strFileName, strLable, Trim(txt(3).Text)) '�ں�
                End If
                If Opt(2).Value = True Then     '����
                    .Col = 2    '���
                    strFileName = Trim$(.Text)
                    For j = 3 To .MaxCols
                        .Col = j
                        strLable = strLable + Trim$(.Text) + ","    'ƴ�ӱ�ǩ
                    Next
                    strLable = Left$(strLable, Len(strLable) - 1)   'ȥ�����һ������
                    '��ʼ��ӡTXT�ļ���ָ��λ��
                    Call PrintLable(strFileName, strLable, Trim(txt(4).Text)) '����
                End If
                    If Opt(3).Value = True Then      '����
                    .Col = 4    '���
                    strFileName = Trim$(.Text)
                    For j = 2 To .MaxCols
                        .Col = j
                        strLable = strLable + Trim$(.Text) + ","    'ƴ�ӱ�ǩ
                    Next
                    strLable = Left$(strLable, Len(strLable) - 1)   'ȥ�����һ������
                    '��ʼ��ӡTXT�ļ���ָ��λ��
                    Call PrintLable(strFileName, strLable, "\\10.160.1.14\BarCode\37\37DIE2-2\") '����
                End If
                  If Opt(4).Value = True Then      '����
                    .Col = 9    '���
                    strFileName = Trim$(.Text)
                    For j = 2 To .MaxCols
                        .Col = j
                        strLable = strLable + Trim$(.Text) + ","    'ƴ�ӱ�ǩ
                    Next
                    strLable = Left$(strLable, Len(strLable) - 1)   'ȥ�����һ������
                    '��ʼ��ӡTXT�ļ���ָ��λ��
                    Call PrintLable(strFileName, strLable, "\\10.160.1.14\BarCode\37\37DIE2-1\") '����
                End If
                
                '����ӡ�������ɾ��
                .DeleteRows i, 1
                .MaxRows = .MaxRows - 1
            End If
        Next

    End With
    
    MsgBox "��ӡ�ɹ���"

End Sub
'��ȡ��ǩ���
Private Function GetLableXH(strLot As String) As String
Dim strSql          As String
Dim Rs              As New adodb.Recordset
Dim strXH           As String
Dim intCount        As Integer
Dim strLot1          As String
    
    If strLot = "" Then Exit Function
    intCount = 0
    strLot1 = Replace(strLot, "M", "")
    strSql = "SELECT dbo.F_GetPrintXH('" & strLot1 & "') ���"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If Not Rs.EOF Then
        strXH = Trim$("" & Rs!���)
        If strXH <> "" Then '����еõ���ţ��͸�������
            strSql = "Update erpdata..tblSysIncrement Set Para='" & strXH & "',ICount=ICount+1 Where Kind='" & strLot & "'"
            INIadoCon2.Execute strSql, intCount
            If intCount <= 0 Then   '��ʾ�����ڴ�LOT��Ϣ���Ͳ���һ��
                strSql = "Insert Into erpdata..tblSysIncrement(Kind,Para,ICount) Values('" & strLot & "','" & strXH & "',1)"
                INIadoCon2.Execute strSql
            End If
        End If
    End If
    Rs.Close
    
    GetLableXH = strXH  '��ֵ��ȥ
    
End Function
'��ǩ��ϴ�ӡ
Private Sub PrintLable(strFileName As String, strTxt As String, strTxtPath As String)
Dim i               As Long
Dim strSql          As String
Dim Rs              As New adodb.Recordset
    
    '���ù���
    Call PrintLabelTxt(strFileName, strTxt, strTxtPath)
    '���´�ӡ��Ǻ�ʱ��
    If Opt(0).Value = True Then     '����
        strSql = "Update erpdata..TblTSV_Tray_details Set PrintFlag=1,PrintTime=getdate() Where TRAYQBOXNUMBER='" & strFileName & "'"
        INIadoCon2.Execute strSql
    End If
    If Opt(1).Value = True Then     '�ں�
        strSql = "Update erpdata..TblTSV_INBOX_DETAILS Set PrintFlag=1,PrintTime=getdate() Where CONTAINERNAME='" & strFileName & "'"
        INIadoCon2.Execute strSql
    End If
    If Opt(2).Value = True Then     '����
        strSql = "Update erpdata..TblTSV_OutBOX_DETAILS Set PrintFlag=1,PrintTime=getdate() Where CONTAINERNAME='" & strFileName & "'"
        INIadoCon2.Execute strSql
    End If
End Sub
'2016-09-08 mwl add дTXT��ǩ�ļ�
Private Sub PrintLabelTxt(FileName As String, msgTxt As String, dirtemp As String)
'�ж�txt�ļ��Ƿ���ڣ��粻���ڣ�����
Dim fileNameTemp        As String
Dim dirNameTemp         As String
Dim fileTemp            As String

    dirNameTemp = dirtemp
    fileNameTemp = Replace(FileName, "'", "") & ".txt"
    fileTemp = dirNameTemp & fileNameTemp
    
    Open fileTemp For Output As #1   'ֱ�Ӹ���
    Print #1, msgTxt
    Close #1

End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Fra(2).Move C_Left, Fra(2).Top, Me.ScaleWidth - C_Left, Fra(2).Height
    Fra(0).Move C_Left, Fra(0).Top, Fra(0).Width, Me.ScaleHeight - Fra(2).Height
    Fra(1).Move Fra(1).Left, Fra(1).Top, Me.ScaleWidth - Fra(0).Width - C_Top, Me.ScaleHeight - Fra(2).Height
    Fps(0).Move C_Left, Fps(0).Top, Fra(1).Width - C_Top, Me.ScaleHeight - Fra(2).Height - 3 * C_Top
End Sub
Private Sub Form_Load()

    '��ʼ���ؼ�
    InitCtrl
    
End Sub

'��ʼ���ؼ�
Private Sub InitCtrl()
Dim i                   As Integer
Dim strSql              As String
Dim Rs                  As New adodb.Recordset
    
    strdjbh = ""
    '���ص�������
    strSql = "SELECT Delivery,MAX(id) FROM erpbase..tblCustomerShippingUp " & _
             " WHERE Flag='Y' AND customershortname='37' " & _
             " GROUP BY Delivery ORDER BY 2 Desc"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    cmbDN.Clear
    If Not Rs.EOF Then
        Do While Not Rs.EOF
            cmbDN.AddItem Trim$("" & Rs!Delivery)
            Rs.MoveNext
        Loop
    End If
    Rs.Close
    '��ʼ��FPS
    InitFps
    
    chk.Value = 0
    DTP(0).Value = Format(Now(), "YYYY/MM/DD")
    DTP(1).Value = Format(Now(), "YYYY/MM/DD")
    Fra(3).Visible = False
   
End Sub
'��ʼ��FPS�ؼ�
Public Sub InitFps()
Dim i                   As Integer
    'Fps��ʼ��
    With Fps(0)
        .ReDraw = False
        .DAutoHeadings = True
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsBest
        .MaxRows = 0
        .ColsFrozen = 2
        .ButtonDrawMode = 1
        .Row = -1
        .Col = -1
        .Lock = True
        .TypeHAlign = TypeHAlignLeft
        .TypeVAlign = TypeVAlignCenter
        '�趨������
        .Col = FpsDetail.e_Choose   'ѡ��
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '�趨�п�
        .ColWidth(-1) = 10
        .ColWidth(FpsDetail.e_Choose) = 4
        .RowHeight(-1) = 10
        '�趨�Ƿ�����
        .UserColAction = UserColActionSort
        For i = 1 To .MaxCols
            .Col = i
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
        Next
        .ZOrder
        .ReDraw = True
    End With
End Sub
Private Sub Fps_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i           As Long
Dim j           As Integer
    
    '�����ѡ��ĵ��Ŷ�ѡ��
    If Row < 1 Then Exit Sub
    If Col <> 1 Then Exit Sub
    With Fps(0)
'        .Col = FpsDetail.e_Choose
'        For i = 1 To .MaxRows
'            .Row = i
'            If i <> Row Then
'                .Col = FpsDetail.e_Choose
'                If Val(.Value) = 1 Then
''                    .Value = 0
'                    .Col = -1
'                    .ForeColor = vbBlack
'                End If
'            End If
'        Next

        .Col = FpsDetail.e_Choose
        .Row = Row
        .Value = Abs(Val(.Value) - 1)
    End With
    
End Sub

'У������
Private Function CheckData() As Boolean
Dim i               As Long
Dim intCount        As Integer

    CheckData = False
    
    intCount = 0
    
    With Fps(0)
        If .MaxRows <= 0 Then
            MsgBox "û���κ�����,���Ȳ�ѯ��", vbInformation, "��ʾ"
            Exit Function
        End If
        '���Ƿ���ѡ��
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsDetail.e_Choose  'ѡ��
            If .Value = 1 Then
                intCount = intCount + 1
            End If
        Next
    End With
    '--------------------------
    If intCount <= 0 Then
        MsgBox "û��ѡ���κ����ϣ�", vbInformation, "��ʾ"
        Exit Function
    End If
    'У���Ƿ��б�ǩ·��
    If Opt(0).Value = True Then '����
        If Trim(txt(2).Text) = "" Then
            MsgBox "û���趨�˿ͻ��ľ��̱�ǩ·��������ϵϵͳ����Ա��", vbInformation, "��ʾ"
            Exit Function
        End If
    End If
    If Opt(1).Value = True Then '�ں�
        If Trim(txt(3).Text) = "" Then
            MsgBox "û���趨�˿ͻ����ںб�ǩ·��������ϵϵͳ����Ա��", vbInformation, "��ʾ"
            Exit Function
        End If
    End If
    If Opt(2).Value = True Then '����
        If Trim(txt(4).Text) = "" Then
            MsgBox "û���趨�˿ͻ��������ǩ·��������ϵϵͳ����Ա��", vbInformation, "��ʾ"
            Exit Function
        End If
    End If
    CheckData = True
End Function
'���̣����䣬����ı仯
Private Sub Opt_Click(Index As Integer)
    If Index = 0 Then
        Fps(0).MaxRows = 0
        chk.Value = 0
        chkChoose.Value = 0
    ElseIf Index = 1 Then
        Fps(0).MaxRows = 0
        chk.Value = 0
        chkChoose.Value = 0
    ElseIf Index = 2 Then
        Fps(0).MaxRows = 0
        chk.Value = 0
        chkChoose.Value = 0
    End If
End Sub
