VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmSemtech_Report 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Semtech�����ѯ"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
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
   ScaleHeight     =   7695
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra 
      ForeColor       =   &H000000FF&
      Height          =   7335
      Index           =   1
      Left            =   3480
      TabIndex        =   17
      Top             =   840
      Width           =   9615
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   3255
         Index           =   0
         Left            =   120
         TabIndex        =   18
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
         SpreadDesigner  =   "FrmSemtech_Report.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "��ѯ����"
      ForeColor       =   &H00FF0000&
      Height          =   7335
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   3495
      Begin VB.ComboBox cmbCombo1 
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
         Index           =   1
         ItemData        =   "FrmSemtech_Report.frx":04E1
         Left            =   1080
         List            =   "FrmSemtech_Report.frx":04E8
         TabIndex        =   19
         Top             =   2400
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   0
         Left            =   1080
         TabIndex        =   15
         Top             =   1080
         Width           =   2355
      End
      Begin VB.ComboBox cmbCombo1 
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
         Index           =   0
         ItemData        =   "FrmSemtech_Report.frx":04F7
         Left            =   1080
         List            =   "FrmSemtech_Report.frx":04F9
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2355
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   204144643
         CurrentDate     =   41387
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   204144643
         CurrentDate     =   41387
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ����"
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
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   2460
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
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
         TabIndex        =   9
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ������"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������ĩ"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job      No"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   720
      End
   End
   Begin VB.Frame Fra 
      Height          =   855
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.CommandButton cmdInvRpt 
         Caption         =   "һ��������汨��"
         Height          =   360
         Left            =   8640
         MaskColor       =   &H8000000F&
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin MSComDlg.CommonDialog Com 
         Left            =   11760
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "�ϴ��ļ�"
         Height          =   480
         Left            =   11160
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "��������"
         Height          =   360
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�� ��"
         Height          =   360
         Left            =   5760
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdExprot 
         Caption         =   "������ǰ����"
         Height          =   360
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "��  ѯ"
         Height          =   360
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   990
      End
   End
End
Attribute VB_Name = "FrmSemtech_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strdjbh         As String
Dim strdjbh1         As String
Dim DirShare        As String
Dim DirFileShare    As String
Dim DirInvRpt       As String
Dim order           As String
Dim RsClone         As New ADODB.Recordset
Const C_Left = 60
Const C_Top = 120

Private Enum FpsDetail
    e_Choose = 1
    e_DJBH = 2
    e_Cust = 3
    e_YDH = 7
End Enum
'��������ֵ
Private Function GetExcelName(ByVal strTitle As String) As String
Dim strSql          As String
Dim Rs              As New ADODB.Recordset
Dim strExFileName   As String
Dim strCurDate      As String
    
    If strTitle = "output Invoice" Then
        strCurDate = Format(Date, "YY-MMDD")
    Else
        strCurDate = Format(Date, "YYYYMMDD")
    End If
    strSql = "select nvl(max(para_3),0)+1 para from tblsys_parameter where sysname='TSVSYS' and kind='Semtech����' and para_1='" & strTitle & "' and para_2='" & strCurDate & "'"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not Rs.EOF Then
        strExFileName = strCurDate + "_" + Trim$("" & Rs!PARA)
        strSql = "Update tblsys_parameter set para_2='" & strCurDate & "',para_3='" & Trim$("" & Rs!PARA) & "' where sysname='TSVSYS' and kind='Semtech����' and para_1='" & strTitle & "'"
        Cnn.Execute strSql
    Else
        strExFileName = strCurDate + "_1"
        strSql = "Update tblsys_parameter set para_2='" & strCurDate & "',para_3='" & Trim$("" & Rs!PARA) & "' where sysname='TSVSYS' and kind='Semtech����' and para_1='" & strTitle & "'"
        Cnn.Execute strSql
    End If
    Rs.Close
    
    If strTitle = "output Invoice" Then
        GetExcelName = "HTKS_SEDC" & strExFileName
    Else
        GetExcelName = strTitle & "_" & strExFileName
    End If
    
End Function

Private Sub cmbCombo1_Click(Index As Integer)
'Dim strSql              As String
'Dim Rs                  As New ADODB.Recordset

    If Index = 0 Then
        If cmbCombo1(0).Text = "��汨��" Then
            '���زֿ�
            cmbCombo1(1).Clear
            cmbCombo1(1).AddItem "����"
            cmbCombo1(1).AddItem "1000"
            cmbCombo1(1).AddItem "6000"
            cmbCombo1(1).AddItem "7000"
            cmbCombo1(1).AddItem "8000"
            cmbCombo1(1).AddItem "9000"
            cmbCombo1(1).AddItem "Scrap"
            cmbCombo1(1).ListIndex = 0
        Else
            cmbCombo1(1).Clear
        End If
    End If
End Sub

Private Sub CmdExit_Click() '�˳�
    Unload Me
End Sub

Private Sub cmdExprot_Click()
Dim strExportName           As String

    'У������
    If Fps(0).MaxRows <= 0 Then
        MsgBox "û�пɵ��������ݣ�", vbInformation, "��ʾ"
        Exit Sub
    End If
    '��������
    If cmbCombo1(0).ListIndex = 6 Or cmbCombo1(0).ListIndex = 7 Or cmbCombo1(0).ListIndex = 8 Then '��汨��,SMTCList,Shipped
        strExportName = Trim(Fra(1).Caption)
        If cmbCombo1(1).Text <> "" Then
            strExportName = Trim(Fra(1).Caption) + "-" + Trim(cmbCombo1(1).Text)
        End If
    Else
        strExportName = GetExcelName(Trim(Fra(1).Caption))
    End If

    If Not ExportFpspreadToExcel(Fps(0), strExportName, strExportName) Then Exit Sub
    
End Sub

Private Sub cmdInvRpt_Click()
'һ��������汨��ָ���ļ�����
Dim strSql                  As String
Dim strSqlDetail            As String
Dim Rs                      As New ADODB.Recordset
Dim i                       As Integer
Dim strFileName             As String
Dim strMsg                  As String
    
    If MsgBox("ȷ��Ҫ������?", vbYesNo Or vbQuestion Or vbDefaultButton2, "��ʾ") = vbNo Then
        Exit Sub
    End If
    '�������ǿ��
    strSql = "SELECT RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Reel_Size,Remark" & _
             ",Move_in_Date FROM Vw_InvStockRptFor37 Where 2>1 "
    For i = 0 To cmbCombo1(1).ListCount
        If cmbCombo1(1).List(i) <> "����" And cmbCombo1(1).List(i) <> "" Then
            If cmbCombo1(1).List(i) = "9000" Then
               
'            strSql = "SELECT b.fab_conv_id as [Wafer Type],b.mpn_desc as [Assy Part#],a.������ as [Fab Lot], " & _
'            " right(replace(rtrim(a.���̿����),'+',''),2) ID,b.date_code as [D/C],b.die as [QTY(���غ������)]," & _
'            " b.test_mtrl_desc as [Job#],b.bag as Bag#,a.��� as Comment,b.alternatename as [HT Part#]" & _
'            " FROM erpdata..tblStockNumsub a," & _
'            " (SELECT  substring(cast(datepart(year,ora.create_date)as nvarchar(20))+substring(CAST ('100'+datepart(week,ora.create_date) as nvarchar(20)),2,2),3,4) as date_code," & _
'            " ora.waferid,ora.firstname ,ora.test_mtrl_desc,ora.mpn_desc,ora.bag,ora.fab_conv_id,ora.die,ora.alternatename FROM " & _
'            " OPENQUERY(ORACLEDB, 'SELECT d.fab_conv_id,a.waferid,a.die," & _
'            " a.create_date+6 as create_date,b.firstname ,d.test_mtrl_desc,d.mpn_desc,get_37bagid(b.containername) bag,e.alternatename" & _
'            " FROM weight37 a,container b ,mappingdatatest c,customeroitbl_test d,product e" & _
'            " where a.waferid||''-A'' = b.containername  and a.waferid = c.substrateid " & _
'            " and e.productid=b.productid and c.filename = d.id ' ) ora ) b where a.�ⷿ��� IN('44','45')" & _
'            " and a.���̿����=b.waferid "
            strSql = " SELECT * FROM DBO.Vw_InvStockRptFor37By9000 "
            strSqlDetail = ""
            Else
              strSql = "SELECT RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Reel_Size,Remark" & _
              ",Move_in_Date FROM Vw_InvStockRptFor37 Where 2>1 "
              strSqlDetail = " And �ֿ�����='" & cmbCombo1(1).List(i) & "'"
            End If
            If Rs.State = adStateOpen Then Rs.Close
            Rs.open strSql + strSqlDetail, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
            If Not Rs.EOF Then '��ʾ�����ݲŵ�������
                strFileName = Format(Now(), "YYMMDD_HHMM") + "_" + Replace(cmbCombo1(1).List(i), "Scrap", "SCR")
                strMsg = strMsg + DirInvRpt + "\" + strFileName + vbCrLf                '��ʾ��Ϣ
                RsExporToExcel Rs, cmbCombo1(1).List(i), strFileName                    '������Excel
            End If
            Rs.Close
        End If
    Next
    '����Shipped
    strSql = "SELECT RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Remark,SO as [DN#],SHIPPED_DATE,Cust_Name " & _
             " FROM Vw_InvShippedRptFor37 "
'             " WHERE SHIPPED_DATE>='" & DateAdd("m", -1, Format(Now(), "YYYY-MM-DD")) & "' and SHIPPED_DATE<'" & DateAdd("d", 1, Format(Now(), "YYYY-MM-DD")) & "' "
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If Not Rs.EOF Then '��ʾ�����ݲŵ�������
        strFileName = Format(Now(), "YYMMDD_HHMM") + "_SHIPPED"
        strMsg = strMsg + DirInvRpt + "\" + strFileName + vbCrLf                '��ʾ��Ϣ
        RsExporToExcel Rs, "SHIPPED", strFileName                               '������Excel
    End If
    Rs.Close
    
    MsgBox "�����ɹ��������ļ�·��Ϊ��" + vbCrLf + strMsg
    
End Sub

Private Sub cmdReport_Click() '��������
Dim strExportName           As String

    If Fps(0).MaxRows <= 0 Then
        MsgBox "û�пɵ��������ݣ�", vbInformation, "��ʾ"
        Exit Sub
    End If
    '��������
    strExportName = GetExcelName(Trim(Fra(1).Caption))
    
    If cmbCombo1(0).ListIndex = 0 Then
        Call SEDCExportPrintExcel(RsClone, strExportName)                     'SEDC����
    ElseIf cmbCombo1(0).ListIndex = 1 Then
        Call InputPackinglistExportPrintExcel(RsClone, strExportName)         'outputPackinglist����
    ElseIf cmbCombo1(0).ListIndex = 2 Then
        Call InputInvoiceExportPrintExcel(RsClone, strExportName)             'outputInvoice����
    ElseIf cmbCombo1(0).ListIndex = 3 Then
        Call Daily_InvExportPrintExcel(RsClone, strExportName)                'Daily_inventory_report
    ElseIf cmbCombo1(0).ListIndex = 4 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel(order, strExportName)      'Shipping Packinglist
    ElseIf cmbCombo1(0).ListIndex = 5 Then
        If Not CheckData Then Exit Sub
        Call ShippingInvoiceExportPrintExcel(order, strExportName)          'Shipping invoice
    End If
End Sub

Private Sub cmdSearch_Click() '��ѯ����
Dim i                   As Long
Dim strSql              As String
Dim Rs                  As New ADODB.Recordset

    '��ʼ��FPS
    InitFps
    '---------------------------------------------
    If cmbCombo1(0).ListIndex = 0 Then  'SEDC
    
              strSql = "select distinct 0 as ѡ��,a.containername,a.creationtimestamp as receive_date,to_char(ibwo.erpcreatedate,'yyyyww') testdc,'' as invlocation " & _
                ",wo.mpn_desc as device_name,wo.SOURCE_BATCH_ID as job_no,conn.firstname as lot_no,a.moveinqty as qty,wo.date_code,'' as invcomment " & _
                ",'new packing' as remark,'' as so,'' as invoice,Get_37Inv_MergeDetails(a.containername,wo.mtrl_num) as merge " & _
                "from a_wiplothistory a ,a_wiplotdetailshistory b ,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf " & _
                ",customeroitbl_test wo " & _
                "where a.specname='5272' and b.wiplothistoryid=a.wiplothistoryid " & _
                "and conn.containername=a.containername " & _
                "and waf.waferscribenumber=substr(conn.containername,1,instr(conn.containername,'-A')-1) " & _
                "and mfg.mfgordername=waf.workordername and ibwo.ordername=mfg.mfgordername " & _
                "and wo.source_batch_id=waf.wafernumber " & _
                " and wo.customershortname='37' and a.containername not like '%-F%' " & _
                "And a.creationtimestamp>='" & DTP(0).Value & "' And a.creationtimestamp<'" & DTP(1).Value + 1 & "'"
                
        If txt(1).Text <> "" Then
            strSql = strSql & " And wo.SOURCE_BATCH_ID='" & Trim(txt(1).Text) & "'"
        End If
        
    ElseIf cmbCombo1(0).ListIndex = 1 Then  'output packing list
     
        strSql = " select ѡ��,creationtimestamp outdate,po_num,po_item,PKG,DEVICE,lot_no,Job_No,Sublot_No,date_code,sum( qty) as qty,sum(Price) as Price,sum(USD) as USD,Test_Reject,Merge_in_Job,RMA_Number,mtrl_num from (" & _
                " select distinct 0 as ѡ��,a.containername, wo.po_num, wo.po_item,'' PKG,wo.mpn_desc DEVICE,wo.MTRL_NUM lot_no,waf.wafernumber Job_No," & _
                " conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No, wo.date_code,a.moveinqty qty, 0 as Price,a.moveinqty * 0 as USD," & _
                "  '' as Test_Reject, Get_37Inv_MergeDetails(a.containername,wo.MTRL_NUM) as Merge_in_Job,'' as RMA_Number, wo.mtrl_num,a.creationtimestamp " & _
                "  from a_wiplothistory a, a_wiplotdetailshistory b,container  conn,mfgorder  mfg,ib_wohistory  ibwo,a_lotwafers waf, customeroitbl_test wo " & _
                " where a.specname = '5272' and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
                "  and waf.waferscribenumber = substr(conn.containername, 1, instr(conn.containername, '-A') - 1) " & _
                "   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername and wo.source_batch_id = waf.wafernumber and status<>2 " & _
                "   and wo.customershortname = '37' and a.containername not like '%-F%' And a.creationtimestamp >= '" & DTP(0).Value & "'" & _
                "   And a.creationtimestamp < '" & DTP(1).Value & "' and   a.containername<>'10001-A-01' and not  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername=a.containername)  " & _
                " union select distinct 0 as ѡ��, a.containername, wo.po_num,wo.po_item,'' PKG, wo.mpn_desc DEVICE, wo.MTRL_NUM lot_no, waf.wafernumber as job_no, " & _
                " conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No,wo.date_code, waf.ndpw qty,0 as Price, a.moveinqty * 0 as USD," & _
                " '' as Test_Reject, Get_37Inv_MergeDetails(a.containername,wo.MTRL_NUM) as Merge_in_Job,'' as RMA_Number , wo.mtrl_num,a.creationtimestamp" & _
                "  from a_wiplothistory  a, a_wiplotdetailshistory b, container conn,mfgorder mfg, ib_wohistory ibwo, a_lotwafers  waf, customeroitbl_test  wo " & _
                " where a.specname = '5272' and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername " & _
                "   and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername and wo.source_batch_id = waf.wafernumber " & _
                "   and wo.customershortname = '37' and a.containername not like '%-F%' And a.creationtimestamp >= '" & DTP(0).Value & "' " & _
                "   And a.creationtimestamp < '" & DTP(1).Value & "' and waf.containerid=conn.containerid  and  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername = a.containername)  " & _
                " )X Where 2>1 "
                
        If txt(1).Text <> "" Then
            strSql = strSql & " And Job_No='" & Trim(txt(1).Text) & "'"
        End If
        strSql = strSql & " group by ѡ��,po_num,po_item,PKG,DEVICE,lot_no,Job_No,Sublot_No,date_code,Test_Reject,Merge_in_Job,RMA_Number,mtrl_num,creationtimestamp order by Sublot_No,Merge_in_Job desc"
    
    ElseIf cmbCombo1(0).ListIndex = 2 Then  'output invoice
      
      strSql = " select ѡ��,creationtimestamp outdate,po_num,po_item, PKG, DEVICE,lot_no,job_no,Sublot_No,date_code, sum(qty), Price,sum(USD)  from (" & _
                "select distinct 0 as ѡ��,a.containername,wo.po_num,wo.po_item,'' PKG,wo.mpn_desc DEVICE,wo.MTRL_NUM lot_no,wo.SOURCE_BATCH_ID job_no " & _
                ",conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No,wo.date_code,a.moveinqty qty,wo.t_price as Price,a.moveinqty*wo.t_price as USD,a.creationtimestamp " & _
                "from a_wiplothistory a ,a_wiplotdetailshistory b ,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf " & _
                ",customeroitbl_test wo " & _
                "where a.specname='5272' and b.wiplothistoryid=a.wiplothistoryid " & _
                "and conn.containername=a.containername " & _
                "and waf.waferscribenumber=substr(conn.containername,1,instr(conn.containername,'-A')-1) " & _
                "and mfg.mfgordername=waf.workordername and ibwo.ordername=mfg.mfgordername " & _
                "and wo.source_batch_id=waf.wafernumber " & _
                " and wo.customershortname='37' and a.containername not like '%-F%' " & _
                "And a.creationtimestamp>='" & DTP(0).Value & "' And a.creationtimestamp<'" & DTP(1).Value & "'" & _
                "and a.containername<>'10001-A-01'and not  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername=a.containername)   union " & _
                " select  distinct 0 as ѡ��,a.containername,wo.po_num,wo.po_item,'' PKG,wo.mpn_desc DEVICE,wo.MTRL_NUM lot_no,waf.wafernumber job_no, " & _
                " conn.firstname||Get_37Inv_MergeStatus(a.containername) Sublot_No,wo.date_code,waf.ndpw qty,wo.t_price as Price,waf.ndpw* wo.t_price as USD,a.creationtimestamp " & _
                " from a_wiplothistory a,a_wiplotdetailshistory b,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf,customeroitbl_test wo " & _
                " where a.specname = '5272' and b.wiplothistoryid = a.wiplothistoryid and conn.containername = a.containername" & _
                " and mfg.mfgordername = waf.workordername and ibwo.ordername = mfg.mfgordername " & _
                " and wo.source_batch_id = waf.wafernumber and wo.customershortname = '37' and a.containername not like '%-F%' " & _
                " And a.creationtimestamp >= '" & DTP(0).Value & "' And a.creationtimestamp < '" & DTP(1).Value & "' and waf.containerid=conn.containerid   and  exists (select 1 from cus37_5272qtymergV2 mer where mer.containername = a.containername) " & _
                " )X Where 2>1 "
                
        If txt(1).Text <> "" Then
            strSql = strSql & " And job_no='" & Trim(txt(1).Text) & "'"
        End If
        strSql = strSql & " group by ѡ��,po_num,po_item, PKG, DEVICE, lot_no, job_no ,Sublot_No,date_code, Price,creationtimestamp order by Sublot_No"
    
    ElseIf cmbCombo1(0).ListIndex = 3 Then  ' Daily Inv
             
        strSql = " select ѡ��,max(RECEIVE_DATE) as RECEIVE_DATE,TESTDC,LOCATION,DEVICE_NAME,job_no,HTLOT_NO,sum(qty) as qty,date_code,CCOMMENT,Reel_Size,Remark,Move_in_Date from (" & _
                "select distinct 0 as ѡ��,a.containername,a.creationtimestamp as RECEIVE_DATE,to_char(ibwo.erpcreatedate,'yyww') TESTDC,'' as LOCATION " & _
                ",wo.mpn_desc DEVICE_NAME,wo.SOURCE_BATCH_ID||Get_37Inv_MergeStatus(a.containername) job_no " & _
                ",conn.firstname||Get_37Inv_MergeStatus(a.containername) HTLOT_NO,a.moveinqty qty,wo.date_code,'' CCOMMENT,'NEW PACKING' Reel_Size,'' Remark,'' Move_in_Date " & _
                "from a_wiplothistory a ,a_wiplotdetailshistory b ,container conn,mfgorder mfg,ib_wohistory ibwo,a_lotwafers waf " & _
                ",customeroitbl_test wo " & _
                "where a.specname='5272' and b.wiplothistoryid=a.wiplothistoryid " & _
                "and conn.containername=a.containername " & _
                "and waf.waferscribenumber=substr(conn.containername,1,instr(conn.containername,'-A')-1) " & _
                "and mfg.mfgordername=waf.workordername and ibwo.ordername=mfg.mfgordername " & _
                "and wo.source_batch_id=waf.wafernumber " & _
                " and wo.customershortname='37' and a.containername not like '%-F%'  and status<>'2' " & _
                "And a.creationtimestamp>='" & DTP(0).Value & "' And a.creationtimestamp<'" & DTP(1).Value + 1 & "'" & _
                " ) X Where 2>1 "
                
        If txt(1).Text <> "" Then
            strSql = strSql & " And job_no='" & Trim(txt(1).Text) & "'"
        End If
        strSql = strSql & " group by ѡ��,TESTDC,LOCATION,DEVICE_NAME,job_no,HTLOT_NO,date_code,CCOMMENT,Reel_Size,Remark,Move_in_Date"
        
    ElseIf cmbCombo1(0).ListIndex = 4 Then  'shipping packing list (INNER JOIN tblCustomerOI d ON CASE WHEN charindex('M',RTRIM(c.batchnumber))>0 THEN LEFT(RTRIM(c.batchnumber),LEN(RTRIM(c.batchnumber))-1) ELSE RTRIM(c.batchnumber) END=d.SOURCE_BATCH_ID)
        strSql = "SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",���,�Ϻ�,mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,����,ë��,�ߴ� " & _
                 " FROM Vw_InvShippedPLFor37 a " & _
                 " WHERE ��������>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and ��������<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "' "
        If txt(0).Text <> "" Then
            strSql = strSql & " And ���ݱ��='" & Trim(txt(0).Text) & "'"
        End If
        If txt(1).Text <> "" Then
            strSql = strSql & " And batchnumber='" & Trim(txt(1).Text) & "'"
        End If
    ElseIf cmbCombo1(0).ListIndex = 5 Then  'shipping invoice (INNER JOIN tblCustomerOI d ON CASE WHEN charindex('M',RTRIM(c.batchnumber))>0 THEN LEFT(RTRIM(c.batchnumber),LEN(RTRIM(c.batchnumber))-1) ELSE RTRIM(c.batchnumber) END=d.SOURCE_BATCH_ID)
        strSql = "SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
                 ",city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",���,�Ϻ�,mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,����,AMount,customerPartNumber " & _
                 " FROM Vw_InvShippedInvoiceFor37 a " & _
                 " WHERE ��������>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and ��������<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "' "
        If txt(0).Text <> "" Then
            strSql = strSql & " And ���ݱ��='" & Trim(txt(0).Text) & "'"
        End If
        If txt(1).Text <> "" Then
            strSql = strSql & " And batchnumber='" & Trim(txt(1).Text) & "'"
        End If
    ElseIf cmbCombo1(0).ListIndex = 6 Then  '��汨��
        If cmbCombo1(1).Text = "9000" Then
            strSql = " SELECT * FROM DBO.Vw_InvStockRptFor37By9000 "
        Else
            strSql = "SELECT 0 ѡ��,RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Reel_Size,Remark" & _
                     ",Move_in_Date FROM Vw_InvStockRptFor37 Where 2>1 "
            If txt(1).Text <> "" Then
                strSql = strSql & " And JOB_NO='" & Trim(txt(1).Text) & "'"
            End If
            If cmbCombo1(1).Text <> "����" Then
                strSql = strSql & " And �ֿ�����='" & Trim(cmbCombo1(1).Text) & "'"
            End If
        End If
    ElseIf cmbCombo1(0).ListIndex = 7 Then  'SMTCList
        strSql = "SELECT 0 ѡ��,Invoice_No,Carton_No,PartName,LotID,QTY,Job_No,DATE_CODE " & _
                 " FROM Vw_InvShippedSMTCListFor37 " & _
                 " WHERE ��������>='" & DTP(0).Value & "' and ��������<'" & DTP(1).Value + 1 & "' "
        If txt(0).Text <> "" Then
            strSql = strSql & " And ���ݱ��='" & Trim(txt(0).Text) & "'"
        End If
        If txt(1).Text <> "" Then
            strSql = strSql & " And Job_No='" & Trim(txt(1).Text) & "'"
        End If
    ElseIf cmbCombo1(0).ListIndex = 8 Then  'Shipped
        strSql = "SELECT 0 ѡ��,RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Remark,SO as [DN#],SHIPPED_DATE,Cust_Name " & _
                 " FROM Vw_InvShippedRptFor37 " & _
                 " WHERE SHIPPED_DATE>='" & DTP(0).Value & "' and SHIPPED_DATE<'" & DTP(1).Value + 1 & "' "
        If txt(0).Text <> "" Then
            strSql = strSql & " And ���ݱ��='" & Trim(txt(0).Text) & "'"
        End If
        If txt(1).Text <> "" Then
            strSql = strSql & " And JOB_NO='" & Trim(txt(1).Text) & "'"
        End If
    End If
    '��ֵ��FRA(1)�� INIadoCon
    Fra(1).Caption = cmbCombo1(0).Text
    If Rs.State = adStateOpen Then Rs.Close
    If cmbCombo1(0).ListIndex = 0 Or cmbCombo1(0).ListIndex = 1 Or cmbCombo1(0).ListIndex = 2 Or cmbCombo1(0).ListIndex = 3 Then
        Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    Else
        Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    End If
    Fps(0).MaxRows = 0
    Set RsClone = Nothing
    If Not Rs.EOF Then
        Set RsClone = Rs.Clone '��¡һ�����ݵ���һ�����ݼ��У�Ϊ����ʹ��
        With Fps(0)
            .MaxRows = 0
            Set .DataSource = Rs
            .MaxRows = Rs.RecordCount
        End With
    End If
    Rs.Close
    '���⼸���������ӻ�����λ
    CalcTotal
    
End Sub
Private Sub CalcTotal()
'�������
Dim i                   As Long
Dim dblTotal            As Double
Dim colTotal            As Integer
    
    If cmbCombo1(0).ListIndex <> 6 And cmbCombo1(0).ListIndex <> 7 And cmbCombo1(0).ListIndex <> 8 Then Exit Sub
    
    dblTotal = 0
    colTotal = 0
    With Fps(0)
        If .MaxRows <= 0 Then Exit Sub
        For i = 1 To .MaxRows
            .Row = i
            If cmbCombo1(0).ListIndex = 6 Or cmbCombo1(0).ListIndex = 8 Then
                colTotal = 8
                .Col = colTotal
            Else
                colTotal = 6
                .Col = colTotal
            End If
            dblTotal = dblTotal + Val(Trim$(.Text))
        Next
        If dblTotal > 0 Then '��ʾ������
            .MaxRows = .MaxRows + 1
            .SetText colTotal, .MaxRows, dblTotal
        End If
        .DeleteCols FpsDetail.e_Choose, 1
        .MaxCols = .MaxCols - 1
    End With
    
End Sub

Private Sub cmdUpload_Click()
Dim strFilePath         As String
Dim strFileName         As String
Dim strSql              As String
Dim image_Data()        As Byte         'ͼƬ������
Dim Rs                  As New ADODB.Recordset
    '��ͼƬ
    Com.Filter = "�ϴ��ļ�(*.xls,*.xlsx)|*.xls;*.xlsx"
    Com.ShowOpen '�򿪶Ի���
    strFilePath = Trim(Com.FileName)  '����·��
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1) '�ļ���
    '��ʼ���浽���Ͽ�
    '����ת��Ϊ��
    Open strFilePath For Binary As #1
    ReDim image_Data(LOF(1) - 1)
    Get #1, , image_Data()
    Close #1
    '��ѯ�Ƿ񱣴����ͼƬ
    strSql = "Select * From TblPMC_PicInfo Where FileName='" & Trim$(strFileName) & "' For Update"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenKeyset, adLockOptimistic
    If Not Rs.EOF Then
        Rs("FileName") = strFileName
        Rs("FilePath") = strFilePath
        Rs("FileComent") = image_Data()
        Rs("Flag") = 1
        Rs.Update
    Else
        Rs.AddNew
        Rs("FileName") = strFileName
        Rs("FilePath") = strFilePath
        Rs("FileComent") = image_Data()
        Rs("Flag") = 1
        '�ǵ�������ݿ���txt��ŵ�·����txt��
        Rs.Update
    End If
    Rs.Close
    
    MsgBox "�ϴ��ɹ�", vbInformation, "��ʾ"
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Fra(2).Move C_Left, Fra(2).Top, Me.ScaleWidth - C_Left, Fra(2).Height
    Fra(0).Move C_Left, Fra(0).Top, Fra(0).Width, Me.ScaleHeight - Fra(2).Height
    Fra(1).Move Fra(1).Left, Fra(1).Top, Me.ScaleWidth - Fra(0).Width - C_Top, Me.ScaleHeight - Fra(2).Height
    Fps(0).Move C_Left, Fps(0).Top, Fra(1).Width - C_Top, Me.ScaleHeight - Fra(2).Height - 3 * C_Top
End Sub
Private Sub Form_Load()
    If gUserName = "07885" Then
        cmdUpload.Visible = True
    End If
    DirShare = App.Path & "\NewSemtechReport"               'ϵͳ·��
    DirFileShare = App.Path & "\SemtechExcelReport"         'ϵͳExcel�ļ�·��
    DirInvRpt = "C:\37-InventoryRpt"                        '��汨����·��
    '�ж��ļ����Ƿ����,�����ھʹ���
    If Dir(DirShare, vbDirectory) = "" Then
        MkDir DirShare                                      '�����ļ���
    End If
    If Dir(DirFileShare, vbDirectory) = "" Then
        MkDir DirFileShare                                  '�����ļ���
    End If
    If Dir(DirInvRpt, vbDirectory) = "" Then
        MkDir DirInvRpt                                     '�����ļ���
    End If
    '��ʼ���ؼ�
    InitCtrl
    
End Sub
Private Sub CheckXls()
Dim strSql              As String
Dim Rs                  As New ADODB.Recordset
Dim image_filename      As String
Dim temp_image()        As Byte
Dim i                   As Integer
    '�趨���״̬
    Screen.MousePointer = 0
    strSql = "Select * From TblPMC_PicInfo Where Flag=1 Order by Create_Date"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not Rs.EOF Then
        For i = 1 To Rs.RecordCount
            '����ͼƬ
            temp_image = Rs("FileComent")
            image_filename = DirShare & "\" & Rs("FileName")
            Open image_filename For Binary As #1
            Put #1, , temp_image()
            Close #1
            Rs.MoveNext
        Next
    End If
    Rs.Close

End Sub
'��ʼ���ؼ�
Private Sub InitCtrl()
Dim i                   As Integer
Dim strSql              As String
Dim Rs                  As New ADODB.Recordset
    
    strdjbh = ""
    '���ص�������
    strSql = "select para_1 from tblsys_parameter where sysname='TSVSYS' and kind='Semtech����' order by id "
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    cmbCombo1(0).Clear
    If Not Rs.EOF Then
        Do While Not Rs.EOF
            cmbCombo1(0).AddItem Trim$("" & Rs!para_1)
            Rs.MoveNext
        Loop
        cmbCombo1(0).ListIndex = 0
    End If
    Rs.Close
'    '���زֿ�
'    cmbCombo1(1).Clear
'    cmbCombo1(1).AddItem "����"
'    cmbCombo1(1).AddItem "1000"
'    cmbCombo1(1).AddItem "6000"
'    cmbCombo1(1).AddItem "7000"
'    cmbCombo1(1).AddItem "8000"
'    cmbCombo1(1).AddItem "Scrap"
'    cmbCombo1(1).ListIndex = 0
    '��ʼ��FPS
    InitFps
    
   DTP(0).Value = Format(Now() - 1, "YYYY-MM-DD")
   DTP(1).Value = Format(Now(), "YYYY-MM-DD")
   '���ģ��
   CheckXls
   
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
Dim strTmp      As String
    
    '�����������⴦��
    If cmbCombo1(0).ListIndex = 6 Or cmbCombo1(0).ListIndex = 7 Or cmbCombo1(0).ListIndex = 8 Then Exit Sub
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
'        strDJBH = ""
        If Val(.Value) = 1 Then
            '������һ���ĵ���ѡ����
            .Col = FpsDetail.e_DJBH
            .Row = Row
            strTmp = Trim$(.Text)
'            strDJBH = Trim$(.Text) '���õĵ��ݱ�ţ��ڵ�����ӡʱ���õ�
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsDetail.e_DJBH
                If Trim$(.Text) = strTmp Then
                    .Col = FpsDetail.e_Choose
                    .Value = 1
                    .Col = -1
                    .ForeColor = &HFF8080
                End If
            Next
            
            order = strTmp & "'" & "," & "'" & order
            
        Else
            '������һ���ĵ���ѡ����
            .Col = FpsDetail.e_DJBH
            .Row = Row
            strTmp = Trim$(.Text)
'            strDJBH = Trim$(.Text) '���õĵ��ݱ�ţ��ڵ�����ӡʱ���õ�
            For i = 1 To .MaxRows
                .Row = i
                .Col = FpsDetail.e_DJBH
                If Trim$(.Text) = strTmp Then
                    .Col = FpsDetail.e_Choose
                    .Value = 0
                    .Col = -1
                    .ForeColor = vbBlack
                End If
            Next
        End If
        
    End With
    
End Sub

'У������
Private Function CheckData() As Boolean
Dim i               As Integer
Dim intCount        As Integer
Dim strCust         As String

    CheckData = False
    
    strdjbh = ""     '--���ݱ�ż�¼
    strCust = ""
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
                .Col = FpsDetail.e_DJBH '���ݱ��
                If InStr(strdjbh, Trim$(.Text)) <= 0 Then
                    strdjbh = strdjbh + Trim$(.Text) + ","
                    strdjbh1 = Mid(strdjbh, 2, Len(strdjbh)) + Trim$(.Text) + ","
                End If
            End If
        Next
    End With
    'ȥ�����ݱ�����һ������

    '--------------------------
    If intCount <= 0 Then
        MsgBox "û��ѡ���κ����ϣ�", vbInformation, "��ʾ"
        Exit Function
    End If
    strdjbh = Left$(strdjbh, Len(strdjbh) - 1)
    strdjbh1 = Left$(strdjbh1, Len(strdjbh1) - 1)
    CheckData = True
End Function
'SEDC
Public Sub SEDCExportPrintExcel(ByVal Rs As ADODB.Recordset, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double '�ܽ��
Dim strExtsion          As String '��׺��
Dim strNewFullPath      As String '��Excel�ļ�

    
    If Rs.RecordCount <= 0 Then
        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "��ʼ������..."
    
    strFileName = DirShare & "\SEDC.xls" 'Ҫ�򿪵��ļ�
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
    Rs.MoveFirst    '���ݼ��ƶ�����һ��
    
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        
        lngRows = 3
        IntInertRow = Rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "���ڵ���..."
        
        Do While Not Rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & Rs.fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & Rs.fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & Rs.fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & Rs.fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & Rs.fields(10).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(11).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(12).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & Rs.fields(13).Value)
            wkst.Cells(lngRows, 13) = Trim$("" & Rs.fields(14).Value)
            
            lngRows = lngRows + 1
            Rs.MoveNext
        Loop
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "���赼�����ݣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""����_GB2312,��?""&10������Q ��   "  ' & Gsmc
'        .CenterHeader = "&""����_GB2312,��Ҏ""���&""����,��Ҏ""" & Chr(10) & "&""����_GB2312,��?""&10�� �ڣ�"
'        .RightHeader = "" & Chr(10) & "&""����_GB2312,��Ҏ""&10��λ����ʿ"
'        .LeftFooter = "&""����_GB2312,��Ҏ""&10�Ʊ��ˣ�"
'        .CenterFooter = "&""����_GB2312,��Ҏ""&10�Ʊ����ڣ�"
        .RightFooter = "&20" & "�� &P ҳ���� &N ҳ"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("���ļ��Ѿ����ڣ��Ƿ�Ҫ����ԭ�ļ�?", vbYesNo Or vbQuestion Or vbDefaultButton2, "��ʾ") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "�����ɹ���"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "��ʾ��"
    Exit Sub
End Sub
'Packing list
Public Sub InputPackinglistExportPrintExcel(ByVal Rs As ADODB.Recordset, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double '�ܽ��
Dim strExtsion          As String '��׺��
Dim strNewFullPath      As String '��Excel�ļ�
Dim strXH               As String '��ӡ���

    
    If Rs.RecordCount <= 0 Then
        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "��ʼ������..."
    
    strFileName = DirShare & "\output_packing_list.xls" 'Ҫ�򿪵��ļ�
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
    strXH = Mid$(strExName, InStrRev(strExName, "_"))    '��ȡ���
    Rs.MoveFirst    '���ݼ��ƶ�����һ��
    
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        '��ֵ��Excel�У���ͷ
        wkst.Cells(8, 12) = Format(Date, "YYYY/mm/DD")
        wkst.Cells(9, 12) = "HTKS-SEDC" & Format(Date, "YY-MMDD") & strXH
        
        lngRows = 17
        IntInertRow = Rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "���ڵ���..."
        
        Do While Not Rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & Rs.fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & Rs.fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & Rs.fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & Rs.fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & Rs.fields(10).Value)
            
            'jiayun �޸� ����0��Ϊ��ֵ
            
'            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(11).Value)
'            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(12).Value)
            wkst.Cells(lngRows, 10) = ""
            wkst.Cells(lngRows, 11) = ""
            
            wkst.Cells(lngRows, 12) = Trim$("" & Rs.fields(13).Value)
            wkst.Cells(lngRows, 13) = Trim$("" & Rs.fields(14).Value)
            wkst.Cells(lngRows, 14) = Trim$("" & Rs.fields(15).Value)
            
            'jiayun add Bag#
            wkst.Cells(lngRows, 15) = Trim$("" & Rs.fields(16).Value)
            
            DblNum = DblNum + Val(Trim$("" & Rs.fields(10).Value))
            DblAmt = DblAmt + Val(Trim$("" & Rs.fields(12).Value))
            
            lngRows = lngRows + 1
            Rs.MoveNext
        Loop
        
        wkst.Cells(lngRows, 9) = DblNum
        'wkst.Cells(lngRows, 11) = DblAmt
        
        wkst.Cells(lngRows, 11) = ""
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "���赼�����ݣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""����_GB2312,��?""&10������Q ��   "  ' & Gsmc
'        .CenterHeader = "&""����_GB2312,��Ҏ""���&""����,��Ҏ""" & Chr(10) & "&""����_GB2312,��?""&10�� �ڣ�"
'        .RightHeader = "" & Chr(10) & "&""����_GB2312,��Ҏ""&10��λ����ʿ"
'        .LeftFooter = "&""����_GB2312,��Ҏ""&10�Ʊ��ˣ�"
'        .CenterFooter = "&""����_GB2312,��Ҏ""&10�Ʊ����ڣ�"
        .RightFooter = "&20" & "�� &P ҳ���� &N ҳ"
    End With
  '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("���ļ��Ѿ����ڣ��Ƿ�Ҫ����ԭ�ļ�?", vbYesNo Or vbQuestion Or vbDefaultButton2, "��ʾ") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "�����ɹ���"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "��ʾ��"
    Exit Sub
End Sub
'Invoice
Public Sub InputInvoiceExportPrintExcel(ByVal Rs As ADODB.Recordset, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double '�ܽ��
Dim strExtsion          As String '��׺��
Dim strNewFullPath      As String '��Excel�ļ�
Dim strXH               As String '���
    
    If Rs.RecordCount <= 0 Then
        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "��ʼ������..."
    
    strFileName = DirShare & "\output_invoice.xls" 'Ҫ�򿪵��ļ�
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
    strXH = Mid$(strExName, InStrRev(strExName, "_"))    '��ȡ���
    Rs.MoveFirst    '���ݼ��ƶ�����һ��
    
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        '��ֵ��Excel�У���ͷ
        wkst.Cells(8, 11) = Format(Date, "YYYY/mm/DD")
        wkst.Cells(9, 11) = "HTKS-SEDC" & Format(Date, "YY-MMDD") & strXH
        
        lngRows = 17
        IntInertRow = Rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "���ڵ���..."
        
        Do While Not Rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & Rs.fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & Rs.fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & Rs.fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & Rs.fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & Rs.fields(10).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(11).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(12).Value)
            
            DblNum = DblNum + Val(Trim$("" & Rs.fields(10).Value))
            DblAmt = DblAmt + Val(Trim$("" & Rs.fields(12).Value))
            
            
            lngRows = lngRows + 1
            Rs.MoveNext
        Loop
        
        wkst.Cells(lngRows, 9) = DblNum
        wkst.Cells(lngRows, 11) = DblAmt
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "���赼�����ݣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""����_GB2312,��?""&10������Q ��   "  ' & Gsmc
'        .CenterHeader = "&""����_GB2312,��Ҏ""���&""����,��Ҏ""" & Chr(10) & "&""����_GB2312,��?""&10�� �ڣ�"
'        .RightHeader = "" & Chr(10) & "&""����_GB2312,��Ҏ""&10��λ����ʿ"
'        .LeftFooter = "&""����_GB2312,��Ҏ""&10�Ʊ��ˣ�"
'        .CenterFooter = "&""����_GB2312,��Ҏ""&10�Ʊ����ڣ�"
        .RightFooter = "&20" & "�� &P ҳ���� &N ҳ"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("���ļ��Ѿ����ڣ��Ƿ�Ҫ����ԭ�ļ�?", vbYesNo Or vbQuestion Or vbDefaultButton2, "��ʾ") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "�����ɹ���"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "��ʾ��"
    Exit Sub
End Sub

'Daily_inventory
Public Sub Daily_InvExportPrintExcel(ByVal Rs As ADODB.Recordset, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double '�ܽ��
Dim strExtsion          As String '��׺��
Dim strNewFullPath      As String '��Excel�ļ�

    
    If Rs.RecordCount <= 0 Then
        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "��ʼ������..."
    
    strFileName = DirShare & "\Daily_inventory_report.xls" 'Ҫ�򿪵��ļ�
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
    Rs.MoveFirst    '���ݼ��ƶ�����һ��
    
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        
        lngRows = 3
        IntInertRow = Rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "���ڵ���..."
        
        Do While Not Rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(1).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & Rs.fields(2).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(3).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(4).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(5).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & Rs.fields(6).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & Rs.fields(7).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & Rs.fields(8).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & Rs.fields(9).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(10).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(11).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & Rs.fields(12).Value)
            
            lngRows = lngRows + 1
            Rs.MoveNext
        Loop
        
    Else
'        ClsP.UnLoad_Form
        MsgBox "���赼�����ݣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""����_GB2312,��?""&10������Q ��   "  ' & Gsmc
'        .CenterHeader = "&""����_GB2312,��Ҏ""���&""����,��Ҏ""" & Chr(10) & "&""����_GB2312,��?""&10�� �ڣ�"
'        .RightHeader = "" & Chr(10) & "&""����_GB2312,��Ҏ""&10��λ����ʿ"
'        .LeftFooter = "&""����_GB2312,��Ҏ""&10�Ʊ��ˣ�"
'        .CenterFooter = "&""����_GB2312,��Ҏ""&10�Ʊ����ڣ�"
        .RightFooter = "&20" & "�� &P ҳ���� &N ҳ"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("���ļ��Ѿ����ڣ��Ƿ�Ҫ����ԭ�ļ�?", vbYesNo Or vbQuestion Or vbDefaultButton2, "��ʾ") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "�����ɹ���"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "��ʾ��"
    Exit Sub
End Sub

'shipping Packing list
Public Sub ShippingPackinglistExportPrintExcel(ByVal Ordertemp As String, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double  '�ܽ��
Dim intBoxNum           As Integer '����
Dim strPBigBox          As String  'ǰ���
Dim strNBigBox          As String  '�����
Dim IntBMegerRow        As Integer
Dim IntEMegerRow        As Integer
Dim DblJZ               As Double   '����
Dim DblMZ               As Double   'ë��
Dim DblJZ1               As Double   '����
Dim DblMZ1               As Double   'ë��
Dim DblJZ2               As Double   '����
Dim DblMZ2               As Double   'ë��
Dim intBegin            As Integer
Dim strdjTmp            As String
Dim SD                  As String
Dim SD1                  As String
Dim strTmp()            As String
Dim strExtsion          As String '��׺��
Dim strNewFullPath      As String '��Excel�ļ�
Dim RsNew               As New ADODB.Recordset '��¼����ĸ��������������������
Dim Rs               As New ADODB.Recordset


    strPBigBox = ""
    strNBigBox = ""
    strdjTmp = ""
    intBoxNum = 1
'    If Rs.RecordCount <= 0 Then
'        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
'        Exit Sub
'    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "��ʼ������..."
    
    strFileName = DirShare & "\shipping_packing_list.xlsx" 'Ҫ�򿪵��ļ�
    
    
    strSql = "SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",���,�Ϻ�,mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,����,ë��,�ߴ� " & _
                 " FROM Vw_InvShippedPLFor37 a  where ���ݱ�� in ('" & Ordertemp & "')  order by ���"

     Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
'    '-------------RS�ж���ɸѡ--------------------------------------------
'    If InStr(strdjbh, ",") > 0 Then
'        strTmp = Split(strdjbh, ",")
'        For i = 0 To UBound(strTmp)
'            strdjTmp = strdjTmp + "���ݱ��='" + strTmp(i) + "' OR "
'        Next
'        strdjTmp = Mid$(strdjTmp, 1, Len(strdjTmp) - 5)
'    Else
'        strdjTmp = "���ݱ��='" & strdjbh & "'"
'    End If
'    Rs.Filter = Trim(strdjTmp)          '���ݼ�ɸѡ
'    Rs.Sort = "���ݱ��,��� ASC"       '���ݼ�����
   strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
   strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
'    Rs.MoveFirst    '���ݼ��ƶ�����һ��
'    '---------------------------------------------------------------------
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblJZ = 0
        DblMZ = 0
    
        '��ֵ��Excel�У���ͷ
        wkst.Cells(8, 2) = Trim$("" & Rs.fields(2).Value)
        wkst.Cells(8, 17) = Trim$("" & Rs.fields(3).Value)
        wkst.Cells(17, 2) = Trim$("" & Rs.fields(4).Value)
        wkst.Cells(18, 2) = Trim$("" & Rs.fields(5).Value)
        wkst.Cells(19, 2) = Trim$("" & Rs.fields(6).Value) & " " & Trim$("" & Rs.fields(7).Value)
        wkst.Cells(20, 2) = Trim$("" & Rs.fields(8).Value) & " " & Trim$("" & Rs.fields(9).Value) & " " & Trim$("" & Rs.fields(10).Value) & " " & Trim$("" & Rs.fields(11).Value)
        wkst.Cells(22, 2) = "Attn:" & Trim$("" & Rs.fields(12).Value) & " ,Tel:" & Trim$("" & Rs.fields(13).Value)
        wkst.Cells(23, 2) = ""
        wkst.Cells(17, 17) = Trim$("" & Rs.fields(11).Value) 'To
        wkst.Cells(25, 3) = Trim$("" & Rs.fields(14).Value)
        wkst.Cells(25, 6) = Trim$("" & Rs.fields(15).Value)
        
        lngRows = 28
        
        IntInertRow = Rs.RecordCount * 2
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
           wkst.Rows(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '�߿���
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "���ڵ���..."
        
        IntBMegerRow = 27
        IntEMegerRow = 30
        intBegin = 1
        Dim QBX As String
        For i = 0 To Rs.RecordCount - 1
'            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '���

            strPBigBox = Trim$("" & Rs.fields(16).Value) '���
            'QBX = strPBigBox
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & Rs.fields(16).Value) '���
                '����
                intBoxNum = intBoxNum + 1
                wkst.Cells(lngRows, 1) = "K" & Trim(intBoxNum - 1) '��Ž���ת��Ϊ�ͻ���Ҫ����
                
                IntBMegerRow = IntBMegerRow + intBegin
                intBegin = 1
            Else
'                '�ϲ�
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).Merge
'                '�趨ˮƽ����ֱ����
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).HorizontalAlignment = xlCenter
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).VerticalAlignment = xlCenter
                '--------------------------
                intBegin = intBegin + 1
            End If
            
            If SD <> Trim$("" & Rs.fields(14).Value) Then
            SD = Trim$("" & Rs.fields(14).Value)
            SD1 = SD1 & SD & " "
           End If
          wkst.Cells(25, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & Rs.fields(19).Value)) / 1000 '������Ϊ��ǧΪ��λ
            DblNum = DblNum + Val(Trim$("" & Rs.fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(21).Value) 'datacode
            wkst.Cells(lngRows, 12) = Trim$("" & Rs.fields(22).Value) 'lotno
            If strPBigBox <> QBX Then
            wkst.Cells(lngRows, 14) = Trim$("" & Rs.fields(24).Value) '����
            wkst.Cells(lngRows, 15) = "KG"   '���ص�λ
            wkst.Cells(lngRows, 18) = "KG"   'ë�ص�λ
            wkst.Cells(lngRows, 19) = Trim$("" & Rs.fields(26).Value)   '�ߴ�
            wkst.Cells(lngRows, 17) = Trim$("" & Rs.fields(25).Value)   'ë��
            End If
           
           DblJZ1 = Val(Trim$("" & Rs.fields(24).Value))
           If strPBigBox <> QBX Then
           DblJZ = DblJZ1 + DblJZ
           End If
            DblMZ1 = Val(Trim$("" & Rs.fields(25).Value))
           If strPBigBox <> QBX Then
            DblMZ = DblMZ + DblMZ1
            End If
            '
            
            
            
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(23).Value)
            
            QBX = strPBigBox
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            Rs.MoveNext
        Next
        '�������
        wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '��������Ϊ��ǧΪ��λ
        wkst.Cells(lngRows + 1, 9) = "KPCS" '��λ
        wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)    '����
        wkst.Cells(lngRows + 1, 14) = Format(DblJZ, "0.00") '����
        wkst.Cells(lngRows + 1, 17) = DblMZ 'ë�أ���¼����������жԱ�
    Else
'        ClsP.UnLoad_Form
        MsgBox "���赼�����ݣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
    '��ѯ��ųߴ磬���������
    Dim strXHCC         As String       '�����ͳߴ�
    Dim DblTJZ          As String       '�����
    Dim order As String
    
    order = Replace(Ordertemp, "A", "")
    
    strXHCC = ""
    DblTJZ = 0
    'strdjbh1 = Mid(strdjbh, 2, Len(strdjbh) - 1)
    strSql = "SELECT COUNT(DISTINCT erpdata.dbo.f_getparent(b.���)) ����,c.�ߴ� " & _
             " FROM erpdata..tblStockMove a " & _
             " INNER JOIN erpdata..tblStockMovesub b ON a.���ݱ�� = b.���ݱ�� AND a.���=b.������� " & _
             " INNER JOIN erpdata..tblStockNumTree c On c.���=erpdata.dbo.f_getparent(b.���) " & _
             " WHERE a.���ݱ�� IN ('" & order & "')" & _
             " GROUP BY c.�ߴ�"
    If RsNew.State = adStateOpen Then RsNew.Close
    RsNew.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If Not RsNew.EOF Then
        Do While Not RsNew.EOF
            'ѭ���õ���ŵ������ͳߴ磬����ƴ��
            strXHCC = strXHCC & Trim$("" & RsNew!����) & "@" & Trim$("" & RsNew!�ߴ�) & "cm;"
            '�Գߴ���зָ���������
            If Trim$("" & RsNew!�ߴ�) <> "" And InStr(Trim$("" & RsNew!�ߴ�), "*") > 0 Then
                strTmp = Split(Trim$("" & RsNew!�ߴ�), "*") '�ָ��ַ�
                '��������ز�����
                DblTJZ = DblTJZ + Val(Trim$("" & RsNew!����)) * strTmp(0) * strTmp(1) * strTmp(2) / 5000
            End If
            RsNew.MoveNext
        Loop
    End If
    RsNew.Close
    '��ֵ�����
    wkst.Cells(lngRows + 3, 4) = Format(DblTJZ, "0.00")
    '�Ƚ�����غ�ë�ؿ��ĸ���,��ȡ�ĸ�
    If DblMZ > DblTJZ Then
        wkst.Cells(lngRows + 3, 11) = Format(DblMZ, "0.00")
    Else
        wkst.Cells(lngRows + 3, 11) = Format(DblTJZ, "0.00")
    End If
    '��ֵ��EXCEL�����ͳߴ�
    wkst.Cells(lngRows + 4, 3) = strXHCC
'    ExApp.Columns.AutoFit '����Ӧ�п�
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""����_GB2312,��?""&10������Q ��   "  ' & Gsmc
'        .CenterHeader = "&""����_GB2312,��Ҏ""���&""����,��Ҏ""" & Chr(10) & "&""����_GB2312,��?""&10�� �ڣ�"
'        .RightHeader = "" & Chr(10) & "&""����_GB2312,��Ҏ""&10��λ����ʿ"
'        .LeftFooter = "&""����_GB2312,��Ҏ""&10�Ʊ��ˣ�"
'        .CenterFooter = "&""����_GB2312,��Ҏ""&10�Ʊ����ڣ�"
       ' .RightFooter = "&20" & "�� &P ҳ���� &N ҳ"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("���ļ��Ѿ����ڣ��Ƿ�Ҫ����ԭ�ļ�?", vbYesNo Or vbQuestion Or vbDefaultButton2, "��ʾ") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "�����ɹ���"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "��ʾ��"
    Exit Sub
End Sub

'shipping invoice
Public Sub ShippingInvoiceExportPrintExcel(ByVal Ordertemp As String, ByVal strExName As String)
Dim strSql              As String
Dim lngRows             As Long
Dim rsQuery             As Excel.QueryTable
Dim ExApp               As Excel.Application
Dim wkbk                As New Workbook
Dim wkst                As New Worksheet
Dim i                   As Long
Dim j                   As Long
Dim IntCols             As Integer
Dim strCols             As String
Dim strFileName         As String
Dim IntInertRow         As Integer, IntMaxDetailRow As Integer
Dim DblNum              As Double
Dim DblAmt              As Double  '�ܽ��
Dim intBoxNum           As Integer '����
Dim strPBigBox          As String  'ǰ���
Dim strNBigBox          As String  '�����
Dim IntBMegerRow        As Integer
Dim IntEMegerRow        As Integer
Dim DblJZ               As Double   '����
Dim DblMZ               As Double   'ë��
Dim intBegin            As Integer
Dim strdjTmp            As String
Dim strTmp()            As String
Dim SD                  As String
Dim SD1                  As String
Dim strExtsion          As String '��׺��
Dim strNewFullPath      As String '��Excel�ļ�
Dim Rs               As New ADODB.Recordset
    strPBigBox = ""
    strNBigBox = ""
    strdjTmp = ""
    intBoxNum = 1
'
'    If Rs.RecordCount <= 0 Then
'        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
'        Exit Sub
'    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "��ʼ������..."
    
    strFileName = DirShare & "\shipping_invoice.xlsx" 'Ҫ�򿪵��ļ�
    
    
                    
    strSql = " SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
                 " ,city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 " ,���,�Ϻ�,mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,����,AMount,customerPartNumber " & _
                 "  FROM Vw_InvShippedInvoiceFor37 a  where ���ݱ�� in ('" & Ordertemp & "')  order by ���  "

     Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
'    '-------------RS�ж���ɸѡ--------------------------------------------
'    If InStr(strdjbh, ",") > 0 Then
'        strTmp = Split(strdjbh, ",")
'        For i = 0 To UBound(strTmp)
'            strdjTmp = strdjTmp + "���ݱ��='" + strTmp(i) + "' OR "
'        Next
'        strdjTmp = Mid$(strdjTmp, 1, Len(strdjTmp) - 5)
'    Else
'        strdjTmp = "���ݱ��='" & strdjbh & "'"
'    End If
'    Rs.Filter = Trim(strdjTmp)          '���ݼ�ɸѡ
'    Rs.Sort = "���ݱ��,��� ASC"       '���ݼ�����
   strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
   strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
'    Rs.MoveFirst    '���ݼ��ƶ�����һ��
'    '---------------------------------------------------------------------
    If Rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
    
    
   
    
    
'    '-------------RS�ж���ɸѡ--------------------------------------------
'    If InStr(strdjbh, ",") > 0 Then
'        'strTmp = Split(strdjbh, ",")
'       ' For i = 0 To UBound(strTmp)
'            'strdjTmp = strdjTmp + "���ݱ��='" + strTmp(i) + "' AND "
'       ' Next
'       ' strdjTmp = Mid$(strdjTmp, 1, Len(strdjTmp) - 5)
'        MsgBox "�޷�ͬʱ����������ݺţ�"
'        Exit Sub
'    Else
'        strdjTmp = "���ݱ��='" & strdjbh & "'"
'    End If
'    Rs.Filter = Trim(strdjTmp)  '���ݼ�ɸѡ
'    Rs.Sort = "���ݱ��,��� ASC" '���ݼ�����
'    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
'    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
'    Rs.MoveFirst    '���ݼ��ƶ�����һ��
'    '---------------------------------------------------------------------
'    If Rs.RecordCount > 0 Then
''        ClsP.ShowProgress 30, "��ʼ��Excel..."
'        Set ExApp = New Excel.Application
'        ExApp.Visible = False   '�Ƿ���ʾ
'
'        Set wkbk = ExApp.Workbooks.open(strFileName)
'        Set wkst = wkbk.Worksheets(1)
''        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        DblJZ = 0
        DblMZ = 0
        '��ֵ��Excel�У���ͷ
        wkst.Cells(13, 10) = Trim$("" & Rs.fields(2).Value)
        wkst.Cells(15, 10) = Trim$("" & Rs.fields(3).Value)
        wkst.Cells(18, 10) = Trim$("" & Rs.fields(11).Value) 'To
        
        wkst.Cells(13, 2) = Trim$("" & Rs.fields(4).Value)
        wkst.Cells(14, 2) = Trim$("" & Rs.fields(5).Value)
        wkst.Cells(15, 2) = Trim$("" & Rs.fields(6).Value) & " " & Trim$("" & Rs.fields(7).Value)
        wkst.Cells(16, 2) = Trim$("" & Rs.fields(8).Value) & " " & Trim$("" & Rs.fields(9).Value) & " " & Trim$("" & Rs.fields(10).Value) & " " & Trim$("" & Rs.fields(11).Value)
        
        wkst.Cells(18, 2) = "Attn:" & Trim$("" & Rs.fields(12).Value) & " ,Tel:" & Trim$("" & Rs.fields(13).Value)
        wkst.Cells(19, 2) = ""

        'wkst.Cells(23, 3) = Trim$("" & Rs.fields(14).Value)
        wkst.Cells(23, 5) = Trim$("" & Rs.fields(15).Value)
        
        lngRows = 27
        
        IntInertRow = Rs.RecordCount * 2
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
           wkst.Range(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '�߿���
        Next i
        IntMaxDetailRow = Rs.RecordCount
        
'        ClsP.ShowProgress 50, "���ڵ���..."
        
        IntBMegerRow = 26
        IntEMegerRow = 29
        intBegin = 1
        Dim QBX As String
        
        For i = 0 To Rs.RecordCount - 1
'            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '���
            strPBigBox = Trim$("" & Rs.fields(16).Value) '���
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & Rs.fields(16).Value) '���
                '����
                intBoxNum = intBoxNum + 1
                wkst.Cells(lngRows, 1) = "K" & Trim(intBoxNum - 1) '��Ž���ת��Ϊ�ͻ���Ҫ����
                QBX = "K" & Trim(intBoxNum - 1)
                
                IntBMegerRow = IntBMegerRow + intBegin
                intBegin = 1
            Else
'                '�ϲ�
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).Merge
'                '�趨ˮƽ����ֱ����
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).HorizontalAlignment = xlCenter
'                wkst.Range(Chr(65) & IntBMegerRow & ":" & Chr(66) & IntEMegerRow).VerticalAlignment = xlCenter
                '--------------------------
                intBegin = intBegin + 1
            End If
              If SD <> Trim$("" & Rs.fields(14).Value) Then
             SD = Trim$("" & Rs.fields(14).Value)
             SD1 = SD1 & SD & " "
             End If
            wkst.Cells(23, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & Rs.fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & Rs.fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & Rs.fields(19).Value)) / 1000 '������Ϊ����1000��ֵ
            DblNum = DblNum + Val(Trim$("" & Rs.fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(21).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & Rs.fields(22).Value)
            wkst.Cells(lngRows, 13) = "US$"
            wkst.Cells(lngRows, 14) = Val(Trim$("" & Rs.fields(23).Value)) * 1000 '����Ϊ����1000��ֵ
            wkst.Cells(lngRows, 15) = "US$"
            wkst.Cells(lngRows, 16) = Trim$("" & Rs.fields(24).Value)
            DblAmt = DblAmt + Val(Trim$("" & Rs.fields(24).Value))
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & Rs.fields(25).Value)
            
            
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            Rs.MoveNext
        Next
        
        '�������
        wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '����
        wkst.Cells(lngRows + 1, 9) = "KPCS" '��λ
        wkst.Cells(lngRows + 1, 16) = DblAmt
        wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)

        
    Else
'        ClsP.UnLoad_Form
        MsgBox "���赼�����ݣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
'    ExApp.Columns.AutoFit '����Ӧ�п�
    With wkst.PageSetup
'        .LeftHeader = "" & Chr(10) & "&""����_GB2312,��?""&10������Q ��   "  ' & Gsmc
'        .CenterHeader = "&""����_GB2312,��Ҏ""���&""����,��Ҏ""" & Chr(10) & "&""����_GB2312,��?""&10�� �ڣ�"
'        .RightHeader = "" & Chr(10) & "&""����_GB2312,��Ҏ""&10��λ����ʿ"
'        .LeftFooter = "&""����_GB2312,��Ҏ""&10�Ʊ��ˣ�"
'        .CenterFooter = "&""����_GB2312,��Ҏ""&10�Ʊ����ڣ�"
        '.RightFooter = "&20" & "�� &P ҳ���� &N ҳ"
    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("���ļ��Ѿ����ڣ��Ƿ�Ҫ����ԭ�ļ�?", vbYesNo Or vbQuestion Or vbDefaultButton2, "��ʾ") = vbNo Then
            Exit Sub
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.Number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    End If
    wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.Saved = True
    '---------------------------------------------------------------------------------------------------------------
'    ClsP.ShowProgress 100, "�����ɹ���"
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    ExApp.Visible = True
    
'    If intFlag = 1 Then
'        wkst.PrintPreview
'        wkbk.Close (False)
'        ExApp.Quit
'    End If
    
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
    Exit Sub
ErrHandle:
    On Error Resume Next
    If Not ExApp Is Nothing Then
        Set wkst = Nothing
        Set wkbk = Nothing
        Set ExApp = Nothing
    End If
'    If Not ClsP Is Nothing Then
'        Set ClsP = Nothing
'    End If
    MsgBox Err.Description, vbInformation, "��ʾ��"
    Exit Sub
End Sub

'����Rs���ݼ���䵼��Excel
Public Sub RsExporToExcel(Rs As ADODB.Recordset, RptName As String, ExcelFileName As String)
Dim Irowcount       As Long
Dim Icolcount       As Integer
Dim strFileName     As String

    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlQuery As Excel.QueryTable
    Screen.MousePointer = 11
    With Rs
        If .RecordCount < 1 Then
            Screen.MousePointer = 0
            MsgBox ("û�пɵ���������")
            Exit Sub
        End If
        Irowcount = .RecordCount
        Icolcount = .fields.Count
    End With
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlApp.Visible = False

    Set xlQuery = xlSheet.QueryTables.Add(Rs, xlSheet.Range("a1"))
    
'    With xlQuery
'        .FieldNames = True
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = True
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = True
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .PreserveColumnInfo = True
'    End With
    xlSheet.Name = RptName
    xlQuery.FieldNames = True '�W
    xlQuery.Refresh
    
    With xlSheet
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Name = "����"
        '�r
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = True
        '
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
        '
'        .Range(.Cells(2, 1), .Cells(Irowcount + 1, Icolcount)).Font.Size = 9
    End With
    '����ļ�
    strFileName = DirInvRpt + "\" + ExcelFileName
    xlBook.SaveAs strFileName, xlNormal, "", "", False, False
    xlBook.Saved = True
    
    Screen.MousePointer = 0
'    xlApp.Visible = True
    Set xlSheet = Nothing
    xlBook.Close
    Set xlBook = Nothing
    xlApp.Quit
    Set xlApp = Nothing
    
    
    
    

End Sub
