VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmSemtech_Report 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Semtech�����ѯ"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14445
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
   ScaleHeight     =   9105
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHBDC 
      BackColor       =   &H0000FF00&
      Caption         =   "�ϲ���������"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5160
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdDXDC 
      BackColor       =   &H000080FF&
      Caption         =   "���������"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3360
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Fra 
      ForeColor       =   &H000000FF&
      Height          =   7335
      Index           =   1
      Left            =   3600
      TabIndex        =   17
      Top             =   960
      Width           =   9615
      Begin VB.CheckBox chooseALL 
         Caption         =   "ȫѡ/��ѡ"
         Height          =   315
         Left            =   120
         MaskColor       =   &H0000FF00&
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin FPSpreadADO.fpSpread Fps 
         Height          =   3375
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   6615
         _Version        =   524288
         _ExtentX        =   11668
         _ExtentY        =   5953
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
      Height          =   13815
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "����"
         Height          =   360
         Left            =   2640
         TabIndex        =   27
         Top             =   3720
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         Height          =   360
         Left            =   2640
         TabIndex        =   26
         Top             =   3240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtLotID 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "����"
         Height          =   285
         Left            =   2640
         TabIndex        =   24
         Top             =   2880
         Width           =   735
      End
      Begin VB.ListBox lstLotID 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   10320
         Left            =   840
         Style           =   1  'Checkbox
         TabIndex        =   23
         Top             =   3240
         Width           =   1695
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
         Format          =   109576195
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
         Format          =   109576195
         CurrentDate     =   41387
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D N"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   22
         Top             =   3240
         Width           =   360
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
         Left            =   9240
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
         Enabled         =   0   'False
         Height          =   360
         Left            =   6720
         MaskColor       =   &H0000FFFF&
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�� ��"
         Height          =   360
         Left            =   8040
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdExprot 
         Caption         =   "������ǰ����"
         Enabled         =   0   'False
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
Dim DirQrShare As String

Dim DirInvRpt       As String
Dim order           As String
Dim RsClone         As New ADODB.Recordset
Const C_Left = 60
Const C_Top = 120

Private Enum fpSDetail
    E_CHOOSE = 1
    e_DJBH = 2
    E_cust = 3
    e_YDH = 7
End Enum
'��������ֵ
Private Function GetExcelName(ByVal strTitle As String) As String
Dim strSql          As String
Dim rs              As New ADODB.Recordset
Dim strExFileName   As String
Dim strCurDate      As String
    
    If strTitle = "output Invoice" Then
        strCurDate = Format(DATE, "YY-MMDD")
    Else
        strCurDate = Format(DATE, "YYYYMMDD")
    End If
    strSql = "select nvl(max(para_3),0)+1 para from tblsys_parameter where sysname='TSVSYS' and kind='Semtech����' and para_1='" & strTitle & "' and para_2='" & strCurDate & "'"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        strExFileName = strCurDate + "_" + Trim$("" & rs!Para)
        strSql = "Update tblsys_parameter set para_2='" & strCurDate & "',para_3='" & Trim$("" & rs!Para) & "' where sysname='TSVSYS' and kind='Semtech����' and para_1='" & strTitle & "'"
        Cnn.Execute strSql
    Else
        strExFileName = strCurDate + "_1"
        strSql = "Update tblsys_parameter set para_2='" & strCurDate & "',para_3='" & Trim$("" & rs!Para) & "' where sysname='TSVSYS' and kind='Semtech����' and para_1='" & strTitle & "'"
        Cnn.Execute strSql
    End If
    rs.Close
    
    If strTitle = "output Invoice" Then
        GetExcelName = "HTKS_SEDC" & strExFileName
    Else
        GetExcelName = strTitle & "_" & strExFileName
    End If
    
End Function
'ȫѡ�ͷ���ȫѡ
Private Sub chooseALL_Click()

Dim i As Integer

If chooseALL.Value = 1 Then

    For i = 1 To Fps(0).MaxRows

        With Fps(0)
            .Row = i
            .Col = 1
            .text = 1

        End With

    Next i

ElseIf chooseALL.Value = 0 Then

    For i = 1 To Fps(0).MaxRows

        With Fps(0)
            .Row = i
            .Col = 1
            .text = 0

        End With

    Next i

End If
End Sub

Private Sub cmbCombo1_Click(Index As Integer)
'Dim strSql              As String
'Dim Rs                  As New ADODB.Recordset


Fps(0).MaxRows = 0

    If Index = 0 Then
        If cmbCombo1(0).text = "��汨��" Then
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


'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       cmdDXDC_Click
' Description:       DN��ַ�ϲ���������
' Created by :       Project Administrator
' Machine    :       DESKTOP-F6L8S2V
' Date-Time  :       2019/10/30-14:38:24
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdHBDC_Click()
    Dim strExportName           As String
    
    cmdReport.Enabled = False
    If Fps(0).MaxRows <= 0 Then
        MsgBox "û�пɵ��������ݣ�", vbInformation, "��ʾ"
        Exit Sub
    End If
    '��������
    strExportName = GetExcelName(Trim(Fra(1).Caption))
    
    If cmbCombo1(0).ListIndex = 0 Then
'        Call SEDCExportPrintExcel(RsClone, strExportName)                     'SEDC����
    ElseIf cmbCombo1(0).ListIndex = 1 Then
'        Call InputPackinglistExportPrintExcel(RsClone, strExportName)         'outputPackinglist����
    ElseIf cmbCombo1(0).ListIndex = 2 Then
'        Call InputInvoiceExportPrintExcel(RsClone, strExportName)             'outputInvoice����
    ElseIf cmbCombo1(0).ListIndex = 3 Then
'        Call Daily_InvExportPrintExcel(RsClone, strExportName)                'Daily_inventory_report
    ElseIf cmbCombo1(0).ListIndex = 4 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel2(strExportName, 0)      'Shipping Packinglist
    ElseIf cmbCombo1(0).ListIndex = 5 Then
        If Not CheckData Then Exit Sub
        Call ShippingInvoiceExportPrintExcel2            'Shipping invoice
    ElseIf cmbCombo1(0).ListIndex = 9 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel2(strExportName, 1)
    End If
    
    cmdReport.Enabled = True
    
End Sub


'--------------------------------------------------------------------------------
' Project    :       ��ʽ����1
' Procedure  :       cmdZXDC_Click
' Description:       ���������
' Created by :       ף�t��
' Machine    :       DESKTOP-F6L8S2V
' Date-Time  :       2019/10/30-14:37:45
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdDXDC_Click()
    Dim strExportName           As String

    cmdReport.Enabled = False
    If Fps(0).MaxRows <= 0 Then
        MsgBox "û�пɵ��������ݣ�", vbInformation, "��ʾ"
        Exit Sub
    End If
    '��������
    strExportName = GetExcelName(Trim(Fra(1).Caption))
    
    If cmbCombo1(0).ListIndex = 0 Then
'        Call SEDCExportPrintExcel(RsClone, strExportName)                     'SEDC����
    ElseIf cmbCombo1(0).ListIndex = 1 Then
'        Call InputPackinglistExportPrintExcel(RsClone, strExportName)         'outputPackinglist����
    ElseIf cmbCombo1(0).ListIndex = 2 Then
'        Call InputInvoiceExportPrintExcel(RsClone, strExportName)             'outputInvoice����
    ElseIf cmbCombo1(0).ListIndex = 3 Then
'        Call Daily_InvExportPrintExcel(RsClone, strExportName)                'Daily_inventory_report
    ElseIf cmbCombo1(0).ListIndex = 4 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel1(strExportName, 0)    'Shipping Packinglist
    ElseIf cmbCombo1(0).ListIndex = 5 Then
        If Not CheckData Then Exit Sub
'        Call ShippingInvoiceExportPrintExcel(order, strExportName)          'Shipping invoice
    ElseIf cmbCombo1(0).ListIndex = 9 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel1(strExportName, 1)
    End If
    
    cmdReport.Enabled = True
    
'Shell (App.Path & "\install.bat")
End Sub

Private Sub cmdExit_Click() '�˳�
    Unload Me
End Sub

Private Sub cmdExprot_Click()
Dim strExportName           As String

cmdExprot.Enabled = False
    'У������
    If Fps(0).MaxRows <= 0 Then
        MsgBox "û�пɵ��������ݣ�", vbInformation, "��ʾ"
        Exit Sub
    End If
    '��������
    If cmbCombo1(0).ListIndex = 6 Or cmbCombo1(0).ListIndex = 7 Or cmbCombo1(0).ListIndex = 8 Then '��汨��,SMTCList,Shipped
        strExportName = Trim(Fra(1).Caption)
        If cmbCombo1(1).text <> "" Then
            strExportName = Trim(Fra(1).Caption) + "-" + Trim(cmbCombo1(1).text)
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
Dim rs                      As New ADODB.Recordset
Dim i                       As Integer
Dim strFileName             As String
Dim strmsg                  As String
    
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
            If rs.State = adStateOpen Then rs.Close
            rs.Open strSql + strSqlDetail, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs.EOF Then '��ʾ�����ݲŵ�������
                strFileName = Format(Now(), "YYMMDD_HHMM") + "_" + Replace(cmbCombo1(1).List(i), "Scrap", "SCR")
                strmsg = strmsg + DirInvRpt + "\" + strFileName + vbCrLf                '��ʾ��Ϣ
                RsExporToExcel rs, cmbCombo1(1).List(i), strFileName                    '������Excel
            End If
            rs.Close
        End If
    Next
    '����Shipped
    strSql = "SELECT RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Remark,SO as [DN#],SHIPPED_DATE,Cust_Name " & _
             " FROM Vw_InvShippedRptFor37 "
'             " WHERE SHIPPED_DATE>='" & DateAdd("m", -1, Format(Now(), "YYYY-MM-DD")) & "' and SHIPPED_DATE<'" & DateAdd("d", 1, Format(Now(), "YYYY-MM-DD")) & "' "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then '��ʾ�����ݲŵ�������
        strFileName = Format(Now(), "YYMMDD_HHMM") + "_SHIPPED"
        strmsg = strmsg + DirInvRpt + "\" + strFileName + vbCrLf                '��ʾ��Ϣ
        RsExporToExcel rs, "SHIPPED", strFileName                               '������Excel
    End If
    rs.Close
    
    MsgBox "�����ɹ��������ļ�·��Ϊ��" + vbCrLf + strmsg
    
End Sub

Private Sub cmdquery_Click()
Dim strKey As String
Dim i      As Integer
Dim bRet   As Boolean

bRet = False
strKey = Trim$(txtLotID.text)
If strKey = "" Then
    MsgBox "������DN", vbInformation, "��ʾ:"
    Exit Sub

End If

With lstLotID

    For i = 0 To .ListCount - 1
        If strKey = .List(i) Then
            .Selected(i) = True
            bRet = True

        End If

    Next

End With

If bRet = False Then
    MsgBox "��ѯ������DN", vbInformation, "��ʾ"

End If




End Sub

Private Sub cmdReport_Click() '��������
    Dim strExportName           As String

    cmdReport.Enabled = False
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
        Call ShippingPackinglistExportPrintExcel(order, strExportName, 0)    'Shipping Packinglist
    ElseIf cmbCombo1(0).ListIndex = 5 Then
        If Not CheckData Then Exit Sub
        Call ShippingInvoiceExportPrintExcel(order, strExportName)          'Shipping invoice
    ElseIf cmbCombo1(0).ListIndex = 9 Then
        If Not CheckData Then Exit Sub
        Call ShippingPackinglistExportPrintExcel(order, strExportName, 1)
    End If
   
    
cmdReport.Enabled = True
End Sub

Private Sub cmdSearch_Click() '��ѯ����
Dim i                   As Long
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
   
        Dim strDNList As String
    '��ʼ��FPS
    
                With lstLotID

            For i = 0 To .ListCount - 1
        
                If .Selected(i) = True Then
                    strDNList = strDNList & Trim$("" & .List(i)) & "','"
    

                End If

            Next

        End With

        
          strDNList = Mid(strDNList, 1, Len(strDNList) - 3)
    
    order = ""
    InitFps
    cmdReport.Enabled = True
    cmdExprot.Enabled = True
    cmdDXDC.Enabled = True
    cmdHBDC.Enabled = True
    
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
                
        If txt(1).text <> "" Then
            strSql = strSql & " And wo.SOURCE_BATCH_ID='" & Trim(txt(1).text) & "'"
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
                
        If txt(1).text <> "" Then
            strSql = strSql & " And Job_No='" & Trim(txt(1).text) & "'"
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
                
        If txt(1).text <> "" Then
            strSql = strSql & " And job_no='" & Trim(txt(1).text) & "'"
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
                
        If txt(1).text <> "" Then
            strSql = strSql & " And job_no='" & Trim(txt(1).text) & "'"
        End If
        strSql = strSql & " group by ѡ��,TESTDC,LOCATION,DEVICE_NAME,job_no,HTLOT_NO,date_code,CCOMMENT,Reel_Size,Remark,Move_in_Date"
        
  ElseIf cmbCombo1(0).ListIndex = 4 Then  'shipping packing list (INNER JOIN tblCustomerOI d ON CASE WHEN charindex('M',RTRIM(c.batchnumber))>0 THEN LEFT(RTRIM(c.batchnumber),LEN(RTRIM(c.batchnumber))-1) ELSE RTRIM(c.batchnumber) END=d.SOURCE_BATCH_ID)
        strSql = "SELECT X.* FROM ( SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",���,�Ϻ�,replace(mpn_desc,'.P2','') as mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,����,ë��,�ߴ� " & _
                 " FROM Vw_InvShippedPLFor37_NEW a " & _
                 " WHERE ��������>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and ��������<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "' " & _
                 " union all " & _
                 "SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",���,�Ϻ�,replace(mpn_desc,'.P2','') as mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,����,ë��,�ߴ� " & _
                 " FROM Vw_InvShippedPLFor37 a " & _
                 " WHERE ��������>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and ��������<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "'  ) X "
                 
        If txt(0).text <> "" Then
            strSql = strSql & " And ���ݱ��='" & Trim(txt(0).text) & "'"
        End If
        If txt(1).text <> "" Then
            strSql = strSql & " And batchnumber='" & Trim(txt(1).text) & "'"
        End If
 
     
'            strSql = "SELECT 0 ѡ��,h.������ + a.���ݱ�� ���ݱ��,c.delivery,dbo.usp_date(a.��������) ��������,ISNULL(dn_address_new, c.shiptoname) AS shiptoname,ISNULL(x.ship_to_street1_new, c.shiptostreet1) AS shiptostreet1, " & _
'"ISNULL(x.ship_to_street2_new, c.shiptostreet2) AS shiptostreet2,ISNULL(x.ship_to_street3_new, c.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, c.city) AS city, ISNULL(x.dn_st_new, c.State) AS state, " & _
'"ISNULL(x.postal_code_new, c.postalcode) AS postalcode,ISNULL(x.country_new, c.countrykey) AS countrykey,ISNULL(x.contact_new, c.contactname) AS contactname,ISNULL(x.phone_new, c.phone) AS phone, " & _
'"c.SalesDocument,'''' + c.PurchasingDocNo AS PurchasingDocNo,RTRIM(erpdata.dbo.f_getparent(b.���)) ���,b.�Ϻ�, " & _
'"CASE WHEN RTRIM(gg.MPN_DESC) = 'UCLAMP0541Z.TFT' THEN RTRIM(gg.MPN_DESC) + '.P2' ELSE REPLACE(REPLACE(gg.MPN_DESC, '.P2', ''), '.P3', '') END AS mpn_desc,SUM(b.����) ����,c.batchnumber, " & _
'"hh.CREATE_DATE DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no,c.customerPartNumber,ROUND(CAST(f.���� AS FLOAT) * 0.4, 2) ����,f.���� ë��,f.�ߴ� FROM erpdata .. tblStockSQfh a " & _
'"INNER JOIN erpdata .. tblStocksqfhsub b ON a.���ݱ�� = b.���ݱ�� AND a.��� = b.������� INNER JOIN erpdata .. tblStockNumTree g ON g.��� = b.��� INNER JOIN erpdata .. tblStockNumTree f " & _
'"ON f.��� = g.�ϼ���� INNER JOIN (SELECT a.BOX_ID,SUBSTRING(a.KEY_VALUE,1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox,SUBSTRING(a.KEY_VALUE,CHARINDEX('|', a.KEY_VALUE) + 1,10) AS job " & _
'"FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON g.��� = aa.qbox INNER JOIN (SELECT dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, " & _
'"dn.shiptostreet3, dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber " & _
'"FROM tblCustomerShippingUp dn GROUP BY dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname, " & _
'"dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,dn.customerPartNumber) c ON c.Delivery = g.DN AND c.BatchNumber = aa.job INNER JOIN dbo.tblstock h " & _
'"ON CONVERT(NVARCHAR(4), h.�ⷿ����) = CONVERT(NVARCHAR(4), a.�ֿ���) INNER JOIN ERPBASE .. tblmappingData ff ON ff.SUBSTRATEID = b.���̿���� " & _
'"INNER JOIN ERPBASE .. tblCustomerOI gg ON CONVERT(VARCHAR(100), gg.ID) = ff.FILENAME AND gg.SOURCE_BATCH_ID = ff.LOTID AND gg.CUSTOMERSHORTNAME = '37' INNER JOIN erpbase .. weight37 hh " & _
'"ON hh.WAFERID = REPLACE(b.���̿����, '+', '') INNER JOIN erpdata .. tblErpInStockMainData jj ON jj.BOX_ID = aa.BOX_ID LEFT JOIN erptemp .. dn_address x ON dn_address = c.ShipToName " & _
'"WHERE a.�ͻ����� = '37' AND a.���ݱ�� LIKE 'F%' AND a.�������� >= CONVERT(VARCHAR(100), GETDATE() - 5, 23) AND c.Delivery = g.DN " & _
'"and c.Delivery in ('" & strDNList & "') GROUP BY h.������, a.���ݱ��, c.delivery,dbo.usp_date(a.��������), c.shiptoname,c.shiptostreet1,c.shiptostreet2,c.shiptostreet3,c.city,c.State,c.postalcode, " & _
'"c.countrykey,c.contactname,c.phone,c.SalesDocument,c.PurchasingDocNo,erpdata.dbo.f_getparent(b.���),b.�Ϻ�,gg.MPN_DESC,c.batchnumber,hh.CREATE_DATE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2), " & _
'"c.customerPartNumber,f.����,f.�ߴ�,dn_address_new,x.ship_to_street1_new,x.ship_to_street2_new,x.ship_to_street3_new,x.city_new,x.dn_st_new,x.postal_code_new,x.country_new,x.contact_new,x.phone_new"

            
    strSql = "SELECT 0 AS  ѡ��,y.������ + b.���ݱ�� AS ���ݱ��,d.Delivery, CONVERT(VARCHAR(100), c.��������,23) AS ��������, ISNULL(dn_address_new, d.shiptoname) AS shiptoname, ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1 " & _
 " ,  ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2, ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,  ISNULL(x.city_new, d.city) AS city,  ISNULL(x.dn_st_new, d.State) AS state,  ISNULL(x.postal_code_new, d.postalcode) AS postalcode, " & _
 "  ISNULL(x.country_new, d.countrykey) AS countrykey,  ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone, d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo " & _
 " ,f.��� ,b.�Ϻ� ,d.MarketingPN,SUM(b.����),d.BatchNumber,d.DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no, d.customerPartNumber, ROUND(CAST(f.���� AS FLOAT) * 0.4, 2) ����, f.���� ë��, " & _
 " f.�ߴ� FROM erpdata..tblStockNumTree a INNER JOIN erpdata..tblStocksqfhsub b ON b.��� = a.��� INNER JOIN erpdata..tblStockSQfh c ON c.���ݱ�� = b.���ݱ�� AND c.��� = b.������� INNER JOIN erpdata..tblstock y " & _
 " ON y.�ⷿ���� = c.�ֿ���  INNER JOIN (SELECT a.BOX_ID,  SUBSTRING(a.KEY_VALUE,  1, CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job " & _
 " FROM erpdata .. tblErpInStockDetailInfo a  WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.��� = aa.qbox  INNER JOIN (SELECT dn.Delivery, dn.shiptoname, dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3 " & _
 " , dn.city, dn.State, dn.postalcode, dn.countrykey, dn.contactname,  dn.phone, dn.SalesDocument,  dn.PurchasingDocNo, dn.BatchNumber, SUM(dn.Quantity) AS Quantity, dn.customerPartNumber, dn.MarketingPN, dn.DATE_CODE " & _
 " FROM ERPBASE..tblCustomerShippingUp dn WHERE dn.Delivery IN ('" & strDNList & "') GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1, dn.shiptostreet2, dn.shiptostreet3,dn.city, dn.State, dn.postalcode, dn.countrykey " & _
 " , dn.contactname,dn.phone, dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber, dn.customerPartNumber, dn.MarketingPN,dn.DATE_CODE) d  ON d.Delivery = a.DN AND d.BatchNumber = aa.job " & _
 " INNER JOIN erpdata..tblStockNumTree e ON e.��� = b.��� INNER JOIN erpdata..tblstocknumtree f  ON f.��� = e.�ϼ���� LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName " & _
"  WHERE a.DN IN ('" & strDNList & "')GROUP BY  b.���ݱ��,c.��������,ISNULL(dn_address_new, d.shiptoname) , ISNULL(x.ship_to_street1_new, d.shiptostreet1) ,ISNULL(x.ship_to_street2_new, d.shiptostreet2) " & _
 " ,ISNULL(x.ship_to_street3_new, d.shiptostreet3),ISNULL(x.city_new, d.city),ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode), ISNULL(x.country_new, d.countrykey) " & _
 " ,ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) , d.SalesDocument,d.PurchasingDocNo ,f.���,b.�Ϻ� ,d.MarketingPN,d.BatchNumber ,d.DATE_CODE " & _
 " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber, f.���� , f.�ߴ�,y.������,d.Delivery order by shiptoname, Delivery,���"



        
    ElseIf cmbCombo1(0).ListIndex = 5 Then  'shipping invoice (INNER JOIN tblCustomerOI d ON CASE WHEN charindex('M',RTRIM(c.batchnumber))>0 THEN LEFT(RTRIM(c.batchnumber),LEN(RTRIM(c.batchnumber))-1) ELSE RTRIM(c.batchnumber) END=d.SOURCE_BATCH_ID)
        strSql = "select x.* from ( SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
                 ",city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",���,�Ϻ�,replace(mpn_desc,'.P2','') as mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,����,AMount,customerPartNumber " & _
                 " FROM Vw_InvShippedInvoiceFor37_NEW a " & _
                 " WHERE ��������>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and ��������<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "' " & _
                 "union all " & _
                 " SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
                 ",city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
                 ",���,�Ϻ�,replace(mpn_desc,'.P2','') as mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,����,AMount,customerPartNumber " & _
                 " FROM Vw_InvShippedInvoiceFor37 a " & _
                 " WHERE ��������>='" & Format(DTP(0).Value, "YYYY-MM-DD") & "' and ��������<'" & Format(DTP(1).Value + 1, "YYYY-MM-DD") & "') x "
                 
        
        
                If txt(0).text <> "" Then
            strSql = strSql & " And ���ݱ��='" & Trim(txt(0).text) & "'"
        End If
        If txt(1).text <> "" Then
            strSql = strSql & " And batchnumber='" & Trim(txt(1).text) & "'"
        End If
        
        
        
        
        
  
        
        
'
'        strSql = "SELECT 0 ѡ��, h.������+a.���ݱ�� ���ݱ��,c.delivery,dbo.usp_date(a.��������) ��������, ISNULL(dn_address_new, c.shiptoname) AS  shiptoname,ISNULL(x.ship_to_street1_new ,c.shiptostreet1)  AS shiptostreet1  " & _
'",ISNULL(x.ship_to_street2_new,c.shiptostreet2) AS shiptostreet2 ,ISNULL(x.ship_to_street3_new,c.shiptostreet3) AS shiptostreet3 ,ISNULL(x.city_new ,c.city) AS city,ISNULL(x.dn_st_new, c.State) AS state,ISNULL(x.postal_code_new,c.postalcode  ) AS  postalcode " & _
'",ISNULL(x.country_new,c.countrykey) AS countrykey,ISNULL(x.contact_new,c.contactname) AS contactname ,ISNULL(x.phone_new,c.phone ) AS phone      " & _
'",c.SalesDocument,'''' + c.PurchasingDocNo AS PurchasingDocNo,RTRIM(erpdata.dbo.f_getparent(b.���)) ���,b.�Ϻ�,CASE WHEN RTRIM(gg.MPN_DESC)='UCLAMP0541Z.TFT' THEN RTRIM(gg.MPN_DESC)+'.P2' " & _
'"ELSE REPLACE(REPLACE(gg.MPN_DESC,'.P2',''),'.P3','') END AS mpn_desc,SUM(b.����) ����,c.batchnumber,hh.CREATE_DATE DATE_CODE " & _
'",SUBSTRING(aa.qbox ,2,CHARINDEX('-R',aa.qbox)-2)  HTlot_no ,ISNULL(ISNULL( BB.��˰���� / AB.��Ʒ��,0)  + ( cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0) AS ���� " & _
'",ROUND( SUM(b.����) * ISNULL(ISNULL( BB.��˰���� / AB.��Ʒ��,0)  +  (cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0),2) AS AMount,c.customerPartNumber  ,e.���۵����  " & _
'",ROUND( SUM(b.����) * ISNULL(ISNULL( BB.��˰���� / AB.��Ʒ��,0)  +  (cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0),2) -  CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( BB.��˰���� / AB.��Ʒ��,0)) AS �ӹ��ѽ�� " & _
'", CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( BB.��˰���� / AB.��Ʒ��,0)) AS �͹��Ͻ�� FROM erpdata..tblStockSQfh a           " & _
'"INNER JOIN erpdata..tblStocksqfhsub b ON a.���ݱ�� = b.���ݱ�� AND a.���=b.������� INNER JOIN erpdata..tblStockNumTree g ON g.���=b.��� " & _
'"INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE,1,CHARINDEX('|',a.KEY_VALUE)-1) AS qbox , SUBSTRING(a.KEY_VALUE,CHARINDEX('|',a.KEY_VALUE)+1,10) AS job " & _
'" FROM erpdata..tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND  a.KEY_VALUE LIKE '%SS%|%')  aa ON g.��� = aa.qbox  " & _
'"INNER JOIN (SELECT dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode "
'
'        strSql = strSql & ",dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber FROM tblCustomerShippingUp dn GROUP BY dn.Delivery,dn.shiptoname,dn.shiptostreet1 " & _
'",dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,dn.customerPartNumber)  c ON c.Delivery = g.DN  AND c.BatchNumber =  aa.job " & _
'"INNER JOIN dbo.tblstock h ON CONVERT(NVARCHAR(4),h.�ⷿ����) = CONVERT(NVARCHAR(4),a.�ֿ���)    INNER JOIN ERPBASE..tblmappingData ff ON ff.SUBSTRATEID = b.���̿���� " & _
'"INNER JOIN ERPBASE..tblCustomerOI gg ON CONVERT(VARCHAR(30), gg.ID) = ff.FILENAME AND gg.SOURCE_BATCH_ID = ff.LOTID AND gg.CUSTOMERSHORTNAME = '37' " & _
'"INNER JOIN erpbase..weight37 hh ON hh.WAFERID = REPLACE(b.���̿����,'+','') INNER JOIN erpdata..tblErpInStockMainData jj ON jj.BOX_ID = aa.BOX_ID " & _
'"LEFT JOIN erpbase..tbltoinrec_wafer  AB ON ab.���� = ff.LOTID AND AB.��ԲID = REPLACE(B.���̿����,'+','') LEFT JOIN erpbase..tbltorec_wafer  ww ON  ww.���� = ab.���� AND  ww.��ԲID = ab.��ԲID  " & _
'"LEFT JOIN ERPBASE..TblToInsub BB ON BB.��ⵥ��� = AB.��ⵥ��� AND BB.�������� = AB.���� AND ww.��������� = bb.��������� AND bb.��˰���� IS NOT NULL " & _
'"LEFT JOIN erptemp..tblBB_CSRPO cb ON cb.PO_NUM = gg.PO_NUM AND cb.FAB_DEVICE = gg.MPN_DESC LEFT JOIN ERPBASE..tblmappingData db ON db.SUBSTRATEID = REPLACE(B.���̿����,'+','')  " & _
'"LEFT JOIN  erpdata..tblSalerec e ON  e.���ݱ�� = a.���ݱ��       AND a.��� = e.�������  AND e.С��� = b.���    " & _
'"LEFT JOIN erptemp..dn_address x ON dn_address = c.ShipToName WHERE a.�ͻ�����='37' and c.Delivery in ('" & strDNList & "') AND a.�������� >= CONVERT(VARCHAR(100),GETDATE()- 8,23) AND  a.���ݱ�� LIKE 'F%' AND a.��Ʒ���� >0  " & _
'"GROUP BY gg.PO_NUM,h.������,a.���ݱ��,c.delivery,dbo.usp_date(a.��������),c.shiptoname,c.shiptostreet1,c.shiptostreet2         " & _
'",c.shiptostreet3,c.city,c.State,c.postalcode,c.countrykey,c.contactname,c.phone,c.SalesDocument,c.PurchasingDocNo ,erpdata.dbo.f_getparent(b.���),b.�Ϻ�,gg.MPN_DESC,c.batchnumber,hh.CREATE_DATE      " & _
'",SUBSTRING(aa.qbox ,2,CHARINDEX('-R',aa.qbox)-2),e.���۵����,c.customerPartNumber ,ISNULL( BB.��˰���� / AB.��Ʒ��,0) ,  cb.WAFER_PRICE,db.PASSBINCOUNT , cb.DIE_PRICE,dn_address_new " & _
'",x.ship_to_street1_new,x.ship_to_street2_new,x.ship_to_street3_new,x.city_new,x.dn_st_new,x.postal_code_new,x.country_new,x.contact_new ,x.phone_new "

'      strsql = " SELECT 0 AS ѡ��,y.������ + b.���ݱ�� AS ���ݱ��,a.DN, CONVERT(VARCHAR(100), c.��������,23) AS ��������, ISNULL(dn_address_new, d.shiptoname) AS shiptoname " & _
' ",ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1,ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2 , ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, d.city) AS city, " & _
' " ISNULL(x.dn_st_new, d.State) AS state, ISNULL(x.postal_code_new, d.postalcode) AS postalcode,ISNULL(x.country_new, d.countrykey) AS countrykey, ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone,d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo, f.���, " & _
' " b.�Ϻ�,d.MarketingPN,SUM(b.����), d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no ,ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  + ( dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0) AS ���� " & _
' " ,ROUND( SUM(b.����) * ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  +  (dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0),2) AS AMount ,d.customerPartNumber ,ROUND( SUM(b.����) * ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  +  (dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0),2) " & _
' "  -  CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( cd.��˰���� / cc.��Ʒ��,0)) AS �ӹ��ѽ�� , CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( cd.��˰���� / cc.��Ʒ��,0)) AS �͹��Ͻ�� " & _
'"   FROM erpdata .. tblStockNumTree a INNER JOIN erpdata .. tblStocksqfhsub b ON b.��� = a.��� INNER JOIN erpdata .. tblStockSQfh c  ON c.���ݱ�� = b.���ݱ�� AND c.��� = b.������� INNER JOIN erpdata..tblstock y ON y.�ⷿ���� = c.�ֿ��� " & _
'"  INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE, 1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job  FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.��� = aa.qbox  INNER JOIN (SELECT dn.Delivery, " & _
' "  dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State,dn.postalcode, dn.countrykey, dn.contactname,dn.phone, dn.SalesDocument,  dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE  FROM ERPBASE .. tblCustomerShippingUp dn " & _
' "   WHERE dn.Delivery IN ('" & strDNList & "')  GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State, dn.postalcode,  dn.countrykey,dn.contactname,dn.phone, dn.SalesDocument, dn.PurchasingDocNo, dn.BatchNumber, dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE) d " & _
' "   ON d.Delivery = a.DN  AND d.BatchNumber = aa.job INNER JOIN erpdata .. tblStockNumTree e  ON e.��� = b.��� INNER JOIN erpdata .. tblstocknumtree f  ON f.��� = e.�ϼ����  INNER JOIN ERPBASE..tblmappingData qq ON qq.SUBSTRATEID = b.���̿���� AND qq.LOTID = b.������ LEFT JOIN ERPBASE..tblCustomerOI bb " & _
' " ON CONVERT(VARCHAR(100), bb.ID) = qq.FILENAME AND bb.SOURCE_BATCH_ID = qq.LOTID  LEFT JOIN ERPBASE..tblToRec_Wafer cc ON cc.��ԲID = REPLACE(qq.SUBSTRATEID,'+','')   AND cc.���� = qq.LOTID  LEFT JOIN ERPBASE..tblToRecEntry cd ON cd.��������� = cc.��������� AND cd.�������� = cc.���� LEFT JOIN erptemp..tblBB_CSRPO dd " & _
'"  ON dd.PO_NUM = bb.PO_NUM AND dd.FAB_DEVICE = bb.MPN_DESC LEFT JOIN  erpdata..tblSalerec j ON j.���ݱ�� = c.���ݱ�� AND j.������� = b.������� AND j.С��� = b.��� LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName  WHERE a.DN IN ('" & strDNList & "')  GROUP BY b.���ݱ��, c.��������, ISNULL(dn_address_new, d.shiptoname), " & _
'"  ISNULL(x.ship_to_street1_new, d.shiptostreet1), ISNULL(x.ship_to_street2_new, d.shiptostreet2),ISNULL(x.ship_to_street3_new, d.shiptostreet3), ISNULL(x.city_new, d.city),  ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode),ISNULL(x.country_new, d.countrykey),ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone), " & _
'" d.SalesDocument, d.PurchasingDocNo,f.���, b.�Ϻ�,d.MarketingPN,d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber,   a.DN , J.���۵����, bb.PO_NUM, cd.��˰����, cc.��Ʒ��, dd.WAFER_PRICE, qq.PASSBINCOUNT, dd.die_price, y.������  order by shiptoname,DN,��� "
'
'
          
   
strSql = "SELECT 0 AS ѡ��,y.������ + b.���ݱ�� AS ���ݱ��,a.DN, CONVERT(VARCHAR(100), c.��������,23) AS ��������, ISNULL(dn_address_new, d.shiptoname) AS shiptoname ,ISNULL(x.ship_to_street1_new " & _
         " , d.shiptostreet1) AS shiptostreet1,ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2 , ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, d.city) AS city " & _
         " ,  ISNULL(x.dn_st_new, d.State) AS state, ISNULL(x.postal_code_new, d.postalcode) AS postalcode,ISNULL(x.country_new, d.countrykey) AS countrykey, ISNULL(x.contact_new, d.contactname) AS contactname " & _
         " , ISNULL(x.phone_new, d.phone) AS phone,d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo, f.���,  b.�Ϻ�,d.MarketingPN,SUM(b.����), d.BatchNumber, d.DATE_CODE " & _
         " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no ,ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  + ( dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0) AS ���� ,ROUND( SUM(b.����) * ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0) " & _
         "  +  (dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0),2) AS AMount ,d.customerPartNumber,ROUND( SUM(b.����) * ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  +  (dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0),2)   -  CONVERT(DECIMAL(18,2) " & _
         " ,SUM(b.����) * ISNULL( cd.��˰���� / cc.��Ʒ��,0)) AS �ӹ��ѽ�� , CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( cd.��˰���� / cc.��Ʒ��,0)) AS �͹��Ͻ�� " & _
         "  FROM erpdata .. tblStockNumTree a INNER JOIN erpdata .. tblStocksqfhsub b ON b.��� = a.��� INNER JOIN erpdata .. tblStockSQfh c  ON c.���ݱ�� = b.���ݱ�� AND c.��� = b.������� " & _
         " INNER JOIN erpdata..tblstock y ON y.�ⷿ���� = c.�ֿ���   INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE, 1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox " & _
         " , SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job  FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.��� = aa.qbox " & _
         " INNER JOIN (SELECT dn.Delivery,   dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State,dn.postalcode, dn.countrykey, dn.contactname,dn.phone, dn.SalesDocument " & _
         " ,  dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE  FROM ERPBASE .. tblCustomerShippingUp dn " & _
         " WHERE dn.Delivery IN ('" & strDNList & "')  GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State, dn.postalcode,  dn.countrykey,dn.contactname,dn.phone " & _
         " , dn.SalesDocument, dn.PurchasingDocNo, dn.BatchNumber, dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE) d    ON d.Delivery = a.DN  AND d.BatchNumber = aa.job " & _
         " INNER JOIN erpdata .. tblstocknumtree f  ON f.��� = a.�ϼ����  INNER JOIN ERPBASE..tblmappingData qq ON qq.SUBSTRATEID = b.���̿���� AND qq.LOTID = b.������ " & _
         " LEFT JOIN ERPBASE..tblCustomerOI bb  ON CONVERT(VARCHAR(100), bb.ID) = qq.FILENAME AND bb.SOURCE_BATCH_ID = qq.LOTID  LEFT JOIN ERPBASE..tblToInRec_Wafer cc " & _
         " ON cc.��ԲID = REPLACE(qq.SUBSTRATEID,'+','')   AND cc.���� = qq.LOTID  LEFT JOIN ERPBASE..TblToInSub cd ON cd.��ⵥ��� = cc.��ⵥ��� AND cd.�������� = cc.���� " & _
         " LEFT JOIN erptemp..tblBB_CSRPO dd   ON dd.PO_NUM = bb.PO_NUM AND dd.FAB_DEVICE = bb.MPN_DESC LEFT JOIN  erpdata..tblSalerec j ON j.���ݱ�� = c.���ݱ�� AND j.������� = b.������� " & _
         " AND j.С��� = b.��� LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName  WHERE a.DN IN ('" & strDNList & "')  GROUP BY b.���ݱ��, c.��������, ISNULL(dn_address_new, d.shiptoname) " & _
         "  ,   ISNULL(x.ship_to_street1_new, d.shiptostreet1), ISNULL(x.ship_to_street2_new, d.shiptostreet2),ISNULL(x.ship_to_street3_new, d.shiptostreet3), ISNULL(x.city_new, d.city) " & _
         "  ,  ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode),ISNULL(x.country_new, d.countrykey),ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) " & _
         " ,  d.SalesDocument, d.PurchasingDocNo,f.���, b.�Ϻ�,d.MarketingPN,d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber,   a.DN " & _
         " , J.���۵����, bb.PO_NUM, cd.��˰����, cc.��Ʒ��, dd.WAFER_PRICE, qq.PASSBINCOUNT, dd.die_price, y.������  order by shiptoname,DN,��� "

          
          
          
          
          
          
    ElseIf cmbCombo1(0).ListIndex = 6 Then  '��汨��
        If cmbCombo1(1).text = "9000" Then
            strSql = " SELECT * FROM DBO.Vw_InvStockRptFor37By9000 "
            strSql = "select * from [dbo].[Vw_InvStockRptFor37By9000_new_temp] "
        
            strSql = "select RECEIVE_DATE,[Wafer Type],[Assy Part#],[Fab Lot],ID,[D/C],[QTY(���غ������)],Job#,Bag#,Comment,NCMR,[HT Part#],[��˰/�Ǳ�˰], [�ⷿ���] from erpbase..Vw_InvStockRptFor37By9000_temp a where a.�ⷿ��� in ('44','45') " & _
"Union select RECEIVE_DATE,[Wafer Type],[Assy Part#],[Fab Lot],ID,[D/C],sum([QTY(���غ������)]),Job#,Bag#,Comment,NCMR,[HT Part#],[��˰/�Ǳ�˰],[�ⷿ���] from erpbase..Vw_InvStockRptFor37By9000_new_temp a where a.�ⷿ��� in ('44','45') " & _
"group by RECEIVE_DATE,[Wafer Type],[Assy Part#],[Fab Lot],ID,[D/C],Job#,Bag#,Comment,NCMR,[HT Part#],[��˰/�Ǳ�˰],[�ⷿ���] " & _
"order by [Assy Part#]"
            
            
        Else
            strSql = "SELECT 0 ѡ��,RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Reel_Size,Remark" & _
                     ",Move_in_Date FROM Vw_InvStockRptFor37 Where 2>1 "
            If txt(1).text <> "" Then
                strSql = strSql & " And JOB_NO='" & Trim(txt(1).text) & "'"
            End If
            If cmbCombo1(1).text <> "����" Then
                strSql = strSql & " And �ֿ�����='" & Trim(cmbCombo1(1).text) & "'"
            End If
        End If
    ElseIf cmbCombo1(0).ListIndex = 7 Then  'SMTCList
        strSql = "SELECT 0 ѡ��,Invoice_No,Carton_No,PartName,LotID,QTY,Job_No,DATE_CODE " & _
                 " FROM Vw_InvShippedSMTCListFor37 " & _
                 " WHERE ��������>='" & DTP(0).Value & "' and ��������<'" & DTP(1).Value + 1 & "' "
        If txt(0).text <> "" Then
            strSql = strSql & " And ���ݱ��='" & Trim(txt(0).text) & "'"
        End If
        If txt(1).text <> "" Then
            strSql = strSql & " And Job_No='" & Trim(txt(1).text) & "'"
        End If
    ElseIf cmbCombo1(0).ListIndex = 8 Then  'Shipped
        strSql = "SELECT 0 ѡ��,RECEIVE_DATE,TEST_DC,LOCATION,DEVICE_NAME,JOB_NO,LOT_NO,QTY,DATA_CODE,Comment,Remark,SO as [DN#],SHIPPED_DATE,Cust_Name " & _
                 " FROM Vw_InvShippedRptFor37 " & _
                 " WHERE SHIPPED_DATE>='" & DTP(0).Value & "' and SHIPPED_DATE<'" & DTP(1).Value + 1 & "' "
        If txt(0).text <> "" Then
            strSql = strSql & " And ���ݱ��='" & Trim(txt(0).text) & "'"
        End If
        If txt(1).text <> "" Then
            strSql = strSql & " And JOB_NO='" & Trim(txt(1).text) & "'"
        End If
    ElseIf cmbCombo1(0).ListIndex = 9 Then  'Shipped
          strSql = "SELECT 0 AS  ѡ��,y.������ + b.���ݱ�� AS ���ݱ��,d.Delivery, CONVERT(VARCHAR(100), c.��������,23) AS ��������, ISNULL(dn_address_new, d.shiptoname) AS shiptoname, ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1 " & _
          " ,  ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2, ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,  ISNULL(x.city_new, d.city) AS city,  ISNULL(x.dn_st_new, d.State) AS state,  ISNULL(x.postal_code_new, d.postalcode) AS postalcode, " & _
          "  ISNULL(x.country_new, d.countrykey) AS countrykey,  ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone, d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo " & _
         " ,f.��� ,b.�Ϻ� ,d.MarketingPN,SUM(b.����),d.BatchNumber,d.DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no, d.customerPartNumber, ROUND(CAST(f.���� AS FLOAT) * 0.4, 2) ����, f.���� ë��, " & _
         " f.�ߴ� FROM erpdata..tblStockNumTree a INNER JOIN erpdata..tblStocksqfhsub b ON b.��� = a.��� INNER JOIN erpdata..tblStockSQfh c ON c.���ݱ�� = b.���ݱ�� AND c.��� = b.������� INNER JOIN erpdata..tblstock y " & _
         " ON y.�ⷿ���� = c.�ֿ���  INNER JOIN (SELECT a.BOX_ID,  SUBSTRING(a.KEY_VALUE,  1, CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job " & _
         " FROM erpdata .. tblErpInStockDetailInfo a  WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.��� = aa.qbox  INNER JOIN (SELECT dn.Delivery, dn.shiptoname, dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3 " & _
         " , dn.city, dn.State, dn.postalcode, dn.countrykey, dn.contactname,  dn.phone, dn.SalesDocument,  dn.PurchasingDocNo, dn.BatchNumber, SUM(dn.Quantity) AS Quantity, dn.customerPartNumber, dn.MarketingPN, dn.DATE_CODE " & _
         " FROM ERPBASE..tblCustomerShippingUp dn WHERE dn.Delivery IN ('" & strDNList & "') GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1, dn.shiptostreet2, dn.shiptostreet3,dn.city, dn.State, dn.postalcode, dn.countrykey " & _
         " , dn.contactname,dn.phone, dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber, dn.customerPartNumber, dn.MarketingPN,dn.DATE_CODE) d  ON d.Delivery = a.DN AND d.BatchNumber = aa.job " & _
         " INNER JOIN erpdata..tblStockNumTree e ON e.��� = b.��� INNER JOIN erpdata..tblstocknumtree f  ON f.��� = e.�ϼ���� LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName " & _
        "  WHERE a.DN IN ('" & strDNList & "')GROUP BY  b.���ݱ��,c.��������,ISNULL(dn_address_new, d.shiptoname) , ISNULL(x.ship_to_street1_new, d.shiptostreet1) ,ISNULL(x.ship_to_street2_new, d.shiptostreet2) " & _
         " ,ISNULL(x.ship_to_street3_new, d.shiptostreet3),ISNULL(x.city_new, d.city),ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode), ISNULL(x.country_new, d.countrykey) " & _
         " ,ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) , d.SalesDocument,d.PurchasingDocNo ,f.���,b.�Ϻ� ,d.MarketingPN,d.BatchNumber ,d.DATE_CODE " & _
         " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber, f.���� , f.�ߴ�,y.������,d.Delivery order by shiptoname, Delivery,���"
   
    End If
    '��ֵ��FRA(1)�� INIadoCon
    Fra(1).Caption = cmbCombo1(0).text
    If rs.State = adStateOpen Then rs.Close
    If cmbCombo1(0).ListIndex = 0 Or cmbCombo1(0).ListIndex = 1 Or cmbCombo1(0).ListIndex = 2 Or cmbCombo1(0).ListIndex = 3 Then
        rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    Else
        rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    End If
    Fps(0).MaxRows = 0
    Set RsClone = Nothing
    
    If rs.EOF Then
        MsgBox "�����ݣ�"
    End If
    
    If Not rs.EOF Then
        Set RsClone = rs.Clone '��¡һ�����ݵ���һ�����ݼ��У�Ϊ����ʹ��
        With Fps(0)
            .MaxRows = 0
            Set .DataSource = rs
            .MaxRows = rs.RecordCount
        End With
    End If
    rs.Close
    '���⼸���������ӻ�����λ
    CalcTotal
    
'      With lstLotID
'
'        For i = 0 To .ListCount - 1
'            .Selected(i) = False
'        Next
'
'    End With
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
            dblTotal = dblTotal + Val(Trim$(.text))
        Next
        If dblTotal > 0 Then '��ʾ������
            .MaxRows = .MaxRows + 1
            .SetText colTotal, .MaxRows, dblTotal
        End If
        .DeleteCols fpSDetail.E_CHOOSE, 1
        .MaxCols = .MaxCols - 1
    End With
    
End Sub

Private Sub cmdUpload_Click()
Dim strFilePath         As String
Dim strFileName         As String
Dim strSql              As String
Dim image_Data()        As Byte         'ͼƬ������
Dim rs                  As New ADODB.Recordset
    '��ͼƬ
    Com.Filter = "�ϴ��ļ�(*.xls,*.xlsx)|*.xls;*.xlsx"
    Com.ShowOpen '�򿪶Ի���
    strFilePath = Trim(Com.filename)  '����·��
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1) '�ļ���
    '��ʼ���浽���Ͽ�
    '����ת��Ϊ��
    Open strFilePath For Binary As #1
    ReDim image_Data(LOF(1) - 1)
    Get #1, , image_Data()
    Close #1
    '��ѯ�Ƿ񱣴����ͼƬ
    
    
    
    'strsql = "Select * From TblPMC_PicInfo Where FileName='" & Trim$(strFilename) & "' For Update"
    
    strSql = " SELECT * FROM erpdata..tblSystemTemplet WHERE SYS_NAME='����' AND UPPER(TEMPLETNAME)='Invoice.xls'"
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon2, adOpenKeyset, adLockOptimistic
    If Not rs.EOF Then
        rs("SYS_NAME") = "����"
        rs("TEMPLETNAME") = "Invoice.xls"
        rs("FILECONTENT") = image_Data()
        rs.Update
    Else
        rs.AddNew
        rs("FileName") = strFileName
        rs("FilePath") = strFilePath
        rs("FileComent") = image_Data()
        rs("Flag") = 1
        '�ǵ�������ݿ���txt��ŵ�·����txt��
        rs.Update
    End If
    rs.Close
    
    MsgBox "�ϴ��ɹ�", vbInformation, "��ʾ"
    
End Sub

Private Sub Command1_Click()
Dim strExportName As String
Dim i             As Integer

cmdReport.Enabled = False
strExportName = GetExcelName(Trim(Fra(1).Caption))
Call ShippingPackinglistExportPrintExcel(order, strExportName, 0)    'Shipping Packinglist

MsgBox "װ�䵥�Ѿ��������", vbInformation, "��ʾ"
Call ShippingInvoiceExportPrintExcel(order, strExportName)          'Shipping invoice
cmdReport.Enabled = True
MsgBox "Invoice�Ѿ��������", vbInformation, "��ʾ"

End Sub

Private Sub Command2_Click()
Dim i As Integer
With lstLotID

        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next

End With
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
    
    DirQrShare = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\37\jpg"
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
Dim rs                  As New ADODB.Recordset
Dim image_filename      As String
Dim temp_image()        As Byte
Dim i                   As Integer
    '�趨���״̬
    Screen.MousePointer = 0
    strSql = "Select * From TblPMC_PicInfo Where Flag=1 Order by Create_Date"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        For i = 1 To rs.RecordCount
            '����ͼƬ
            temp_image = rs("FileComent")
            image_filename = DirShare & "\" & rs("FileName")
            Open image_filename For Binary As #1
            Close #1
          '  Put #1, , temp_image()
          '  Close #1
            rs.MoveNext
        Next
    End If
    rs.Close

End Sub
'��ʼ���ؼ�
Private Sub InitCtrl()
Dim i                   As Integer
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
 Dim Rs2                  As New ADODB.Recordset
    
    strSql = "select  distinct dn_num from packing_detailed  where create_date > sysdate - 30  order by dn_num desc"
     
    Set Rs2 = Get_OracleRs(strSql)
'Show
lstLotID.Clear
If Not Rs2.EOF Then

    Do While Not Rs2.EOF
        lstLotID.AddItem Trim("" & Rs2!DN_NUM)
        Rs2.MoveNext
    Loop
Else
End If
    
    
    strdjbh = ""
    '���ص�������
    strSql = "select para_1 from tblsys_parameter where sysname='TSVSYS' and kind='Semtech����' order by id "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    cmbCombo1(0).Clear
    If Not rs.EOF Then
        Do While Not rs.EOF
            cmbCombo1(0).AddItem Trim$("" & rs!para_1)
            rs.MoveNext
        Loop
        cmbCombo1(0).ListIndex = 0
    End If
    rs.Close
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
    .TypeMaxEditLen = 500
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
        .Col = fpSDetail.E_CHOOSE   'ѡ��
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '�趨�п�
        .ColWidth(-1) = 10
        .ColWidth(fpSDetail.E_CHOOSE) = 4
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
Private Sub fps_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim i           As Long
Dim j           As Integer
Dim strTmp      As String
    
    '�����������⴦��
    
    cmdReport.Enabled = True
    cmdExprot.Enabled = True
    
    
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

        .Col = fpSDetail.E_CHOOSE
        .Row = Row
        .Value = Abs(Val(.Value) - 1)
'        strDJBH = ""
        If Val(.Value) = 1 Then
            '������һ���ĵ���ѡ����
            .Col = fpSDetail.e_DJBH
            .Row = Row
            strTmp = Trim$(.text)
'            strDJBH = Trim$(.Text) '���õĵ��ݱ�ţ��ڵ�����ӡʱ���õ�
            For i = 1 To .MaxRows
                .Row = i
                .Col = fpSDetail.e_DJBH
                If Trim$(.text) = strTmp Then
                    .Col = fpSDetail.E_CHOOSE
                    .Value = 1
                    .Col = -1
                    .ForeColor = &HFF8080
                End If
            Next
            
            order = strTmp & "'" & "," & "'" & order
            
        Else
            '������һ���ĵ���ѡ����
            .Col = fpSDetail.e_DJBH
            .Row = Row
            strTmp = Trim$(.text)
'            strDJBH = Trim$(.Text) '���õĵ��ݱ�ţ��ڵ�����ӡʱ���õ�
            For i = 1 To .MaxRows
                .Row = i
                .Col = fpSDetail.e_DJBH
                If Trim$(.text) = strTmp Then
                    .Col = fpSDetail.E_CHOOSE
                    .Value = 0
                    .Col = -1
                    .ForeColor = vbBlack
                End If
            Next
            
            order = Replace(order, strTmp & "'" & "," & "'", "")
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
            .Col = fpSDetail.E_CHOOSE  'ѡ��
            If .Value = 1 Then
                intCount = intCount + 1
                .Col = fpSDetail.e_DJBH '���ݱ��
                If InStr(strdjbh, Trim$(.text)) <= 0 Then
                    strdjbh = strdjbh + Trim$(.text) + ","
                    strdjbh1 = Mid(strdjbh, 2, Len(strdjbh)) + Trim$(.text) + ","
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
Public Sub SEDCExportPrintExcel(ByVal rs As ADODB.Recordset, ByVal strExName As String)
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

    
    If rs.RecordCount <= 0 Then
        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
    
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "��ʼ������..."
    
    strFileName = DirShare & "\SEDC.xls" 'Ҫ�򿪵��ļ�
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
    rs.MoveFirst    '���ݼ��ƶ�����һ��
    
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        
        lngRows = 3
        IntInertRow = rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = rs.RecordCount
        
'        ClsP.ShowProgress 50, "���ڵ���..."
        
        Do While Not rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & rs.Fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & rs.Fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & rs.Fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & rs.Fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & rs.Fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & rs.Fields(10).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(11).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(12).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(13).Value)
            wkst.Cells(lngRows, 13) = Trim$("" & rs.Fields(14).Value)
            
            lngRows = lngRows + 1
            rs.MoveNext
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
            If Err.number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
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
    MsgBox Err.DESCRIPTION, vbInformation, "��ʾ��"
    Exit Sub
End Sub

'Packing list
Public Sub InputPackinglistExportPrintExcel(ByVal rs As ADODB.Recordset, ByVal strExName As String)
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

    If rs.RecordCount <= 0 Then
        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
'   ClsP.Init 100, True
'   ClsP.ShowProgress 10, "��ʼ������..."
    
    strFileName = DirShare & "\output_packing_list.xls" 'Ҫ�򿪵��ļ�
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
    strXH = Mid$(strExName, InStrRev(strExName, "_"))    '��ȡ���
    rs.MoveFirst    '���ݼ��ƶ�����һ��
        
    
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        '��ֵ��Excel�У���ͷ
        
        wkst.Cells(8, 12) = Format(DATE, "YYYY/mm/DD")
        wkst.Cells(9, 12) = "HTKS-SEDC" & Format(DATE, "YY-MMDD") & strXH
        
        lngRows = 17
        IntInertRow = rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = rs.RecordCount
        
'        ClsP.ShowProgress 50, "���ڵ���..."
        
        Do While Not rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & rs.Fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & rs.Fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & rs.Fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & rs.Fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & rs.Fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & rs.Fields(10).Value)
            
            'jiayun �޸� ����0��Ϊ��ֵ
            
'            wkst.Cells(lngRows, 10) = Trim$("" & Rs.fields(11).Value)
'            wkst.Cells(lngRows, 11) = Trim$("" & Rs.fields(12).Value)
            wkst.Cells(lngRows, 10) = ""
            wkst.Cells(lngRows, 11) = ""
            
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(13).Value)
            wkst.Cells(lngRows, 13) = Trim$("" & rs.Fields(14).Value)
            wkst.Cells(lngRows, 14) = Trim$("" & rs.Fields(15).Value)
            
            'jiayun add Bag#
            wkst.Cells(lngRows, 15) = Trim$("" & rs.Fields(16).Value)
            
            DblNum = DblNum + Val(Trim$("" & rs.Fields(10).Value))
            DblAmt = DblAmt + Val(Trim$("" & rs.Fields(12).Value))
            
            lngRows = lngRows + 1
            rs.MoveNext
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
            If Err.number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
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
    MsgBox Err.DESCRIPTION, vbInformation, "��ʾ��"
    Exit Sub
End Sub
'Invoice
Public Sub InputInvoiceExportPrintExcel(ByVal rs As ADODB.Recordset, ByVal strExName As String)
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
    
    If rs.RecordCount <= 0 Then
        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "��ʼ������..."
    
    strFileName = DirShare & "\output_invoice.xls" 'Ҫ�򿪵��ļ�
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
    strXH = Mid$(strExName, InStrRev(strExName, "_"))    '��ȡ���
    rs.MoveFirst    '���ݼ��ƶ�����һ��
    
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        '��ֵ��Excel�У���ͷ
        wkst.Cells(8, 11) = Format(DATE, "YYYY/mm/DD")
        wkst.Cells(9, 11) = "HTKS-SEDC" & Format(DATE, "YY-MMDD") & strXH
        
        lngRows = 17
        IntInertRow = rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = rs.RecordCount
        
'        ClsP.ShowProgress 50, "���ڵ���..."
        
        Do While Not rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & rs.Fields(2).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & rs.Fields(3).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(6).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & rs.Fields(7).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & rs.Fields(8).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & rs.Fields(9).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & rs.Fields(10).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(11).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(12).Value)
            
            DblNum = DblNum + Val(Trim$("" & rs.Fields(10).Value))
            DblAmt = DblAmt + Val(Trim$("" & rs.Fields(12).Value))
            
            
            lngRows = lngRows + 1
            rs.MoveNext
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
            If Err.number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
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
    MsgBox Err.DESCRIPTION, vbInformation, "��ʾ��"
    Exit Sub
End Sub

'Daily_inventory
Public Sub Daily_InvExportPrintExcel(ByVal rs As ADODB.Recordset, ByVal strExName As String)
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

    
    If rs.RecordCount <= 0 Then
        MsgBox "û��Ҫ���������ϣ�", vbInformation, "��ʾ��"
        Exit Sub
    End If
'    ClsP.Init 100, True
'    ClsP.ShowProgress 10, "��ʼ������..."
    
    strFileName = DirShare & "\Daily_inventory_report.xls" 'Ҫ�򿪵��ļ�
    strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
    strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
    rs.MoveFirst    '���ݼ��ƶ�����һ��
    
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
'        ExApp.ActiveWindow.DisplayGridlines = False
        
        DblNum = 0
        DblAmt = 0
        
        lngRows = 3
        IntInertRow = rs.RecordCount
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
        Next i
        IntMaxDetailRow = rs.RecordCount
        
'        ClsP.ShowProgress 50, "���ڵ���..."
        
        Do While Not rs.EOF
            wkst.Cells(lngRows, 1) = Trim$("" & rs.Fields(1).Value)
            wkst.Cells(lngRows, 2) = Trim$("" & rs.Fields(2).Value)
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(3).Value)
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(lngRows, 6) = Trim$("" & rs.Fields(6).Value)
            wkst.Cells(lngRows, 7) = Trim$("" & rs.Fields(7).Value)
            wkst.Cells(lngRows, 8) = Trim$("" & rs.Fields(8).Value)
            wkst.Cells(lngRows, 9) = Trim$("" & rs.Fields(9).Value)
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(10).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(11).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(12).Value)
            
            lngRows = lngRows + 1
            rs.MoveNext
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
            If Err.number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
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
    MsgBox Err.DESCRIPTION, vbInformation, "��ʾ��"
    Exit Sub
End Sub

'shipping Packing list
Public Sub ShippingPackinglistExportPrintExcel(ByVal ordertemp As String, _
                                               ByVal strExName As String, lxflag As Integer)

    Dim strSql         As String

    Dim lngRows        As Long

    Dim rsQuery        As Excel.QueryTable

    Dim ExApp          As Excel.Application

    Dim wkbk           As New Workbook

    Dim wkst           As New Worksheet

    Dim i              As Long

    Dim j              As Long

    Dim IntCols        As Integer

    Dim strCols        As String

    Dim strFileName    As String

    Dim IntInertRow    As Integer, IntMaxDetailRow As Integer

    Dim DblNum         As Double

    Dim DblAmt         As Double  '�ܽ��

    Dim intBoxNum      As Integer '����

    Dim strPBigBox     As String  'ǰ���

    Dim strNBigBox     As String  '�����

    Dim IntBMegerRow   As Integer

    Dim IntEMegerRow   As Integer

    Dim DblJZ          As Double   '����

    Dim DblMZ          As Double   'ë��

    Dim DblJZ1         As Double   '����

    Dim DblMZ1         As Double   'ë��

    Dim DblJZ2         As Double   '����

    Dim DblMZ2         As Double   'ë��

    Dim intBegin       As Integer

    Dim strdjTmp       As String

    Dim SD             As String

    Dim SD1            As String

    Dim strTmp()       As String

    Dim strExtsion     As String '��׺��

    Dim strNewFullPath As String '��Excel�ļ�
    
    Dim strNewFullPathNew As String
    
    Dim DirFileShare1 As String

    Dim RsNew          As New ADODB.Recordset  '��¼����ĸ��������������������

    Dim rs             As New ADODB.Recordset

    Dim dnnum          As String

    Dim dnnum1         As String

    strNewFullPathNew = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\37\excel"
    
    dnnum = ""
    dnnum1 = ""
    
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
    
    If lxflag = 0 Then
         strFileName = DirShare & "\shipping_packing_list.xlsx" 'Ҫ�򿪵��ļ�
         strExtsion = ".xls"
         strExtsion = Mid$(strFileName, InStrRev(strFileName, "."))      '��ȡ��׺��
         strNewFullPath = DirFileShare & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��

    Else
         strFileName = DirShare & "\shipping_packing_list_2.xlsx" 'Ҫ�򿪵��ļ�
         strExtsion = ".xls"
         DirFileShare1 = "C:\�ϰ�_2"
         strNewFullPath = DirFileShare1 & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
         If Dir("C:\�ϰ�_2", vbDirectory) = "" Then '�ж��ļ����Ƿ����
            MkDir ("C:\�ϰ�_2") '�����ļ��� msgbox ("�������")
            MsgBox ("�ļ����Ѵ�����·��Ϊ C:\�ϰ�_2")
         End If
    End If
    
    '    strSql = "SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
    '                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
    '                 ",���,�Ϻ�,mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,����,ë��,�ߴ� " & _
    '                 " FROM Vw_InvShippedPLFor37 a  where ���ݱ�� in ('" & Ordertemp & "')  order by ���"

    strSql = " select * from ( SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & ",���,�Ϻ�,replace(mpn_desc,'.P2','') AS mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,����,ë��,�ߴ� " & " FROM Vw_InvShippedPLFor37_NEW a  where ���ݱ�� in ('" & ordertemp & "')  " & " union all " & "SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & ",���,�Ϻ�,replace(mpn_desc,'.P2','') AS mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,����,ë��,�ߴ� " & " FROM Vw_InvShippedPLFor37 a  where ���ݱ�� in ('" & ordertemp & "') ) x order by x.��� "
    



        Dim strDNList As String
        With lstLotID

            For i = 0 To .ListCount - 1
        
                If .Selected(i) = True Then
                    strDNList = strDNList & Trim$("" & .List(i)) & "','"
                     

                End If

            Next

        End With

        
strDNList = Mid(strDNList, 1, Len(strDNList) - 3)

'    strSql = "SELECT 0 ѡ��,h.������ + a.���ݱ�� ���ݱ��,c.delivery,dbo.usp_date(a.��������) ��������,ISNULL(dn_address_new, c.shiptoname) AS shiptoname,ISNULL(x.ship_to_street1_new, c.shiptostreet1) AS shiptostreet1, " & _
'"ISNULL(x.ship_to_street2_new, c.shiptostreet2) AS shiptostreet2,ISNULL(x.ship_to_street3_new, c.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, c.city) AS city, ISNULL(x.dn_st_new, c.State) AS state, " & _
'"ISNULL(x.postal_code_new, c.postalcode) AS postalcode,ISNULL(x.country_new, c.countrykey) AS countrykey,ISNULL(x.contact_new, c.contactname) AS contactname,ISNULL(x.phone_new, c.phone) AS phone, " & _
'"c.SalesDocument,'''' + c.PurchasingDocNo AS PurchasingDocNo,RTRIM(erpdata.dbo.f_getparent(b.���)) ���,b.�Ϻ�, " & _
'"CASE WHEN RTRIM(gg.MPN_DESC) = 'UCLAMP0541Z.TFT' THEN RTRIM(gg.MPN_DESC) + '.P2' ELSE REPLACE(REPLACE(gg.MPN_DESC, '.P2', ''), '.P3', '') END AS mpn_desc,SUM(b.����) ����,c.batchnumber, " & _
'"hh.CREATE_DATE DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no,c.customerPartNumber,ROUND(CAST(f.���� AS FLOAT) * 0.4, 2) ����,f.���� ë��,f.�ߴ� FROM erpdata .. tblStockSQfh a " & _
'"INNER JOIN erpdata .. tblStocksqfhsub b ON a.���ݱ�� = b.���ݱ�� AND a.��� = b.������� INNER JOIN erpdata .. tblStockNumTree g ON g.��� = b.��� INNER JOIN erpdata .. tblStockNumTree f " & _
'"ON f.��� = g.�ϼ���� INNER JOIN (SELECT a.BOX_ID,SUBSTRING(a.KEY_VALUE,1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox,SUBSTRING(a.KEY_VALUE,CHARINDEX('|', a.KEY_VALUE) + 1,10) AS job " & _
'"FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON g.��� = aa.qbox INNER JOIN (SELECT dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, " & _
'"dn.shiptostreet3, dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber " & _
'"FROM tblCustomerShippingUp dn GROUP BY dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname, " & _
'"dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,dn.customerPartNumber) c ON c.Delivery = g.DN AND c.BatchNumber = aa.job INNER JOIN dbo.tblstock h " & _
'"ON CONVERT(NVARCHAR(4), h.�ⷿ����) = CONVERT(NVARCHAR(4), a.�ֿ���) INNER JOIN ERPBASE .. tblmappingData ff ON ff.SUBSTRATEID = b.���̿���� " & _
'"INNER JOIN ERPBASE .. tblCustomerOI gg ON CONVERT(VARCHAR(100), gg.ID) = ff.FILENAME AND gg.SOURCE_BATCH_ID = ff.LOTID AND gg.CUSTOMERSHORTNAME = '37' INNER JOIN erpbase .. weight37 hh " & _
'"ON hh.WAFERID = REPLACE(b.���̿����, '+', '') INNER JOIN erpdata .. tblErpInStockMainData jj ON jj.BOX_ID = aa.BOX_ID LEFT JOIN erptemp .. dn_address x ON dn_address = c.ShipToName " & _
'"WHERE a.�ͻ����� = '37' AND a.���ݱ�� LIKE 'F%' AND a.�������� >= CONVERT(VARCHAR(100), GETDATE() - 5, 23) AND c.Delivery = g.DN " & _
'"and c.Delivery in ('" & strDNList & "') GROUP BY h.������, a.���ݱ��, c.delivery,dbo.usp_date(a.��������), c.shiptoname,c.shiptostreet1,c.shiptostreet2,c.shiptostreet3,c.city,c.State,c.postalcode, " & _
'"c.countrykey,c.contactname,c.phone,c.SalesDocument,c.PurchasingDocNo,erpdata.dbo.f_getparent(b.���),b.�Ϻ�,gg.MPN_DESC,c.batchnumber,hh.CREATE_DATE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2), " & _
'"c.customerPartNumber,f.����,f.�ߴ�,dn_address_new,x.ship_to_street1_new,x.ship_to_street2_new,x.ship_to_street3_new,x.city_new,x.dn_st_new,x.postal_code_new,x.country_new,x.contact_new,x.phone_new order by Delivery"

 strSql = "SELECT 0 AS  ѡ��,y.������ + b.���ݱ�� AS ���ݱ��,d.Delivery, CONVERT(VARCHAR(100), c.��������,23) AS ��������, ISNULL(dn_address_new, d.shiptoname) AS shiptoname, ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1 " & _
 " ,  ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2, ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,  ISNULL(x.city_new, d.city) AS city,  ISNULL(x.dn_st_new, d.State) AS state,  ISNULL(x.postal_code_new, d.postalcode) AS postalcode, " & _
 "  ISNULL(x.country_new, d.countrykey) AS countrykey,  ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone, d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo " & _
 " ,f.��� ,b.�Ϻ� ,d.MarketingPN,SUM(b.����),d.BatchNumber,d.DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no, d.customerPartNumber, ROUND(CAST(f.���� AS FLOAT) * 0.4, 2) ����, f.���� ë��, " & _
 " f.�ߴ� FROM erpdata..tblStockNumTree a INNER JOIN erpdata..tblStocksqfhsub b ON b.��� = a.��� INNER JOIN erpdata..tblStockSQfh c ON c.���ݱ�� = b.���ݱ�� AND c.��� = b.������� INNER JOIN erpdata..tblstock y " & _
 " ON y.�ⷿ���� = c.�ֿ���  INNER JOIN (SELECT a.BOX_ID,  SUBSTRING(a.KEY_VALUE,  1, CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job " & _
 " FROM erpdata .. tblErpInStockDetailInfo a  WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.��� = aa.qbox  INNER JOIN (SELECT dn.Delivery, dn.shiptoname, dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3 " & _
 " , dn.city, dn.State, dn.postalcode, dn.countrykey, dn.contactname,  dn.phone, dn.SalesDocument,  dn.PurchasingDocNo, dn.BatchNumber, SUM(dn.Quantity) AS Quantity, dn.customerPartNumber, dn.MarketingPN, dn.DATE_CODE " & _
 " FROM ERPBASE..tblCustomerShippingUp dn WHERE dn.Delivery in ('" & strDNList & "') GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1, dn.shiptostreet2, dn.shiptostreet3,dn.city, dn.State, dn.postalcode, dn.countrykey " & _
 " , dn.contactname,dn.phone, dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber, dn.customerPartNumber, dn.MarketingPN,dn.DATE_CODE) d  ON d.Delivery = a.DN AND d.BatchNumber = aa.job " & _
 " INNER JOIN erpdata..tblStockNumTree e ON e.��� = b.��� INNER JOIN erpdata..tblstocknumtree f  ON f.��� = e.�ϼ���� LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName " & _
"  WHERE a.DN IN ('" & strDNList & "')GROUP BY  b.���ݱ��,c.��������,ISNULL(dn_address_new, d.shiptoname) , ISNULL(x.ship_to_street1_new, d.shiptostreet1) ,ISNULL(x.ship_to_street2_new, d.shiptostreet2) " & _
 " ,ISNULL(x.ship_to_street3_new, d.shiptostreet3),ISNULL(x.city_new, d.city),ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode), ISNULL(x.country_new, d.countrykey) " & _
 " ,ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) , d.SalesDocument,d.PurchasingDocNo ,f.���,b.�Ϻ� ,d.MarketingPN,d.BatchNumber ,d.DATE_CODE " & _
 " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber, f.���� , f.�ߴ�,y.������,d.Delivery order by Delivery,���"

    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
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

    '    Rs.MoveFirst    '���ݼ��ƶ�����һ��
    '    '---------------------------------------------------------------------
    If rs.RecordCount > 0 Then
        '        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
        '        ExApp.ActiveWindow.DisplayGridlines = False
        
        ' wkbk.ActiveSheet.Range(A3).Select
        
        DblNum = 0
        DblJZ = 0
        DblMZ = 0
        '

        '��ֵ��Excel�У���ͷ
        'wkst.Cells(8, 2) = Trim$("" & rs.Fields(2).Value)
        wkst.Cells(8, 17) = Trim$("" & rs.Fields(3).Value)
        'shipto  ����װ�䵥ʱ��ִ�� ����Sold to = Ship to 2019-12-5
        If lxflag = 1 Then
            wkst.Cells(10, 2) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(11, 2) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(12, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
            wkst.Cells(13, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
            wkst.Cells(14, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
            wkst.Cells(15, 2) = ""
        End If
        'sold
        wkst.Cells(17, 2) = Trim$("" & rs.Fields(4).Value)
        wkst.Cells(18, 2) = Trim$("" & rs.Fields(5).Value)
        wkst.Cells(19, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
        wkst.Cells(20, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
        wkst.Cells(22, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
        wkst.Cells(23, 2) = ""
        wkst.Cells(17, 17) = Trim$("" & rs.Fields(11).Value) 'To
        wkst.Cells(25, 3) = Trim$("" & rs.Fields(14).Value)
        wkst.Cells(25, 6) = Trim$("'" & rs.Fields(15).Value)
        
        lngRows = 28
        
        IntInertRow = rs.RecordCount * 2

        For i = 1 To IntInertRow - 1
            wkst.Rows(lngRows & ":" & lngRows).Select
            ExApp.Selection.Copy
            ExApp.Selection.Insert Shift:=xlDown
            wkst.Rows(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '�߿���
        Next i

        IntMaxDetailRow = rs.RecordCount
        
        '        ClsP.ShowProgress 50, "���ڵ���..."
        
        IntBMegerRow = 27
        IntEMegerRow = 30
        intBegin = 1

        Dim QBX As String

        For i = 0 To rs.RecordCount - 1

            '            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '���
            If dnnum1 <> Trim$("" & rs.Fields(2).Value) Then
                dnnum = Trim$("" & rs.Fields(2).Value) + "/" + dnnum
                dnnum1 = Trim$("" & rs.Fields(2).Value)

            End If
             
            strPBigBox = Trim$("" & rs.Fields(16).Value) '���

            'QBX = strPBigBox
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & rs.Fields(16).Value) '���
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
            
            If SD <> Trim$("" & rs.Fields(14).Value) Then
                SD = Trim$("" & rs.Fields(14).Value)
                SD1 = SD1 & SD & " "

            End If

            wkst.Cells(25, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & rs.Fields(19).Value)) / 1000 '������Ϊ��ǧΪ��λ
            DblNum = DblNum + Val(Trim$("" & rs.Fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(21).Value) 'datacode
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(22).Value) 'lotno

            If strPBigBox <> QBX Then
                wkst.Cells(lngRows, 14) = Trim$("" & rs.Fields(24).Value) '����
                wkst.Cells(lngRows, 15) = "KG"   '���ص�λ
                wkst.Cells(lngRows, 18) = "KG"   'ë�ص�λ
                wkst.Cells(lngRows, 19) = Trim$("" & rs.Fields(26).Value)   '�ߴ�
                wkst.Cells(lngRows, 17) = Trim$("" & rs.Fields(25).Value)   'ë��
            
            End If
           
            DblJZ1 = Val(Trim$("" & rs.Fields(24).Value))
            
            If strPBigBox <> QBX Then
                DblJZ = DblJZ1 + DblJZ

            End If

            DblMZ1 = Val(Trim$("" & rs.Fields(25).Value))

            If strPBigBox <> QBX Then
                DblMZ = DblMZ + DblMZ1

            End If

            '
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(23).Value)
            
            QBX = strPBigBox
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            rs.MoveNext
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
    Dim strXHCC As String       '�����ͳߴ�

    Dim DblTJZ  As String       '�����

    Dim order   As String
    
    order = Replace(ordertemp, "A", "")
    order = Replace$(order, "B", "")
    
    strXHCC = ""
    DblTJZ = 0
    'strdjbh1 = Mid(strdjbh, 2, Len(strdjbh) - 1)
    '    strSql = "SELECT COUNT(DISTINCT erpdata.dbo.f_getparent(b.���)) ����,c.�ߴ� " & _
    '             " FROM erpdata..tblStockMove a " & _
    '             " INNER JOIN erpdata..tblStockMovesub b ON a.���ݱ�� = b.���ݱ�� AND a.���=b.������� " & _
    '             " INNER JOIN erpdata..tblStockNumTree c On c.���=erpdata.dbo.f_getparent(b.���) " & _
    '             " WHERE a.���ݱ�� IN ('" & order & "')" & _
    '             " GROUP BY c.�ߴ�"
             
    strSql = "SELECT  COUNT(DISTINCT d.���) ����,d.�ߴ�  " & " FROM erpdata..tblStockSQfh  a  " & "  INNER JOIN erpdata..tblStockSQfhsub b ON a.���ݱ�� = b.���ݱ�� AND a.���=b.������� " & "   INNER JOIN erpdata..tblStockNumTree c On c.���=b.��� AND c.������ = 0 " & "   INNER JOIN erpdata..tblStockNumTree d On d.��� = c.�ϼ���� AND d.������ = 1 " & " WHERE a.���ݱ�� IN ('" & order & "','') GROUP BY d.�ߴ� "
             
    
    If RsNew.State = adStateOpen Then RsNew.Close
    RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

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
    
    wkst.Cells(8, 2) = Mid(dnnum, 1, Len(dnnum) - 1)
    
    wkst.Cells(8, 5) = ""
    
   ' wkst.Cells(7, 10).Select
    
    ' ���ɶ�ά��
    Dim strQrCodePath As String
    
    strNewFullPathNew = strNewFullPathNew & "\" & strExName & strExtsion
    strNewFullPathNew = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\37\excel" & "\" & strExName & strExtsion
    strQrCodePath = DirQrShare & "\" & strExName & ".JPG"
    strQrCodePath = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\37\jpg" & "\" & strExName & ".JPG"
    test.Visible = False

    test.QRmaker1.InputData = wkst.Cells(8, 2)
    test.QRmaker1.Refresh
    test.QRmaker1.CreateQrMetaFile hDC, strQrCodePath, 2
    Unload test

    'wkst.Pictures.Insert (App.Path + "\dn.bmp")
    wkst.Shapes.AddPicture _
    strQrCodePath, _
    True, True, 1100, 200, 400, 400
    
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

            If Err.number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub

            End If

        End If

    End If

    ' wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    ' wkbk.Saved = True
    
    'wkbk.SaveAs strNewFullPathNew, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPathNew
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
    MsgBox Err.DESCRIPTION, vbInformation, "��ʾ��"
    Exit Sub


End Sub

'DN��ַ��ͬ�ϲ�����
Public Sub ShippingPackinglistExportPrintExcel2(ByVal strExName As String, lxflag As Integer)

   Dim strDNList As String  'ȷ��DN��Χ
   Dim i As Integer
   Dim ShiptoName As String
   

        With lstLotID
            For i = 0 To .ListCount - 1
                If .Selected(i) = True Then
                    strDNList = strDNList & Trim$("" & .List(i)) & "','"
                    
                End If

            Next

        End With

 
 strDNList = Mid(strDNList, 1, Len(strDNList) - 3)
        
        
   With Fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 5
                If Trim(.text) <> ShiptoName Then
                    ShiptoName = .text
                    strExName = strDNList
                   Call SPPE2(strDNList, ShiptoName, strExName, lxflag)
                End If
            End If
            Next
        
    End With
    MsgBox "������ɣ�"
End Sub

'DN��ַ��ͬ�ϲ������Ӻ���
Public Sub SPPE2(strDNList As String, ShiptoName As String, strExName As String, lxflag As Integer)
    Dim strSql         As String

    Dim lngRows        As Long

    Dim rsQuery        As Excel.QueryTable

    Dim ExApp          As Excel.Application

    Dim wkbk           As New Workbook

    Dim wkst           As New Worksheet

    Dim i              As Long

    Dim j              As Long

    Dim IntCols        As Integer

    Dim strCols        As String

    Dim strFileName    As String

    Dim IntInertRow    As Integer, IntMaxDetailRow As Integer

    Dim DblNum         As Double

    Dim DblAmt         As Double  '�ܽ��

    Dim intBoxNum      As Integer '����

    Dim strPBigBox     As String  'ǰ���

    Dim strNBigBox     As String  '�����

    Dim IntBMegerRow   As Integer

    Dim IntEMegerRow   As Integer

    Dim DblJZ          As Double   '����

    Dim DblMZ          As Double   'ë��

    Dim DblJZ1         As Double   '����

    Dim DblMZ1         As Double   'ë��

    Dim DblJZ2         As Double   '����

    Dim DblMZ2         As Double   'ë��

    Dim intBegin       As Integer

    Dim strdjTmp       As String

    Dim SD             As String

    Dim SD1            As String

    Dim strTmp()       As String

    Dim strExtsion     As String '��׺��

    Dim strNewFullPath As String '��Excel�ļ�
    
    Dim strNewFullPathNew As String
    
    Dim order1 As String
    
    ' Dim strExName As String
    
    Dim ordertemp  As String
    Dim ordertemp1  As String

    Dim RsNew          As New ADODB.Recordset  '��¼����ĸ��������������������

    Dim rs             As New ADODB.Recordset

    Dim dnnum          As String

    Dim dnnum1         As String
    Dim DirFileShare1 As String
    
    order1 = ""
    dnnum = ""
    dnnum1 = ""
    ordertemp = ""
    ordertemp1 = ""
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
    

 
     If lxflag = 0 Then
        strNewFullPathNew = "C:\�ϲ�_spl"
        strFileName = DirShare & "\shipping_packing_list.xlsx" 'Ҫ�򿪵��ļ�
        strExtsion = ".xlsx"
        DirFileShare1 = "C:\�ϲ�_spl"
        strNewFullPath = DirFileShare1 & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
        If Dir("C:\�ϲ�_spl", vbDirectory) = "" Then '�ж��ļ����Ƿ����
            MkDir ("C:\�ϲ�_spl") '�����ļ��� msgbox ("�������")
            MsgBox ("�ļ����Ѵ�����·��Ϊ C:\�ϲ�_spl")
        Else
            'MsgBox ("�ļ�������")
        End If
  
    Else
        strNewFullPathNew = "C:\�ϲ�_spl2"
        strFileName = DirShare & "\shipping_packing_list_2.xlsx" 'Ҫ�򿪵��ļ�
        strExtsion = ".xlsx"
        DirFileShare1 = "C:\�ϲ�_spl2"
        strNewFullPath = DirFileShare1 & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
        If Dir("C:\�ϲ�_spl2", vbDirectory) = "" Then '�ж��ļ����Ƿ����
            MkDir ("C:\�ϲ�_spl2") '�����ļ��� msgbox ("�������")
            MsgBox ("�ļ����Ѵ�����·��Ϊ C:\�ϲ�_spl2")
        Else
            'MsgBox ("�ļ�������")
        End If
    End If
    
    
     
    'strExName = GetExcelName(Trim(Fra(1).Caption))

    '    strSql = "SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
    '                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
    '                 ",���,�Ϻ�,mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,����,ë��,�ߴ� " & _
    '                 " FROM Vw_InvShippedPLFor37 a  where ���ݱ�� in ('" & Ordertemp & "')  order by ���"

 strSql = "select aaa.* from (SELECT 0 AS  ѡ��,y.������ + b.���ݱ�� AS ���ݱ��,d.Delivery, CONVERT(VARCHAR(100), c.��������,23) AS ��������, ISNULL(dn_address_new, d.shiptoname) AS shiptoname, ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1 " & _
 " ,  ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2, ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,  ISNULL(x.city_new, d.city) AS city,  ISNULL(x.dn_st_new, d.State) AS state,  ISNULL(x.postal_code_new, d.postalcode) AS postalcode, " & _
 "  ISNULL(x.country_new, d.countrykey) AS countrykey,  ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone, d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo " & _
 " ,f.��� ,b.�Ϻ� ,d.MarketingPN,SUM(b.����) as sum,d.BatchNumber,d.DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no, d.customerPartNumber, ROUND(CAST(f.���� AS FLOAT) * 0.4, 2) ����, f.���� ë��, " & _
 " f.�ߴ� FROM erpdata..tblStockNumTree a INNER JOIN erpdata..tblStocksqfhsub b ON b.��� = a.��� INNER JOIN erpdata..tblStockSQfh c ON c.���ݱ�� = b.���ݱ�� AND c.��� = b.������� INNER JOIN erpdata..tblstock y " & _
 " ON y.�ⷿ���� = c.�ֿ���  INNER JOIN (SELECT a.BOX_ID,  SUBSTRING(a.KEY_VALUE,  1, CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job " & _
 " FROM erpdata .. tblErpInStockDetailInfo a  WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.��� = aa.qbox  INNER JOIN (SELECT dn.Delivery, dn.shiptoname, dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3 " & _
 " , dn.city, dn.State, dn.postalcode, dn.countrykey, dn.contactname,  dn.phone, dn.SalesDocument,  dn.PurchasingDocNo, dn.BatchNumber, SUM(dn.Quantity) AS Quantity, dn.customerPartNumber, dn.MarketingPN, dn.DATE_CODE " & _
 " FROM ERPBASE..tblCustomerShippingUp dn WHERE dn.Delivery in ('" & strDNList & "') GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1, dn.shiptostreet2, dn.shiptostreet3,dn.city, dn.State, dn.postalcode, dn.countrykey " & _
 " , dn.contactname,dn.phone, dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber, dn.customerPartNumber, dn.MarketingPN,dn.DATE_CODE) d  ON d.Delivery = a.DN AND d.BatchNumber = aa.job " & _
 " INNER JOIN erpdata..tblStockNumTree e ON e.��� = b.��� INNER JOIN erpdata..tblstocknumtree f  ON f.��� = e.�ϼ���� LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName " & _
"  WHERE a.DN IN ('" & strDNList & "') GROUP BY  b.���ݱ��,c.��������,ISNULL(dn_address_new, d.shiptoname) , ISNULL(x.ship_to_street1_new, d.shiptostreet1) ,ISNULL(x.ship_to_street2_new, d.shiptostreet2) " & _
 " ,ISNULL(x.ship_to_street3_new, d.shiptostreet3),ISNULL(x.city_new, d.city),ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode), ISNULL(x.country_new, d.countrykey) " & _
 " ,ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) , d.SalesDocument,d.PurchasingDocNo ,f.���,b.�Ϻ� ,d.MarketingPN,d.BatchNumber ,d.DATE_CODE " & _
 " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber, f.���� , f.�ߴ�,y.������,d.Delivery ) aaa where aaa.shiptoname = '" & ShiptoName & "' order by aaa.shiptoname, aaa.Delivery,aaa.��� "


    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
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
   
   If Not rs.EOF Then

        Do While Not rs.EOF
            If ordertemp <> Trim$(rs.Fields("���ݱ��")) Then
                ordertemp = Trim$(rs.Fields("���ݱ��"))
                ordertemp1 = ordertemp1 & ordertemp & "','"
            End If
            rs.MoveNext
        Loop

   End If
        'MsgBox "" & ordertemp1
        
     rs.MoveFirst    '���ݼ��ƶ�����һ��
    '    '---------------------------------------------------------------------
    If rs.RecordCount > 0 Then
            
        '        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
        '        ExApp.ActiveWindow.DisplayGridlines = False
        
        ' wkbk.ActiveSheet.Range(A3).Select
        
        DblNum = 0
        DblJZ = 0
        DblMZ = 0
        '

        '��ֵ��Excel�У���ͷ
        'wkst.Cells(8, 2) = Trim$("" & rs.Fields(2).Value)
        wkst.Cells(8, 17) = Trim$("" & rs.Fields(3).Value)
        'shipto  ����װ�䵥ʱ��ִ�� ����Sold to = Ship to 2019-12-5
        If lxflag = 1 Then
            wkst.Cells(10, 2) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(11, 2) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(12, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
            wkst.Cells(13, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
            wkst.Cells(14, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
            wkst.Cells(15, 2) = ""
        End If
        
        wkst.Cells(17, 2) = Trim$("" & rs.Fields(4).Value)
        wkst.Cells(18, 2) = Trim$("" & rs.Fields(5).Value)
        wkst.Cells(19, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
        wkst.Cells(20, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
        wkst.Cells(22, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
        If UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION (CAMARILLO)" Or UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION" Then
        
           wkst.Cells(23, 2) = "TAX ID: 95-2119684"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INTERCONNECT" Then
        
           wkst.Cells(23, 2) = "TAX ID: 82-5035949"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INCORPORATED(FEDERAL��" Then
        
            wkst.Cells(23, 2) = "TAX ID: 82-5035949"
        
        Else
           
            wkst.Cells(23, 2) = ""
        
        End If
        
     '   wkst.Cells(23, 2) = ""
        wkst.Cells(17, 17) = Trim$("" & rs.Fields(11).Value) 'To
        wkst.Cells(25, 3) = Trim$("" & rs.Fields(14).Value)
        wkst.Cells(25, 6) = Trim$("'" & rs.Fields(15).Value)
        
        lngRows = 28
        
        IntInertRow = rs.RecordCount * 2

        For i = 1 To IntInertRow - 1
            
            wkst.Rows(lngRows & ":" & lngRows).Select
            ExApp.Selection.Copy
            ExApp.Selection.Insert Shift:=xlDown
            wkst.Rows(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '�߿���
        Next i

        IntMaxDetailRow = rs.RecordCount
        
        '        ClsP.ShowProgress 50, "���ڵ���..."
        
        IntBMegerRow = 27
        IntEMegerRow = 30
        intBegin = 1

        Dim QBX As String

        For i = 0 To rs.RecordCount - 1

            '            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '���
            If dnnum1 <> Trim$("" & rs.Fields(2).Value) Then
                dnnum = Trim$("" & rs.Fields(2).Value) + "/" + dnnum
                dnnum1 = Trim$("" & rs.Fields(2).Value)

            End If
             
            strPBigBox = Trim$("" & rs.Fields(16).Value) '���

            'QBX = strPBigBox
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & rs.Fields(16).Value) '���
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
            
            If SD <> Trim$("" & rs.Fields(14).Value) Then
                SD = Trim$("" & rs.Fields(14).Value)
                SD1 = SD1 & SD & " "

            End If

            wkst.Cells(25, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & rs.Fields(19).Value)) / 1000 '������Ϊ��ǧΪ��λ
            DblNum = DblNum + Val(Trim$("" & rs.Fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(21).Value) 'datacode
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(22).Value) 'lotno

            If strPBigBox <> QBX Then
                
                If Trim$("" & rs.Fields(24).Value) = "" Or Trim$("" & rs.Fields(25).Value) = "" Or Trim$("" & rs.Fields(26).Value) = "" Then
                    
                    MsgBox "û��ά�������ߴ�", vbInformation, "��ʾ"
                    Exit Sub
                    
                End If
                
                wkst.Cells(lngRows, 14) = Trim$("" & rs.Fields(24).Value) '����
                wkst.Cells(lngRows, 15) = "KG"   '���ص�λ
                wkst.Cells(lngRows, 18) = "KG"   'ë�ص�λ
                wkst.Cells(lngRows, 19) = Trim$("" & rs.Fields(26).Value)   '�ߴ�
                wkst.Cells(lngRows, 17) = Trim$("" & rs.Fields(25).Value)   'ë��
            
            End If
           
            DblJZ1 = Val(Trim$("" & rs.Fields(24).Value))
            
            If strPBigBox <> QBX Then
                DblJZ = DblJZ1 + DblJZ

            End If

            DblMZ1 = Val(Trim$("" & rs.Fields(25).Value))

            If strPBigBox <> QBX Then
                DblMZ = DblMZ + DblMZ1

            End If

            '
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(23).Value)
            
            QBX = strPBigBox
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            rs.MoveNext
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
    Dim strXHCC As String       '�����ͳߴ�

    Dim DblTJZ  As String       '�����

   ' Dim order   As String
    ordertemp1 = Mid(ordertemp1, 1, Len(ordertemp1) - 3)
    order1 = Replace(ordertemp1, "A", "")
    order1 = Replace$(order1, "B", "")
    
    strXHCC = ""
    DblTJZ = 0
    'strdjbh1 = Mid(strdjbh, 2, Len(strdjbh) - 1)
    '    strSql = "SELECT COUNT(DISTINCT erpdata.dbo.f_getparent(b.���)) ����,c.�ߴ� " & _
    '             " FROM erpdata..tblStockMove a " & _
    '             " INNER JOIN erpdata..tblStockMovesub b ON a.���ݱ�� = b.���ݱ�� AND a.���=b.������� " & _
    '             " INNER JOIN erpdata..tblStockNumTree c On c.���=erpdata.dbo.f_getparent(b.���) " & _
    '             " WHERE a.���ݱ�� IN ('" & order & "')" & _
    '             " GROUP BY c.�ߴ�"
             
    strSql = "SELECT  COUNT(DISTINCT d.���) ����,d.�ߴ�  " & " FROM erpdata..tblStockSQfh  a  " & "  INNER JOIN erpdata..tblStockSQfhsub b ON a.���ݱ�� = b.���ݱ�� AND a.���=b.������� " & "   INNER JOIN erpdata..tblStockNumTree c On c.���=b.��� AND c.������ = 0 " & "   INNER JOIN erpdata..tblStockNumTree d On d.��� = c.�ϼ���� AND d.������ = 1 " & " WHERE a.���ݱ�� IN ('" & order1 & "') GROUP BY d.�ߴ� "
             
    
    If RsNew.State = adStateOpen Then RsNew.Close
    RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

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
    
    wkst.Cells(8, 2) = Mid(dnnum, 1, Len(dnnum) - 1)
    
    wkst.Cells(8, 5) = ""
    
   ' wkst.Cells(7, 10).Select
    
    ' ���ɶ�ά��
    Dim strQrCodePath As String
    
    strNewFullPathNew = strNewFullPathNew & "\" & strExName & strExtsion
    'strNewFullPathNew = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\37\excel" & "\" & strExName & strExtsion
    strQrCodePath = DirQrShare & "\" & strExName & ".JPG"
    strQrCodePath = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\37\jpg" & "\" & strExName & ".JPG"
    test.Visible = False

    test.QRmaker1.InputData = wkst.Cells(8, 2)
    test.QRmaker1.Refresh
    test.QRmaker1.CreateQrMetaFile hDC, strQrCodePath, 2
    Unload test

    'wkst.Pictures.Insert (App.Path + "\dn.bmp")
    wkst.Shapes.AddPicture _
    strQrCodePath, _
    True, True, 1100, 200, 400, 400
    
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

            If Err.number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub

            End If

        End If

    End If

    ' wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    ' wkbk.Saved = True
    
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
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
    MsgBox Err.DESCRIPTION, vbInformation, "��ʾ��ִ�����"
    Exit Sub
End Sub

'����DN�����
Public Sub ShippingPackinglistExportPrintExcel1(ByVal strExName As String, lxflag As Integer)

    Dim i              As Long

    Dim j              As Long
    Dim strDNList      As String
    

        

    
        With lstLotID

            For i = 0 To .ListCount - 1
        
                If .Selected(i) = True Then
                    'strDNList = strDNList & Trim$("" & .List(i)) & "','"
                    strDNList = Trim(.List(i))
                    strExName = strDNList
                    Call SPEPrintExcel1(strDNList, strExName, lxflag)
                    
                End If

            Next

        End With
    MsgBox "������ɣ�"
        
         'strDNList = Mid(strDNList, 1, Len(strDNList) - 3)
End Sub


'����DN�����ѭ�������Ӻ���
Private Function SPEPrintExcel1(strDNList As String, strExName As String, lxflag As Integer)
    Dim strSql         As String

    Dim lngRows        As Long

    Dim rsQuery        As Excel.QueryTable

    Dim ExApp          As Excel.Application

    Dim wkbk           As New Workbook

    Dim wkst           As New Worksheet

    Dim i              As Long

    Dim j              As Long

    Dim IntCols        As Integer

    Dim strCols        As String

    Dim strFileName    As String

    Dim IntInertRow    As Integer, IntMaxDetailRow As Integer

    Dim DblNum         As Double

    Dim DblAmt         As Double  '�ܽ��

    Dim intBoxNum      As Integer '����

    Dim strPBigBox     As String  'ǰ���

    Dim strNBigBox     As String  '�����

    Dim IntBMegerRow   As Integer

    Dim IntEMegerRow   As Integer

    Dim DblJZ          As Double   '����

    Dim DblMZ          As Double   'ë��

    Dim DblJZ1         As Double   '����

    Dim DblMZ1         As Double   'ë��

    Dim DblJZ2         As Double   '����

    Dim DblMZ2         As Double   'ë��

    Dim intBegin       As Integer

    Dim strdjTmp       As String

    Dim SD             As String

    Dim SD1            As String

    Dim strTmp()       As String

    Dim strExtsion     As String '��׺��

    Dim strNewFullPath As String '��Excel�ļ�
    
    Dim strNewFullPathNew As String
    
    'Dim strExName As String
    
    Dim ordertemp As String
    Dim ordertemp1 As String
    
    Dim order1 As String

    Dim RsNew          As New ADODB.Recordset  '��¼����ĸ��������������������

    Dim rs             As New ADODB.Recordset

    Dim dnnum          As String

    Dim dnnum1         As String
    Dim DirFileShare1 As String

    dnnum = ""
    dnnum1 = ""
    ordertemp = ""
    ordertemp1 = ""
    order1 = ""
    
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
       
    strNewFullPathNew = "C:\����"
    
    If lxflag = 0 Then
         strFileName = DirShare & "\shipping_packing_list.xlsx" 'Ҫ�򿪵��ļ�
         strExtsion = ".xlsx"
         DirFileShare1 = "C:\����"
         strNewFullPath = DirFileShare1 & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
         If Dir("C:\����", vbDirectory) = "" Then '�ж��ļ����Ƿ����
            MkDir ("C:\����") '�����ļ��� msgbox ("�������")
            MsgBox ("�ļ����Ѵ�����·��Ϊ C:\����")
         End If

   Else
         strFileName = DirShare & "\shipping_packing_list_2.xlsx" 'Ҫ�򿪵��ļ�
         strExtsion = ".xlsx"
         DirFileShare1 = "C:\����_2"
         strNewFullPath = DirFileShare1 & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
         If Dir("C:\����_2", vbDirectory) = "" Then '�ж��ļ����Ƿ����
            MkDir ("C:\����_2") '�����ļ��� msgbox ("�������")
            MsgBox ("�ļ����Ѵ�����·��Ϊ C:\����_2")
         End If
   End If
    'strExName = GetExcelName(Trim(Fra(1).Caption))

    '    strSql = "SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3," & _
    '                 "city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
    '                 ",���,�Ϻ�,mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,customerPartNumber,����,ë��,�ߴ� " & _
    '                 " FROM Vw_InvShippedPLFor37 a  where ���ݱ�� in ('" & Ordertemp & "')  order by ���"


 strSql = "SELECT 0 AS  ѡ��,y.������ + b.���ݱ�� AS ���ݱ��,d.Delivery, CONVERT(VARCHAR(100), c.��������,23) AS ��������, ISNULL(dn_address_new, d.shiptoname) AS shiptoname, ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1 " & _
 " ,  ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2, ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,  ISNULL(x.city_new, d.city) AS city,  ISNULL(x.dn_st_new, d.State) AS state,  ISNULL(x.postal_code_new, d.postalcode) AS postalcode, " & _
 "  ISNULL(x.country_new, d.countrykey) AS countrykey,  ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone, d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo " & _
 " ,f.��� ,b.�Ϻ� ,d.MarketingPN,SUM(b.����),d.BatchNumber,d.DATE_CODE,SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no, d.customerPartNumber, ROUND(CAST(f.���� AS FLOAT) * 0.4, 2) ����, f.���� ë��, " & _
 " f.�ߴ� FROM erpdata..tblStockNumTree a INNER JOIN erpdata..tblStocksqfhsub b ON b.��� = a.��� INNER JOIN erpdata..tblStockSQfh c ON c.���ݱ�� = b.���ݱ�� AND c.��� = b.������� INNER JOIN erpdata..tblstock y " & _
 " ON y.�ⷿ���� = c.�ֿ���  INNER JOIN (SELECT a.BOX_ID,  SUBSTRING(a.KEY_VALUE,  1, CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job " & _
 " FROM erpdata .. tblErpInStockDetailInfo a  WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.��� = aa.qbox  INNER JOIN (SELECT dn.Delivery, dn.shiptoname, dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3 " & _
 " , dn.city, dn.State, dn.postalcode, dn.countrykey, dn.contactname,  dn.phone, dn.SalesDocument,  dn.PurchasingDocNo, dn.BatchNumber, SUM(dn.Quantity) AS Quantity, dn.customerPartNumber, dn.MarketingPN, dn.DATE_CODE " & _
 " FROM ERPBASE..tblCustomerShippingUp dn WHERE dn.Delivery in ('" & strDNList & "') GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1, dn.shiptostreet2, dn.shiptostreet3,dn.city, dn.State, dn.postalcode, dn.countrykey " & _
 " , dn.contactname,dn.phone, dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber, dn.customerPartNumber, dn.MarketingPN,dn.DATE_CODE) d  ON d.Delivery = a.DN AND d.BatchNumber = aa.job " & _
 " INNER JOIN erpdata..tblStockNumTree e ON e.��� = b.��� INNER JOIN erpdata..tblstocknumtree f  ON f.��� = e.�ϼ���� LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName " & _
 "  WHERE a.DN IN ('" & strDNList & "')GROUP BY  b.���ݱ��,c.��������,ISNULL(dn_address_new, d.shiptoname) , ISNULL(x.ship_to_street1_new, d.shiptostreet1) ,ISNULL(x.ship_to_street2_new, d.shiptostreet2) " & _
 " ,ISNULL(x.ship_to_street3_new, d.shiptostreet3),ISNULL(x.city_new, d.city),ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode), ISNULL(x.country_new, d.countrykey) " & _
 " ,ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) , d.SalesDocument,d.PurchasingDocNo ,f.���,b.�Ϻ� ,d.MarketingPN,d.BatchNumber ,d.DATE_CODE " & _
 " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber, f.���� , f.�ߴ�,y.������,d.Delivery order by Delivery,���"

    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
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
    
    'strExtsion = Mid$(StrFileName, InStrRev(StrFileName, "."))      '��ȡ��׺��

    '�Ҹ�DN��ȫ�����ݱ��
       If Not rs.EOF Then

        Do While Not rs.EOF
            If ordertemp <> Trim$(rs.Fields("���ݱ��")) Then
                ordertemp = Trim$(rs.Fields("���ݱ��"))
                ordertemp1 = ordertemp1 & ordertemp & "','"
            End If
            rs.MoveNext
        Loop

       End If
    
    '�����ǰDNû�����ݣ���ִ����һ��DN����������������������
    If rs.RecordCount = 0 Then
        Exit Function
    End If
        
     rs.MoveFirst    '���ݼ��ƶ�����һ��
     
    If rs.RecordCount > 0 Then
        '        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
        '        ExApp.ActiveWindow.DisplayGridlines = False
        
        ' wkbk.ActiveSheet.Range(A3).Select
        
       
        
        DblNum = 0
        DblJZ = 0
        DblMZ = 0
        '

        '��ֵ��Excel�У���ͷ
        'wkst.Cells(8, 2) = Trim$("" & rs.Fields(2).Value)
        wkst.Cells(8, 17) = Trim$("" & rs.Fields(3).Value)
        'shipto  ����װ�䵥ʱ��ִ�� ����Sold to = Ship to 2019-12-5
        If lxflag = 1 Then
            wkst.Cells(10, 2) = Trim$("" & rs.Fields(4).Value)
            wkst.Cells(11, 2) = Trim$("" & rs.Fields(5).Value)
            wkst.Cells(12, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
            wkst.Cells(13, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
            wkst.Cells(14, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
            wkst.Cells(15, 2) = ""
        End If
        'soldto
        wkst.Cells(17, 2) = Trim$("" & rs.Fields(4).Value)
        wkst.Cells(18, 2) = Trim$("" & rs.Fields(5).Value)
        wkst.Cells(19, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
        wkst.Cells(20, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
        wkst.Cells(22, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)


        If UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION (CAMARILLO)" Or UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION" Then
        
           wkst.Cells(23, 2) = "TAX ID: 95-2119684"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INTERCONNECT" Then
        
           wkst.Cells(23, 2) = "TAX ID: 82-5035949"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INCORPORATED(FEDERAL��" Then
        
            wkst.Cells(23, 2) = "TAX ID: 82-5035949"
        
        Else
           
        wkst.Cells(23, 2) = ""
        
        End If
        wkst.Cells(17, 17) = Trim$("" & rs.Fields(11).Value) 'To
        wkst.Cells(25, 3) = Trim$("" & rs.Fields(14).Value)
        wkst.Cells(25, 6) = Trim$("'" & rs.Fields(15).Value)
        
        lngRows = 28
        
        IntInertRow = rs.RecordCount * 2

        For i = 1 To IntInertRow - 1
            wkst.Rows(lngRows & ":" & lngRows).Select
            ExApp.Selection.Copy
            ExApp.Selection.Insert Shift:=xlDown
            wkst.Rows(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '�߿���
        Next i

        IntMaxDetailRow = rs.RecordCount
        
        '        ClsP.ShowProgress 50, "���ڵ���..."
        
        IntBMegerRow = 27
        IntEMegerRow = 30
        intBegin = 1

        Dim QBX As String

        For i = 0 To rs.RecordCount - 1

            '            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '���
            If dnnum1 <> Trim$("" & rs.Fields(2).Value) Then
                dnnum = Trim$("" & rs.Fields(2).Value) + "/" + dnnum
                dnnum1 = Trim$("" & rs.Fields(2).Value)

            End If
             
            strPBigBox = Trim$("" & rs.Fields(16).Value) '���

            'QBX = strPBigBox
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & rs.Fields(16).Value) '���
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
            
            If SD <> Trim$("" & rs.Fields(14).Value) Then
                SD = Trim$("" & rs.Fields(14).Value)
                SD1 = SD1 & SD & " "

            End If

            wkst.Cells(25, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & rs.Fields(19).Value)) / 1000 '������Ϊ��ǧΪ��λ
            DblNum = DblNum + Val(Trim$("" & rs.Fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(21).Value) 'datacode
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(22).Value) 'lotno

            If strPBigBox <> QBX Then
                wkst.Cells(lngRows, 14) = Trim$("" & rs.Fields(24).Value) '����
                wkst.Cells(lngRows, 15) = "KG"   '���ص�λ
                wkst.Cells(lngRows, 18) = "KG"   'ë�ص�λ
                wkst.Cells(lngRows, 19) = Trim$("" & rs.Fields(26).Value)   '�ߴ�
                wkst.Cells(lngRows, 17) = Trim$("" & rs.Fields(25).Value)   'ë��
            
            End If
           
            DblJZ1 = Val(Trim$("" & rs.Fields(24).Value))
            
            If strPBigBox <> QBX Then
                DblJZ = DblJZ1 + DblJZ

            End If

            DblMZ1 = Val(Trim$("" & rs.Fields(25).Value))

            If strPBigBox <> QBX Then
                DblMZ = DblMZ + DblMZ1

            End If

            '
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(23).Value)
            
            QBX = strPBigBox
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            rs.MoveNext
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
        Exit Function

    End If

    '��ѯ��ųߴ磬���������
    Dim strXHCC As String       '�����ͳߴ�

    Dim DblTJZ  As String       '�����

   ' Dim order   As String
    ordertemp1 = Mid(ordertemp1, 1, Len(ordertemp1) - 3)
    order1 = Replace(ordertemp1, "A", "")
    order1 = Replace$(order1, "B", "")
    
    strXHCC = ""
    DblTJZ = 0
    'strdjbh1 = Mid(strdjbh, 2, Len(strdjbh) - 1)
    '    strSql = "SELECT COUNT(DISTINCT erpdata.dbo.f_getparent(b.���)) ����,c.�ߴ� " & _
    '             " FROM erpdata..tblStockMove a " & _
    '             " INNER JOIN erpdata..tblStockMovesub b ON a.���ݱ�� = b.���ݱ�� AND a.���=b.������� " & _
    '             " INNER JOIN erpdata..tblStockNumTree c On c.���=erpdata.dbo.f_getparent(b.���) " & _
    '             " WHERE a.���ݱ�� IN ('" & order & "')" & _
    '             " GROUP BY c.�ߴ�"
             
    strSql = "SELECT  COUNT(DISTINCT d.���) ����,d.�ߴ�  " & " FROM erpdata..tblStockSQfh  a  " & "  INNER JOIN erpdata..tblStockSQfhsub b ON a.���ݱ�� = b.���ݱ�� AND a.���=b.������� " & "   INNER JOIN erpdata..tblStockNumTree c On c.���=b.��� AND c.������ = 0 " & "   INNER JOIN erpdata..tblStockNumTree d On d.��� = c.�ϼ���� AND d.������ = 1 " & " WHERE a.���ݱ�� IN ('" & order1 & "','') GROUP BY d.�ߴ� "
             
    
    If RsNew.State = adStateOpen Then RsNew.Close
    RsNew.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText

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
    
    wkst.Cells(8, 2) = Mid(dnnum, 1, Len(dnnum) - 1)
    
    wkst.Cells(8, 5) = ""
    
   ' wkst.Cells(7, 10).Select
    
    ' ���ɶ�ά��
    Dim strQrCodePath As String
    
    strNewFullPathNew = strNewFullPathNew & "\" & strExName & strExtsion
    'strNewFullPathNew = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\37\excel" & "\" & strExName & strExtsion
    strQrCodePath = DirQrShare & "\" & strExName & ".JPG"
    strQrCodePath = "\\10.160.1.84\public\FileServer\35.�г�������ά��ϵͳ\37\jpg" & "\" & strExName & ".JPG"
    test.Visible = False

    test.QRmaker1.InputData = wkst.Cells(8, 2)
    test.QRmaker1.Refresh
    test.QRmaker1.CreateQrMetaFile hDC, strQrCodePath, 2
    Unload test

    'wkst.Pictures.Insert (App.Path + "\dn.bmp")
    wkst.Shapes.AddPicture _
    strQrCodePath, _
    True, True, 1100, 200, 400, 400
    
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
            Exit Function
        Else

            On Error Resume Next

            Kill strNewFullPath

            If Err.number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Function

            End If

        End If

    End If

    ' wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    ' wkbk.Saved = True
    
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
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

    Exit Function
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
    MsgBox Err.DESCRIPTION, vbInformation, "��ʾ��ִ�����"
    Exit Function


End Function
'shipping invoice
Public Sub ShippingInvoiceExportPrintExcel(ByVal ordertemp As String, ByVal strExName As String)
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
    Dim rs               As New ADODB.Recordset
    Dim dnnum            As String
    Dim dnnum1            As String
    Dim jine1 As Double
    Dim jine2 As Double
     
    dnnum = ""
    dnnum1 = ""
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
    
       Dim strDNList As String
        With lstLotID

            For i = 0 To .ListCount - 1
        
                If .Selected(i) = True Then
                    strDNList = strDNList & Trim$("" & .List(i)) & "','"

                End If
                
            Next

        End With

        
          strDNList = Mid(strDNList, 1, Len(strDNList) - 3)

                    
'    strSql = " SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
'                 " ,city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
'                 " ,���,�Ϻ�,mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,����,AMount,customerPartNumber " & _
'                 "  FROM Vw_InvShippedInvoiceFor37 a  where ���ݱ�� in ('" & Ordertemp & "')  order by ���  "


' strSql = "   select * from ( SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
'                 " ,city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
'                 " ,���,�Ϻ�,replace(mpn_desc,'.P2','') as mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,����,AMount,customerPartNumber,�ӹ��ѽ��,�͹��Ͻ��  " & _
'                 "  FROM Vw_InvShippedInvoiceFor37_NEW a  where ���ݱ�� in ('" & Ordertemp & "')    " & _
'                 "union all " & _
'                 " SELECT 0 ѡ��,���ݱ��,delivery,��������,shiptoname,shiptostreet1,shiptostreet2,shiptostreet3" & _
'                 " ,city,State,postalcode,countrykey,contactname,phone,SalesDocument,PurchasingDocNo" & _
'                 " ,���,�Ϻ�,replace(mpn_desc,'.P2','') as mpn_desc,����,batchnumber,DATE_CODE,HTlot_no,����,AMount,customerPartNumber ,�ӹ��ѽ��,�͹��Ͻ��  " & _
'                 "  FROM Vw_InvShippedInvoiceFor37 a  where ���ݱ�� in ('" & Ordertemp & "') )  x order by x.���  "
'
    
'    strSql = "SELECT 0 ѡ��, h.������+a.���ݱ�� ���ݱ��,c.delivery,dbo.usp_date(a.��������) ��������, ISNULL(dn_address_new, c.shiptoname) AS  shiptoname,ISNULL(x.ship_to_street1_new ,c.shiptostreet1)  AS shiptostreet1  " & _
'",ISNULL(x.ship_to_street2_new,c.shiptostreet2) AS shiptostreet2 ,ISNULL(x.ship_to_street3_new,c.shiptostreet3) AS shiptostreet3 ,ISNULL(x.city_new ,c.city) AS city,ISNULL(x.dn_st_new, c.State) AS state,ISNULL(x.postal_code_new,c.postalcode  ) AS  postalcode " & _
'",ISNULL(x.country_new,c.countrykey) AS countrykey,ISNULL(x.contact_new,c.contactname) AS contactname ,ISNULL(x.phone_new,c.phone ) AS phone      " & _
'",c.SalesDocument,'''' + c.PurchasingDocNo AS PurchasingDocNo,RTRIM(erpdata.dbo.f_getparent(b.���)) ���,b.�Ϻ�,CASE WHEN RTRIM(gg.MPN_DESC)='UCLAMP0541Z.TFT' THEN RTRIM(gg.MPN_DESC)+'.P2' " & _
'"ELSE REPLACE(REPLACE(gg.MPN_DESC,'.P2',''),'.P3','') END AS mpn_desc,SUM(b.����) ����,c.batchnumber,hh.CREATE_DATE DATE_CODE " & _
'",SUBSTRING(aa.qbox ,2,CHARINDEX('-R',aa.qbox)-2)  HTlot_no ,ISNULL(ISNULL( BB.��˰���� / AB.��Ʒ��,0)  + ( cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0) AS ���� " & _
'",ROUND( SUM(b.����) * ISNULL(ISNULL( BB.��˰���� / AB.��Ʒ��,0)  +  (cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0),2) AS AMount,c.customerPartNumber    " & _
'",ROUND( SUM(b.����) * ISNULL(ISNULL( BB.��˰���� / AB.��Ʒ��,0)  +  (cb.WAFER_PRICE/db.PASSBINCOUNT + cb.DIE_PRICE),0),2) -  CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( BB.��˰���� / AB.��Ʒ��,0)) AS �ӹ��ѽ�� " & _
'", CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( BB.��˰���� / AB.��Ʒ��,0)) AS �͹��Ͻ�� FROM erpdata..tblStockSQfh a           " & _
'"INNER JOIN erpdata..tblStocksqfhsub b ON a.���ݱ�� = b.���ݱ�� AND a.���=b.������� INNER JOIN erpdata..tblStockNumTree g ON g.���=b.��� " & _
'"INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE,1,CHARINDEX('|',a.KEY_VALUE)-1) AS qbox , SUBSTRING(a.KEY_VALUE,CHARINDEX('|',a.KEY_VALUE)+1,10) AS job " & _
'" FROM erpdata..tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND  a.KEY_VALUE LIKE '%SS%|%')  aa ON g.��� = aa.qbox  " & _
'"INNER JOIN (SELECT dn.Delivery,dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode "
'
'
'     strSql = strSql & ",dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber FROM tblCustomerShippingUp dn GROUP BY dn.Delivery,dn.shiptoname,dn.shiptostreet1 " & _
'",dn.shiptostreet2,dn.shiptostreet3,dn.city,dn.State,dn.postalcode,dn.countrykey,dn.contactname,dn.phone,dn.SalesDocument,dn.PurchasingDocNo,dn.BatchNumber,dn.customerPartNumber)  c ON c.Delivery = g.DN  AND c.BatchNumber =  aa.job " & _
'"INNER JOIN dbo.tblstock h ON CONVERT(NVARCHAR(4),h.�ⷿ����) = CONVERT(NVARCHAR(4),a.�ֿ���)    INNER JOIN ERPBASE..tblmappingData ff ON ff.SUBSTRATEID = b.���̿���� " & _
'"INNER JOIN ERPBASE..tblCustomerOI gg ON CONVERT(VARCHAR(30), gg.ID) = ff.FILENAME AND gg.SOURCE_BATCH_ID = ff.LOTID AND gg.CUSTOMERSHORTNAME = '37' " & _
'"INNER JOIN erpbase..weight37 hh ON hh.WAFERID = REPLACE(b.���̿����,'+','') INNER JOIN erpdata..tblErpInStockMainData jj ON jj.BOX_ID = aa.BOX_ID " & _
'"LEFT JOIN erpbase..tbltoinrec_wafer  AB ON ab.���� = ff.LOTID AND AB.��ԲID = REPLACE(B.���̿����,'+','') LEFT JOIN erpbase..tbltorec_wafer  ww ON  ww.���� = ab.���� AND  ww.��ԲID = ab.��ԲID  " & _
'"LEFT JOIN ERPBASE..TblToInsub BB ON BB.��ⵥ��� = AB.��ⵥ��� AND BB.�������� = AB.���� AND ww.��������� = bb.��������� AND bb.��˰���� IS NOT NULL " & _
'"LEFT JOIN erptemp..tblBB_CSRPO cb ON cb.PO_NUM = gg.PO_NUM AND cb.FAB_DEVICE = gg.MPN_DESC LEFT JOIN ERPBASE..tblmappingData db ON db.SUBSTRATEID = REPLACE(B.���̿����,'+','')  " & _
'"LEFT JOIN  erpdata..tblSalerec e ON  e.���ݱ�� = a.���ݱ��       AND a.��� = e.�������  AND e.С��� = b.���    " & _
'"LEFT JOIN erptemp..dn_address x ON dn_address = c.ShipToName WHERE a.�ͻ�����='37' and c.Delivery in ('" & strDNList & "') AND a.�������� >= CONVERT(VARCHAR(100),GETDATE()- 8,23) AND  a.���ݱ�� LIKE 'F%' AND a.��Ʒ���� >0  " & _
'"GROUP BY gg.PO_NUM,h.������,a.���ݱ��,c.delivery,dbo.usp_date(a.��������),c.shiptoname,c.shiptostreet1,c.shiptostreet2         " & _
'",c.shiptostreet3,c.city,c.State,c.postalcode,c.countrykey,c.contactname,c.phone,c.SalesDocument,c.PurchasingDocNo ,erpdata.dbo.f_getparent(b.���),b.�Ϻ�,gg.MPN_DESC,c.batchnumber,hh.CREATE_DATE      " & _
'",SUBSTRING(aa.qbox ,2,CHARINDEX('-R',aa.qbox)-2),e.���۵����,c.customerPartNumber ,ISNULL( BB.��˰���� / AB.��Ʒ��,0) ,  cb.WAFER_PRICE,db.PASSBINCOUNT , cb.DIE_PRICE,dn_address_new " & _
'",x.ship_to_street1_new,x.ship_to_street2_new,x.ship_to_street3_new,x.city_new,x.dn_st_new,x.postal_code_new,x.country_new,x.contact_new ,x.phone_new  order by Delivery "


    strSql = " SELECT 0 AS ѡ��,y.������ + b.���ݱ�� AS ���ݱ��,a.DN, CONVERT(VARCHAR(100), c.��������,23) AS ��������, ISNULL(dn_address_new, d.shiptoname) AS shiptoname " & _
 ",ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1,ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2 , ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, d.city) AS city, " & _
 " ISNULL(x.dn_st_new, d.State) AS state, ISNULL(x.postal_code_new, d.postalcode) AS postalcode,ISNULL(x.country_new, d.countrykey) AS countrykey, ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone,d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo, f.���, " & _
 " b.�Ϻ�,d.MarketingPN,SUM(b.����), d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no ,ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  + ( dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0) AS ���� " & _
 " ,ROUND( SUM(b.����) * ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  +  (dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0),2) AS AMount ,d.customerPartNumber ,ROUND( SUM(b.����) * ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  +  (dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0),2) " & _
 "  -  CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( cd.��˰���� / cc.��Ʒ��,0)) AS �ӹ��ѽ�� , CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( cd.��˰���� / cc.��Ʒ��,0)) AS �͹��Ͻ�� " & _
"   FROM erpdata .. tblStockNumTree a INNER JOIN erpdata .. tblStocksqfhsub b ON b.��� = a.��� INNER JOIN erpdata .. tblStockSQfh c  ON c.���ݱ�� = b.���ݱ�� AND c.��� = b.������� INNER JOIN erpdata..tblstock y ON y.�ⷿ���� = c.�ֿ��� " & _
"  INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE, 1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job  FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.��� = aa.qbox  INNER JOIN (SELECT dn.Delivery, " & _
 "  dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State,dn.postalcode, dn.countrykey, dn.contactname,dn.phone, dn.SalesDocument,  dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE  FROM ERPBASE .. tblCustomerShippingUp dn " & _
 "   WHERE dn.Delivery IN ('" & strDNList & "')  GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State, dn.postalcode,  dn.countrykey,dn.contactname,dn.phone, dn.SalesDocument, dn.PurchasingDocNo, dn.BatchNumber, dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE) d " & _
 "   ON d.Delivery = a.DN  AND d.BatchNumber = aa.job INNER JOIN erpdata .. tblStockNumTree e  ON e.��� = b.��� INNER JOIN erpdata .. tblstocknumtree f  ON f.��� = e.�ϼ����  INNER JOIN ERPBASE..tblmappingData qq ON qq.SUBSTRATEID = b.���̿���� AND qq.LOTID = b.������ LEFT JOIN ERPBASE..tblCustomerOI bb " & _
 " ON CONVERT(VARCHAR(100), bb.ID) = qq.FILENAME AND bb.SOURCE_BATCH_ID = qq.LOTID  LEFT JOIN ERPBASE..tblToRec_Wafer cc ON cc.��ԲID = REPLACE(qq.SUBSTRATEID,'+','')   AND cc.���� = qq.LOTID  LEFT JOIN ERPBASE..tblToRecEntry cd ON cd.��������� = cc.��������� AND cd.�������� = cc.���� LEFT JOIN erptemp..tblBB_CSRPO dd " & _
"  ON dd.PO_NUM = bb.PO_NUM AND dd.FAB_DEVICE = bb.MPN_DESC LEFT JOIN  erpdata..tblSalerec j ON j.���ݱ�� = c.���ݱ�� AND j.������� = b.������� AND j.С��� = b.��� LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName  WHERE a.DN IN ('" & strDNList & "')  GROUP BY b.���ݱ��, c.��������, ISNULL(dn_address_new, d.shiptoname), " & _
"  ISNULL(x.ship_to_street1_new, d.shiptostreet1), ISNULL(x.ship_to_street2_new, d.shiptostreet2),ISNULL(x.ship_to_street3_new, d.shiptostreet3), ISNULL(x.city_new, d.city),  ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode),ISNULL(x.country_new, d.countrykey),ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone), " & _
" d.SalesDocument, d.PurchasingDocNo,f.���, b.�Ϻ�,d.MarketingPN,d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber,   a.DN , J.���۵����, bb.PO_NUM, cd.��˰����, cc.��Ʒ��, dd.WAFER_PRICE, qq.PASSBINCOUNT, dd.die_price, y.������ order by DN,��� "
  
     rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     
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
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
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
        'wkst.Cells(13, 10) = Trim$("" & rs.Fields(2).Value)
        wkst.Cells(15, 10) = Trim$("" & rs.Fields(3).Value)
        wkst.Cells(18, 10) = Trim$("" & rs.Fields(11).Value) 'To
        
        wkst.Cells(13, 2) = Trim$("" & rs.Fields(4).Value)
        wkst.Cells(14, 2) = Trim$("" & rs.Fields(5).Value)
        wkst.Cells(15, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
        wkst.Cells(16, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
        
        wkst.Cells(18, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
        wkst.Cells(19, 2) = ""

        'wkst.Cells(23, 3) = Trim$("" & Rs.fields(14).Value)
        wkst.Cells(23, 5) = Trim$("'" & rs.Fields(15).Value)
        
        lngRows = 27
        
        IntInertRow = rs.RecordCount * 2
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
           wkst.Range(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '�߿���
        Next i
        IntMaxDetailRow = rs.RecordCount
        
'        ClsP.ShowProgress 50, "���ڵ���..."
        
        IntBMegerRow = 26
        IntEMegerRow = 29
        intBegin = 1
        Dim QBX As String
        
        For i = 0 To rs.RecordCount - 1
'            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '���

             If dnnum1 <> Trim$("" & rs.Fields(2).Value) And InStr(dnnum, Trim$("" & rs.Fields(2).Value)) = 0 Then
                dnnum = Trim$("" & rs.Fields(2).Value) + "/" + dnnum
                dnnum1 = Trim$("" & rs.Fields(2).Value)
             End If
             

            
            strPBigBox = Trim$("" & rs.Fields(16).Value) '���
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & rs.Fields(16).Value) '���
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
              If SD <> Trim$("" & rs.Fields(14).Value) Then
             SD = Trim$("" & rs.Fields(14).Value)
             SD1 = SD1 & SD & " "
             End If
            wkst.Cells(23, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & rs.Fields(19).Value)) / 1000 '������Ϊ����1000��ֵ
            DblNum = DblNum + Val(Trim$("" & rs.Fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(21).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(22).Value)
            wkst.Cells(lngRows, 13) = "US$"
            wkst.Cells(lngRows, 14) = Val(Trim$("" & rs.Fields(23).Value)) * 1000 '����Ϊ����1000��ֵ
            wkst.Cells(lngRows, 15) = "US$"
            wkst.Cells(lngRows, 16) = Trim$("" & rs.Fields(24).Value)
            DblAmt = DblAmt + Val(Trim$("" & rs.Fields(24).Value))
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(25).Value)
            
            jine1 = jine1 + Val(Trim$("" & rs.Fields(26).Value))
            
            jine2 = jine2 + Val(Trim$("" & rs.Fields(27).Value))
            
            
            
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            rs.MoveNext
        Next
        
        wkst.Cells(13, 10) = Mid(dnnum, 1, Len(dnnum) - 1)
        
        '�������
        wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '����
        wkst.Cells(lngRows + 1, 9) = "KPCS" '��λ
        wkst.Cells(lngRows + 1, 16) = DblAmt
        wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)
        
        wkst.Cells(lngRows + 8, 1) = "Process Amount��US$ " + Str(jine1)
        wkst.Cells(lngRows + 9, 1) = "Wafer Amount��US$ " + Str(jine2)

        
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
            If Err.number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
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
    MsgBox Err.DESCRIPTION, vbInformation, "��ʾ��"
    Exit Sub
End Sub

'shipping invoice
Public Sub ShippingInvoiceExportPrintExcel2()

    Dim strDNList As String
    Dim ShiptoName As String
    Dim strExName As String
    Dim i As Integer

    If Dir("C:\�ϲ�_voice", vbDirectory) = "" Then '�ж��ļ����Ƿ����
        MkDir ("C:\�ϲ�_voice") '�����ļ��� msgbox ("�������")
        MsgBox ("�ļ����Ѵ�����·��Ϊ C:\�ϲ�_voice")
    Else
        'MsgBox ("�ļ�������")
    End If
    
        With lstLotID

            For i = 0 To .ListCount - 1
        
                If .Selected(i) = True Then
                    strDNList = strDNList & Trim$("" & .List(i)) & "','"

                End If
                
            Next

        End With
        strDNList = Mid(strDNList, 1, Len(strDNList) - 3)
        
      With Fps(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1

            If .Value = 1 Then
                .Col = 5
                If Trim(.text) <> ShiptoName Then
                    ShiptoName = Trim(.text)
                    strExName = strDNList
                    Call SPLSVoice(strDNList, strExName, ShiptoName)
                End If
            End If
            Next
        
    End With
    MsgBox "������ɣ�"
          

End Sub

Public Function SPLSVoice(strDNList As String, strExName As String, ShiptoName As String)
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
    Dim rs               As New ADODB.Recordset
    Dim dnnum            As String
    Dim dnnum1            As String
    Dim jine1 As Double
    Dim jine2 As Double
     
    dnnum = ""
    dnnum1 = ""
    strPBigBox = ""
    strNBigBox = ""
    strdjTmp = ""
    intBoxNum = 1
    
    strFileName = DirShare & "\shipping_invoice.xlsx" 'Ҫ�򿪵��ļ�

'    strsql = "select * from( SELECT 0 AS ѡ��,y.������ + b.���ݱ�� AS ���ݱ��,a.DN, CONVERT(VARCHAR(100), c.��������,23) AS ��������, ISNULL(dn_address_new, d.shiptoname) AS shiptoname " & _
' ",ISNULL(x.ship_to_street1_new, d.shiptostreet1) AS shiptostreet1,ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2 , ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, d.city) AS city, " & _
' " ISNULL(x.dn_st_new, d.State) AS state, ISNULL(x.postal_code_new, d.postalcode) AS postalcode,ISNULL(x.country_new, d.countrykey) AS countrykey, ISNULL(x.contact_new, d.contactname) AS contactname, ISNULL(x.phone_new, d.phone) AS phone,d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo, f.���, " & _
' " b.�Ϻ�,d.MarketingPN,SUM(b.����) as ����, d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no ,ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  + ( dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0) AS ���� " & _
' " ,ROUND( SUM(b.����) * ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  +  (dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0),2) AS AMount ,d.customerPartNumber ,ROUND( SUM(b.����) * ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  +  (dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0),2) " & _
' "  -  CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( cd.��˰���� / cc.��Ʒ��,0)) AS �ӹ��ѽ�� , CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( cd.��˰���� / cc.��Ʒ��,0)) AS �͹��Ͻ�� " & _
'"   FROM erpdata .. tblStockNumTree a INNER JOIN erpdata .. tblStocksqfhsub b ON b.��� = a.��� INNER JOIN erpdata .. tblStockSQfh c  ON c.���ݱ�� = b.���ݱ�� AND c.��� = b.������� INNER JOIN erpdata..tblstock y ON y.�ⷿ���� = c.�ֿ��� " & _
'"  INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE, 1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox, SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job  FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.��� = aa.qbox  INNER JOIN (SELECT dn.Delivery, " & _
' "  dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State,dn.postalcode, dn.countrykey, dn.contactname,dn.phone, dn.SalesDocument,  dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE  FROM ERPBASE .. tblCustomerShippingUp dn " & _
' "   WHERE dn.Delivery IN ('" & strDNList & "')  GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State, dn.postalcode,  dn.countrykey,dn.contactname,dn.phone, dn.SalesDocument, dn.PurchasingDocNo, dn.BatchNumber, dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE) d " & _
' "   ON d.Delivery = a.DN  AND d.BatchNumber = aa.job INNER JOIN erpdata .. tblStockNumTree e  ON e.��� = b.��� INNER JOIN erpdata .. tblstocknumtree f  ON f.��� = e.�ϼ����  INNER JOIN ERPBASE..tblmappingData qq ON qq.SUBSTRATEID = b.���̿���� AND qq.LOTID = b.������ LEFT JOIN ERPBASE..tblCustomerOI bb " & _
' " ON CONVERT(VARCHAR(100), bb.ID) = qq.FILENAME AND bb.SOURCE_BATCH_ID = qq.LOTID  LEFT JOIN ERPBASE..tblToRec_Wafer cc ON cc.��ԲID = REPLACE(qq.SUBSTRATEID,'+','')   AND cc.���� = qq.LOTID  LEFT JOIN ERPBASE..tblToRecEntry cd ON cd.��������� = cc.��������� AND cd.�������� = cc.���� LEFT JOIN erptemp..tblBB_CSRPO dd " & _
'"  ON dd.PO_NUM = bb.PO_NUM AND dd.FAB_DEVICE = bb.MPN_DESC LEFT JOIN  erpdata..tblSalerec j ON j.���ݱ�� = c.���ݱ�� AND j.������� = b.������� AND j.С��� = b.��� LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName  WHERE a.DN IN ('" & strDNList & "')  GROUP BY b.���ݱ��, c.��������, ISNULL(dn_address_new, d.shiptoname), " & _
'"  ISNULL(x.ship_to_street1_new, d.shiptostreet1), ISNULL(x.ship_to_street2_new, d.shiptostreet2),ISNULL(x.ship_to_street3_new, d.shiptostreet3), ISNULL(x.city_new, d.city),  ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode),ISNULL(x.country_new, d.countrykey),ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone), " & _
'" d.SalesDocument, d.PurchasingDocNo,f.���, b.�Ϻ�,d.MarketingPN,d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber,   a.DN , J.���۵����, bb.PO_NUM, cd.��˰����, cc.��Ʒ��, dd.WAFER_PRICE, qq.PASSBINCOUNT, dd.die_price, y.������  )aaa where aaa.shiptoname = '" & ShiptoName & "' order by aaa.shiptoname,aaa.DN,aaa.���"
'
'
   strSql = "select * from(SELECT 0 AS ѡ��,y.������ + b.���ݱ�� AS ���ݱ��,a.DN, CONVERT(VARCHAR(100), c.��������,23) AS ��������, ISNULL(dn_address_new, d.shiptoname) AS shiptoname ,ISNULL(x.ship_to_street1_new " & _
         " , d.shiptostreet1) AS shiptostreet1,ISNULL(x.ship_to_street2_new, d.shiptostreet2) AS shiptostreet2 , ISNULL(x.ship_to_street3_new, d.shiptostreet3) AS shiptostreet3,ISNULL(x.city_new, d.city) AS city " & _
         " ,  ISNULL(x.dn_st_new, d.State) AS state, ISNULL(x.postal_code_new, d.postalcode) AS postalcode,ISNULL(x.country_new, d.countrykey) AS countrykey, ISNULL(x.contact_new, d.contactname) AS contactname " & _
         " , ISNULL(x.phone_new, d.phone) AS phone,d.SalesDocument, '' + d.PurchasingDocNo AS PurchasingDocNo, f.���,  b.�Ϻ�,d.MarketingPN,SUM(b.����) as ����, d.BatchNumber, d.DATE_CODE " & _
         " , SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2) HTlot_no ,ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  + ( dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0) AS ���� ,ROUND( SUM(b.����) * ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0) " & _
         "  +  (dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0),2) AS AMount ,d.customerPartNumber,ROUND( SUM(b.����) * ISNULL(ISNULL( cd.��˰���� / cc.��Ʒ��,0)  +  (dd.WAFER_PRICE/cc.��Ʒ�� + dd.DIE_PRICE),0),2)   -  CONVERT(DECIMAL(18,2) " & _
         " ,SUM(b.����) * ISNULL( cd.��˰���� / cc.��Ʒ��,0)) AS �ӹ��ѽ�� , CONVERT(DECIMAL(18,2),SUM(b.����) * ISNULL( cd.��˰���� / cc.��Ʒ��,0)) AS �͹��Ͻ�� " & _
         "  FROM erpdata .. tblStockNumTree a INNER JOIN erpdata .. tblStocksqfhsub b ON b.��� = a.��� INNER JOIN erpdata .. tblStockSQfh c  ON c.���ݱ�� = b.���ݱ�� AND c.��� = b.������� " & _
         " INNER JOIN erpdata..tblstock y ON y.�ⷿ���� = c.�ֿ���   INNER JOIN (SELECT a.BOX_ID, SUBSTRING(a.KEY_VALUE, 1,CHARINDEX('|', a.KEY_VALUE) - 1) AS qbox " & _
         " , SUBSTRING(a.KEY_VALUE, CHARINDEX('|', a.KEY_VALUE) + 1, 10) AS job  FROM erpdata .. tblErpInStockDetailInfo a WHERE a.KEY_TYPE = 'T' AND a.KEY_VALUE LIKE '%SS%|%') aa ON b.��� = aa.qbox " & _
         " INNER JOIN (SELECT dn.Delivery,   dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State,dn.postalcode, dn.countrykey, dn.contactname,dn.phone, dn.SalesDocument " & _
         " ,  dn.PurchasingDocNo,dn.BatchNumber,SUM(dn.Quantity) AS Quantity,dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE  FROM ERPBASE .. tblCustomerShippingUp dn " & _
         " WHERE dn.Delivery IN ('" & strDNList & "')  GROUP BY dn.Delivery, dn.shiptoname,dn.shiptostreet1,dn.shiptostreet2, dn.shiptostreet3, dn.city,dn.State, dn.postalcode,  dn.countrykey,dn.contactname,dn.phone " & _
         " , dn.SalesDocument, dn.PurchasingDocNo, dn.BatchNumber, dn.customerPartNumber,dn.MarketingPN, dn.DATE_CODE) d    ON d.Delivery = a.DN  AND d.BatchNumber = aa.job " & _
         " INNER JOIN erpdata .. tblstocknumtree f  ON f.��� = a.�ϼ����  INNER JOIN ERPBASE..tblmappingData qq ON qq.SUBSTRATEID = b.���̿���� AND qq.LOTID = b.������ " & _
         " LEFT JOIN ERPBASE..tblCustomerOI bb  ON CONVERT(VARCHAR(100), bb.ID) = qq.FILENAME AND bb.SOURCE_BATCH_ID = qq.LOTID  LEFT JOIN ERPBASE..tblToInRec_Wafer cc " & _
         " ON cc.��ԲID = REPLACE(qq.SUBSTRATEID,'+','')   AND cc.���� = qq.LOTID  LEFT JOIN ERPBASE..TblToInSub cd ON cd.��ⵥ��� = cc.��ⵥ��� AND cd.�������� = cc.���� " & _
         " LEFT JOIN erptemp..tblBB_CSRPO dd   ON dd.PO_NUM = bb.PO_NUM AND dd.FAB_DEVICE = bb.MPN_DESC LEFT JOIN  erpdata..tblSalerec j ON j.���ݱ�� = c.���ݱ�� AND j.������� = b.������� " & _
         " AND j.С��� = b.��� LEFT JOIN erptemp .. dn_address x  ON dn_address = d.ShipToName  WHERE a.DN IN ('" & strDNList & "')  GROUP BY b.���ݱ��, c.��������, ISNULL(dn_address_new, d.shiptoname) " & _
         "  ,   ISNULL(x.ship_to_street1_new, d.shiptostreet1), ISNULL(x.ship_to_street2_new, d.shiptostreet2),ISNULL(x.ship_to_street3_new, d.shiptostreet3), ISNULL(x.city_new, d.city) " & _
         "  ,  ISNULL(x.dn_st_new, d.State), ISNULL(x.postal_code_new, d.postalcode),ISNULL(x.country_new, d.countrykey),ISNULL(x.contact_new, d.contactname),ISNULL(x.phone_new, d.phone) " & _
         " ,  d.SalesDocument, d.PurchasingDocNo,f.���, b.�Ϻ�,d.MarketingPN,d.BatchNumber, d.DATE_CODE, SUBSTRING(aa.qbox, 2, CHARINDEX('-R', aa.qbox) - 2),d.customerPartNumber,   a.DN " & _
         " , J.���۵����, bb.PO_NUM, cd.��˰����, cc.��Ʒ��, dd.WAFER_PRICE, qq.PASSBINCOUNT, dd.die_price, y.������  )aaa where aaa.shiptoname = '" & ShiptoName & "' order by aaa.shiptoname,aaa.DN,aaa.���"
   
   rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
     

   strExtsion = ".xlsx"    '��ȡ��׺��
   strNewFullPath = "C:\�ϲ�_voice" & "\" & strExName & strExtsion    '��ȡ���ļ�Ҫ�����·��
'    Rs.MoveFirst    '���ݼ��ƶ�����һ��
'    '---------------------------------------------------------------------
    If rs.RecordCount > 0 Then
'        ClsP.ShowProgress 30, "��ʼ��Excel..."
        Set ExApp = New Excel.Application
        ExApp.Visible = False   '�Ƿ���ʾ
        
        Set wkbk = ExApp.Workbooks.Open(strFileName)
        Set wkst = wkbk.Worksheets(1)
      
        DblNum = 0
        DblAmt = 0
        DblJZ = 0
        DblMZ = 0

        wkst.Cells(15, 10) = Trim$("" & rs.Fields(3).Value)
        wkst.Cells(18, 10) = Trim$("" & rs.Fields(11).Value) 'To
        
        wkst.Cells(13, 2) = Trim$("" & rs.Fields(4).Value)
        wkst.Cells(14, 2) = Trim$("" & rs.Fields(5).Value)
        wkst.Cells(15, 2) = Trim$("" & rs.Fields(6).Value) & " " & Trim$("" & rs.Fields(7).Value)
        wkst.Cells(16, 2) = Trim$("" & rs.Fields(8).Value) & " " & Trim$("" & rs.Fields(9).Value) & " " & Trim$("" & rs.Fields(10).Value) & " " & Trim$("" & rs.Fields(11).Value)
        
        wkst.Cells(18, 2) = "Attn:" & Trim$("" & rs.Fields(12).Value) & " ,Tel:" & Trim$("" & rs.Fields(13).Value)
      ' wkst.Cells(19, 2) = ""
        If UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION (CAMARILLO)" Or UCase(rs.Fields(4).Value) = "SEMTECH CORPORATION" Then
        
           wkst.Cells(19, 2) = "TAX ID: 95-2119684"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INTERCONNECT" Then
        
           wkst.Cells(19, 2) = "TAX ID: 82-5035949"
           
        ElseIf UCase(rs.Fields(4).Value) = "SEMTECH COLORADO INCORPORATED(FEDERAL��" Then
        
            wkst.Cells(19, 2) = "TAX ID: 82-5035949"
        
        Else
           
            wkst.Cells(19, 2) = ""
        
        End If
        'wkst.Cells(23, 3) = Trim$("" & Rs.fields(14).Value)
        wkst.Cells(23, 5) = Trim$("'" & rs.Fields(15).Value)
        
        lngRows = 27
        
        IntInertRow = rs.RecordCount * 2
        For i = 1 To IntInertRow - 1
           wkst.Rows(lngRows & ":" & lngRows).Select
           ExApp.Selection.Copy
           ExApp.Selection.Insert Shift:=xlDown
           wkst.Range(lngRows + 1 & ":" & lngRows + 1).Borders.LineStyle = xlNone '�߿���
        Next i
        IntMaxDetailRow = rs.RecordCount
        
        IntBMegerRow = 26
        IntEMegerRow = 29
        intBegin = 1
        Dim QBX As String
        
        For i = 0 To rs.RecordCount - 1
'            wkst.Cells(lngRows, 1) = Trim$("" & Rs.fields(16).Value) '���

             If dnnum1 <> Trim$("" & rs.Fields(2).Value) And InStr(dnnum, Trim$("" & rs.Fields(2).Value)) = 0 Then
                dnnum = Trim$("" & rs.Fields(2).Value) + "/" + dnnum
                dnnum1 = Trim$("" & rs.Fields(2).Value)
             End If
             

            
            strPBigBox = Trim$("" & rs.Fields(16).Value) '���
            If strPBigBox <> strNBigBox Then
                strNBigBox = Trim$("" & rs.Fields(16).Value) '���
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
              If SD <> Trim$("" & rs.Fields(14).Value) Then
             SD = Trim$("" & rs.Fields(14).Value)
             SD1 = SD1 & SD & " "
             End If
            wkst.Cells(23, 3) = SD1
            
            wkst.Cells(lngRows, 3) = Trim$("" & rs.Fields(15).Value) 'PO
            wkst.Cells(lngRows, 4) = Trim$("" & rs.Fields(17).Value)
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(18).Value)
            wkst.Cells(lngRows, 7) = Val(Trim$("" & rs.Fields(19).Value)) / 1000 '������Ϊ����1000��ֵ
            DblNum = DblNum + Val(Trim$("" & rs.Fields(19).Value))
            wkst.Cells(lngRows, 9) = "KPCS"
            wkst.Cells(lngRows, 10) = Trim$("" & rs.Fields(20).Value)
            wkst.Cells(lngRows, 11) = Trim$("" & rs.Fields(21).Value)
            wkst.Cells(lngRows, 12) = Trim$("" & rs.Fields(22).Value)
            wkst.Cells(lngRows, 13) = "US$"
            wkst.Cells(lngRows, 14) = Val(Trim$("" & rs.Fields(23).Value)) * 1000 '����Ϊ����1000��ֵ
            wkst.Cells(lngRows, 15) = "US$"
            wkst.Cells(lngRows, 16) = Trim$("" & rs.Fields(24).Value)
            DblAmt = DblAmt + Val(Trim$("" & rs.Fields(24).Value))
            lngRows = lngRows + 1
            
            wkst.Cells(lngRows, 4) = "CPN:"
            wkst.Cells(lngRows, 5) = Trim$("" & rs.Fields(25).Value)
            
            jine1 = jine1 + Val(Trim$("" & rs.Fields(26).Value))
            
            jine2 = jine2 + Val(Trim$("" & rs.Fields(27).Value))
            
            
            
            
            lngRows = lngRows + 1
            IntEMegerRow = lngRows
            rs.MoveNext
        Next
        
        wkst.Cells(13, 10) = Mid(dnnum, 1, Len(dnnum) - 1)
        
        '�������
        wkst.Cells(lngRows + 1, 7) = DblNum / 1000 '����
        wkst.Cells(lngRows + 1, 9) = "KPCS" '��λ
        wkst.Cells(lngRows + 1, 16) = DblAmt
        wkst.Cells(lngRows + 1, 1) = Trim(intBoxNum - 1)
        
        wkst.Cells(lngRows + 8, 1) = "Process Amount��US$ " + Str(jine1)
        wkst.Cells(lngRows + 9, 1) = "Wafer Amount��US$ " + Str(jine2)

        
    Else
'        ClsP.UnLoad_Form
        MsgBox "���赼�����ݣ�", vbInformation, "��ʾ��"
        Exit Function
    End If

    With wkst.PageSetup

    End With
    '---------------------------------------------------------------------------------------------------------------
    If Len(Dir(strNewFullPath)) > 0 Then
        If MsgBox("���ļ��Ѿ����ڣ��Ƿ�Ҫ����ԭ�ļ�?", vbYesNo Or vbQuestion Or vbDefaultButton2, "��ʾ") = vbNo Then
            Exit Function
        Else
            On Error Resume Next
            Kill strNewFullPath
            If Err.number <> 0 Then
                MsgBox "�����ļ�ʧ�ܣ����ֶ�ɾ���ļ��ٵ�����", vbInformation, "��ʾ"
                Exit Function
            End If
        End If
    End If
    'wkbk.SaveAs strNewFullPath, xlNormal, "", "", False, False
    wkbk.SaveAs strNewFullPath
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
    Exit Function
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
    MsgBox Err.DESCRIPTION, vbInformation, "��ʾ��"
    Exit Function
End Function


'����Rs���ݼ���䵼��Excel
Public Sub RsExporToExcel(rs As ADODB.Recordset, RptName As String, ExcelFileName As String)
Dim Irowcount       As Long
Dim Icolcount       As Integer
Dim strFileName     As String

    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlQuery As Excel.QueryTable
    Screen.MousePointer = 11
    With rs
        If .RecordCount < 1 Then
            Screen.MousePointer = 0
            MsgBox ("û�пɵ���������")
            Exit Sub
        End If
        Irowcount = .RecordCount
        Icolcount = .Fields.count
    End With
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlBook = xlApp.Workbooks().Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlApp.Visible = False

    Set xlQuery = xlSheet.QueryTables.Add(rs, xlSheet.Range("a1"))
    
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
    xlSheet.name = RptName
    xlQuery.FieldNames = True '�W
    xlQuery.Refresh
    
    With xlSheet
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.name = "����"
        '�r
        .Range(.Cells(1, 1), .Cells(1, Icolcount)).Font.Bold = True
        '
        .Range(.Cells(1, 1), .Cells(Irowcount + 1, Icolcount)).Borders.LineStyle = xlContinuous
        '
'        .Range(.Cells(2, 1), .Cells(Irowcount + 1, Icolcount)).Font.Size = 9
    End With
    '����ļ�
    strFileName = DirInvRpt + "\" + ExcelFileName
    'xlBook.SaveAs strFileName, xlNormal, "", "", False, False
    xlBook.SaveAs strFileName
    xlBook.Saved = True
    
    Screen.MousePointer = 0
'    xlApp.Visible = True
    Set xlSheet = Nothing
    xlBook.Close
    Set xlBook = Nothing
    xlApp.Quit
    Set xlApp = Nothing
  

End Sub








