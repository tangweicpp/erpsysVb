VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmSemtech_LablePrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Semtech标签打印"
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
         Caption         =   "查  询"
         Height          =   360
         Index           =   0
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmd 
         Caption         =   "导出当前数据"
         Height          =   360
         Index           =   1
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd 
         Caption         =   "退 出"
         Height          =   360
         Index           =   3
         Left            =   5760
         TabIndex        =   11
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmd 
         Caption         =   "打  印"
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
         Caption         =   "散袋LOT"
         Height          =   195
         Left            =   7200
         TabIndex        =   33
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "查询条件"
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
            Caption         =   "结束日期"
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
            Caption         =   "开始日期"
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
         Caption         =   "标签补打"
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
            Name            =   "宋体"
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
         Caption         =   "数        量"
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
         Caption         =   "外箱路径"
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
         Caption         =   "内盒路径"
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
         Caption         =   "卷盘路径"
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
            Name            =   "宋体"
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
         Caption         =   "标签类型"
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
         Caption         =   "散袋1"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   6120
         TabIndex        =   35
         Top             =   0
         Width           =   975
      End
      Begin VB.OptionButton Opt 
         Caption         =   "散袋2"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   32
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkChoose 
         Caption         =   "全选"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton Opt 
         Caption         =   "外箱"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   28
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton Opt 
         Caption         =   "内箱"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   27
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton Opt 
         Caption         =   "卷盘"
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

'Check框的变化
Private Sub chk_Click()
    If chk.Value = 1 Then
        Fra(3).Visible = True
    Else
        Fra(3).Visible = False
    End If
End Sub
'全选或取消
Private Sub chkChoose_Click()
Dim i           As Long

    If Fps(0).MaxRows <= 0 Then
        chkChoose.Value = 0
    Else
        With Fps(0)
            For i = 1 To .MaxRows
                If chkChoose.Value = 1 Then '全选
                    .SetText FpsDetail.e_Choose, i, "1"
                Else    '取消
                    .SetText FpsDetail.e_Choose, i, "0"
                End If
            Next
        End With
    End If
    
End Sub

'DN的变化
Private Sub cmbDN_Click()
Dim i                   As Long
Dim strSql              As String
Dim Rs                  As New adodb.Recordset
    
    '先初始化控件
    For i = 0 To txt.UBound
        txt(i).Text = ""
    Next
    chk.Value = 0
    Fps(0).MaxRows = 0
    chkChoose.Value = 0
    '查询数据赋值到控件
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
            strShipToCust = Trim$("" & Rs!ShipToCustomer)   '出货地
            Rs.MoveNext
        Loop
    End If
    Rs.Close

End Sub

Private Sub cmd_Click(Index As Integer) '查询报表
Dim i                   As Long
Dim strSql              As String
Dim Rs                  As New adodb.Recordset
Dim strExportName       As String
Dim lotTemp As String
lotTemp = Trim(txtText1.Text)
    '---------------------------------------------
    If Index = 0 Then          '查询
        If cmbDN.Text = "" And Opt(3).Value = False And Opt(4).Value = False Then
            MsgBox "请先选择DN号！"
            Exit Sub
        End If
        If Opt(0).Value = True Then '卷盘
            If strShipToCust = "2000561" Then   '出LG的标签
                strSql = "SELECT 0 '选择',a.TRAYQBOXNUMBER QBOXNUMBER,b.CustomerPartNumber,b.BatchNumber,a.HTLOTID,a.PoDateCode,a.QTY " & _
                         " FROM erpdata..TblTSV_Tray_details a " & _
                         " INNER JOIN (SELECT DISTINCT Delivery,CustomerPartNumber,BatchNumber,VendorLotNumber " & _
                         " FROM erpbase..tblCustomerShippingUp WHERE flag='Y' AND customershortname='37') b ON a.CUSTOMERLOTID=b.BatchNumber " & _
                         " Inner Join erpdata..tblstocknumsub c on a.TRAYQBOXNUMBER=rtrim(c.箱号) " & _
                         " WHERE b.Delivery='" & Trim$(cmbDN.Text) & "'"
            Else
                strSql = "SELECT 0 '选择',a.TRAYQBOXNUMBER QBOXNUMBER,b.CustomerPartNumber,'TVS DIODES' Specification" & _
                         ",CASE WHEN CHARINDEX('E2',b.LabelRequirement)>0 THEN 'PO TYPE,:E2' ELSE ',' END Potype" & _
                         ",a.CUSTOMERLOTID,a.QTY,b.MarketingPN,'DPTK' VendorCode " & _
                         " FROM erpdata..TblTSV_Tray_details a " & _
                         " INNER JOIN (SELECT DISTINCT Delivery,CustomerPartNumber,LabelRequirement,MarketingPN,BatchNumber " & _
                         " FROM erpbase..tblCustomerShippingUp WHERE flag='Y' AND customershortname='37') b ON a.CUSTOMERLOTID=b.BatchNumber " & _
                         " Inner Join erpdata..tblstocknumsub c on a.TRAYQBOXNUMBER=rtrim(c.箱号) " & _
                         " WHERE b.Delivery='" & Trim$(cmbDN.Text) & "'"
            End If
        End If
        If Opt(1).Value = True Then '内盒
             If strShipToCust = "2000561" Then   '出LG的标签
                strSql = "SELECT 0 '选择',a.CONTAINERNAME QBOXNUMBER,b.CustomerPartNumber,b.BatchNumber,a.HTLOTID,a.PoDateCode,SUM(a.QTY) Qty " & _
                         " FROM erpdata..TblTSV_INBOX_DETAILS a " & _
                         " INNER JOIN (SELECT DISTINCT Delivery,CustomerPartNumber,BatchNumber,VendorLotNumber " & _
                         " FROM erpbase..tblCustomerShippingUp WHERE flag='Y' AND customershortname='37') b ON a.CUSTOMERLOTID=b.BatchNumber " & _
                         " Inner Join erpdata..tblstocknumsub c on a.SUBCONTAINERNAME=rtrim(c.箱号) " & _
                         " WHERE b.Delivery='" & Trim$(cmbDN.Text) & "'"
             Else
                strSql = "SELECT 0 '选择',a.NHBox QBOXNUMBER,b.CustomerPartNumber,'TVS DIODES' Specification" & _
                         ",CASE WHEN CHARINDEX('E2',b.LabelRequirement)>0 THEN 'PO TYPE,:E2' ELSE ',' END Potype" & _
                         ",SUM(a.QTY) Qty,b.MarketingPN,'DPTK' VendorCode " & _
                         " FROM erpdata..TblTSV_INBOX_DETAILS a " & _
                         " INNER JOIN (SELECT DISTINCT Delivery,CustomerPartNumber,LabelRequirement,MarketingPN,BatchNumber " & _
                         " FROM erpbase..tblCustomerShippingUp WHERE flag='Y' AND customershortname='37') b ON a.CUSTOMERLOTID=b.BatchNumber " & _
                         " Inner Join erpdata..tblstocknumsub c on a.SUBCONTAINERNAME=rtrim(c.箱号) " & _
                         " WHERE b.Delivery='" & Trim$(cmbDN.Text) & "'"
            End If
        End If
        If Opt(2).Value = True Then '外箱
             strSql = "SELECT 0 '选择',a.CONTAINERNAME QboxNumber,a.Invoice,'I'+a.Invoice Invoice1,left(a.PONO,10),'K'+left(a.PONO,10) PONO1,CASE WHEN CHARINDEX('E2',b.LabelRequirement)>0 THEN 'E2' ELSE '' END Potype" & _
                      ",left(a.customerPT,11),'P'+left(a.customerPT,11) customerPT1,a.MFGPT,'Z'+a.MFGPT MFGPT1,SUM(a.Qty) 数量,'Q'+Rtrim(SUM(a.Qty)) 数量1,a.Forwarder,a.coo " & _
                      ",left(a.shiptoname,33),a.shiptostreet1,a.shiptostreet2,a.shiptostreet3,a.shiptostreet4,a.countrykey " & _
                      ",'Attn:'+a.contactname+';Tel:'+a.phone 联系人, 'N/A','P' +'N/A','N/A','9D' + 'N/A'" & _
                      " FROM erpdata..TblTSV_OutBOX_DETAILS a " & _
                      " INNER JOIN (SELECT DISTINCT Delivery,LabelRequirement FROM erpbase..tblCustomerShippingUp WHERE flag='Y' AND customershortname='37') b ON a.Invoice=b.Delivery " & _
                      " WHERE a.Invoice='" & Trim$(cmbDN.Text) & "'"
        End If
        '这段通用
        If chk.Value = 0 Then
            strSql = strSql & " And a.PrintFlag=0 "
        Else    '补打标签
            strSql = strSql & " And a.PrintFlag=1 And a.PrintTime>='" & DTP(0).Value & "' And a.PrintTime<'" & DTP(1).Value + 1 & "'"
        End If
        
        If Opt(1).Value = True Then '内盒
            If strShipToCust = "2000561" Then   '出LG的标签
                strSql = strSql & " GROUP BY a.CONTAINERNAME,b.CustomerPartNumber,b.BatchNumber,a.HTLOTID,a.PoDateCode "
            Else
                strSql = strSql & " GROUP BY  a.NHBox,b.CustomerPartNumber " & _
                        ",CASE WHEN CHARINDEX('E2',b.LabelRequirement)>0 THEN 'PO TYPE,:E2' ELSE ',' END,b.MarketingPN"
            End If
        End If
        If Opt(2).Value = True Then '外箱
            strSql = strSql & " GROUP BY a.CONTAINERNAME,a.Invoice,'I'+a.Invoice,left(a.PONO,10),'K'+left(a.PONO,10),CASE WHEN CHARINDEX('E2',b.LabelRequirement)>0 THEN 'E2' ELSE '' END" & _
                     ",left(a.customerPT,11),'P'+left(a.customerPT,11),a.MFGPT,'Z'+a.MFGPT,a.forwarder,a.coo,left(a.shiptoname,33),a.shiptostreet1,a.shiptostreet2 " & _
                     ",a.shiptostreet3,a.shiptostreet4,a.countrykey,'Attn:'+a.contactname+';Tel:'+a.phone"
        End If
        
        If Opt(3).Value = True Then
        If Len(lotTemp) < 2 Then
           MsgBox "请输入LOT号！"
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
           MsgBox "请输入LOT号！"
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
        
        '查询数据到FPS
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
    ElseIf Index = 1 Then   '导出
        
        If Opt(0).Value = True Then
            strExportName = Opt(0).Caption + "标签信息"
        ElseIf Opt(1).Value = True Then
            strExportName = Opt(1).Caption + "标签信息"
        ElseIf Opt(2).Value = True Then
            strExportName = Opt(2).Caption + "标签信息"
        End If
        If Not ExportFpspreadToExcel(Fps(0), strExportName, strExportName) Then Exit Sub
    
    ElseIf Index = 2 Then   '打印
      
        '校验数据
        If Not CheckData Then Exit Sub
        '初始化Fps数据，构造标签头，打印标签
        Call IniLable
        
    ElseIf Index = 3 Then   '退出
        Unload Me
    End If
'
End Sub
'初始化标签组合
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
            If .Value = 1 Then  '选择了打印的行
                strLable = ""
                strFileName = ""
                If Opt(0).Value = True Then     '卷盘
                    If strShipToCust = "2000561" Then   '出LG的标签
                        .Col = 2    '箱号
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
                        '构造标签语句
                        strLable = strTmp(0) + "," + strTmp(1) + "," + strTmp(2) + "," + strTmp(3) + "," + strTmp(4)
                    Else
                        .Col = 2    '箱号
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
                        strTmp(4) = GetLableXH(strTmp(3)) '序号
                        .Col = 7    'Qty
                        strTmp(5) = Trim$(.Text)
                        .Col = 8    'VenDor P/N
                        strTmp(6) = Trim$(.Text)
                        .Col = 9    'VenDor Code
                        strTmp(7) = Trim$(.Text)
                        '构造标签语句
                        strLable = strTmp(0) + strTmp(7) + strTmp(2) + Left$(strTmp(3) + "00000000", 8) + strTmp(4) + Right$("000000" + strTmp(5), 6) + ","
                        strLable = strLable + strTmp(0) + "," + strTmp(1) + "," + strTmp(8) + "," + Left$(strTmp(3) + "00000000", 8) + strTmp(4) + "," + strTmp(5) + ","
                        strLable = strLable + strTmp(6) + "," + strTmp(7)
                    End If
                    '开始打印TXT文件到指定位置
                    Call PrintLable(strFileName, strLable, Trim(txt(2).Text)) '卷盘
                End If
                If Opt(1).Value = True Then     '内盒
                    If strShipToCust = "2000561" Then   '出LG的标签
                        .Col = 2    '箱号
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
                        '构造标签语句
                        strLable = strTmp(0) + "," + strTmp(1) + "," + strTmp(2) + "," + strTmp(3) + "," + strTmp(4)
                    Else
                        .Col = 2    '箱号
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
                        '构造标签语句
                        strLable = strTmp(0) + strTmp(5) + strTmp(2) + Right$("000000" + strTmp(3), 6) + ","
                        strLable = strLable + strTmp(0) + "," + strTmp(1) + "," + strTmp(6) + "," + strTmp(3) + ","
                        strLable = strLable + strTmp(4) + "," + strTmp(5)
                    End If
                    '开始打印TXT文件到指定位置
                    Call PrintLable(strFileName, strLable, Trim(txt(3).Text)) '内盒
                End If
                If Opt(2).Value = True Then     '外箱
                    .Col = 2    '箱号
                    strFileName = Trim$(.Text)
                    For j = 3 To .MaxCols
                        .Col = j
                        strLable = strLable + Trim$(.Text) + ","    '拼接标签
                    Next
                    strLable = Left$(strLable, Len(strLable) - 1)   '去除最后一个逗号
                    '开始打印TXT文件到指定位置
                    Call PrintLable(strFileName, strLable, Trim(txt(4).Text)) '外箱
                End If
                    If Opt(3).Value = True Then      '外箱
                    .Col = 4    '箱号
                    strFileName = Trim$(.Text)
                    For j = 2 To .MaxCols
                        .Col = j
                        strLable = strLable + Trim$(.Text) + ","    '拼接标签
                    Next
                    strLable = Left$(strLable, Len(strLable) - 1)   '去除最后一个逗号
                    '开始打印TXT文件到指定位置
                    Call PrintLable(strFileName, strLable, "\\10.160.1.14\BarCode\37\37DIE2-2\") '外箱
                End If
                  If Opt(4).Value = True Then      '外箱
                    .Col = 9    '箱号
                    strFileName = Trim$(.Text)
                    For j = 2 To .MaxCols
                        .Col = j
                        strLable = strLable + Trim$(.Text) + ","    '拼接标签
                    Next
                    strLable = Left$(strLable, Len(strLable) - 1)   '去除最后一个逗号
                    '开始打印TXT文件到指定位置
                    Call PrintLable(strFileName, strLable, "\\10.160.1.14\BarCode\37\37DIE2-1\") '外箱
                End If
                
                '将打印过的箱号删除
                .DeleteRows i, 1
                .MaxRows = .MaxRows - 1
            End If
        Next

    End With
    
    MsgBox "打印成功！"

End Sub
'获取标签序号
Private Function GetLableXH(strLot As String) As String
Dim strSql          As String
Dim Rs              As New adodb.Recordset
Dim strXH           As String
Dim intCount        As Integer
Dim strLot1          As String
    
    If strLot = "" Then Exit Function
    intCount = 0
    strLot1 = Replace(strLot, "M", "")
    strSql = "SELECT dbo.F_GetPrintXH('" & strLot1 & "') 序号"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    If Not Rs.EOF Then
        strXH = Trim$("" & Rs!序号)
        If strXH <> "" Then '如果有得到序号，就更新数据
            strSql = "Update erpdata..tblSysIncrement Set Para='" & strXH & "',ICount=ICount+1 Where Kind='" & strLot & "'"
            INIadoCon2.Execute strSql, intCount
            If intCount <= 0 Then   '表示不存在此LOT信息，就插入一笔
                strSql = "Insert Into erpdata..tblSysIncrement(Kind,Para,ICount) Values('" & strLot & "','" & strXH & "',1)"
                INIadoCon2.Execute strSql
            End If
        End If
    End If
    Rs.Close
    
    GetLableXH = strXH  '赋值回去
    
End Function
'标签组合打印
Private Sub PrintLable(strFileName As String, strTxt As String, strTxtPath As String)
Dim i               As Long
Dim strSql          As String
Dim Rs              As New adodb.Recordset
    
    '调用过程
    Call PrintLabelTxt(strFileName, strTxt, strTxtPath)
    '更新打印标记和时间
    If Opt(0).Value = True Then     '卷盘
        strSql = "Update erpdata..TblTSV_Tray_details Set PrintFlag=1,PrintTime=getdate() Where TRAYQBOXNUMBER='" & strFileName & "'"
        INIadoCon2.Execute strSql
    End If
    If Opt(1).Value = True Then     '内盒
        strSql = "Update erpdata..TblTSV_INBOX_DETAILS Set PrintFlag=1,PrintTime=getdate() Where CONTAINERNAME='" & strFileName & "'"
        INIadoCon2.Execute strSql
    End If
    If Opt(2).Value = True Then     '外箱
        strSql = "Update erpdata..TblTSV_OutBOX_DETAILS Set PrintFlag=1,PrintTime=getdate() Where CONTAINERNAME='" & strFileName & "'"
        INIadoCon2.Execute strSql
    End If
End Sub
'2016-09-08 mwl add 写TXT标签文件
Private Sub PrintLabelTxt(FileName As String, msgTxt As String, dirtemp As String)
'判断txt文件是否存在，如不存在，则建立
Dim fileNameTemp        As String
Dim dirNameTemp         As String
Dim fileTemp            As String

    dirNameTemp = dirtemp
    fileNameTemp = Replace(FileName, "'", "") & ".txt"
    fileTemp = dirNameTemp & fileNameTemp
    
    Open fileTemp For Output As #1   '直接覆盖
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

    '初始化控件
    InitCtrl
    
End Sub

'初始化控件
Private Sub InitCtrl()
Dim i                   As Integer
Dim strSql              As String
Dim Rs                  As New adodb.Recordset
    
    strdjbh = ""
    '加载单据类型
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
    '初始化FPS
    InitFps
    
    chk.Value = 0
    DTP(0).Value = Format(Now(), "YYYY/MM/DD")
    DTP(1).Value = Format(Now(), "YYYY/MM/DD")
    Fra(3).Visible = False
   
End Sub
'初始化FPS控件
Public Sub InitFps()
Dim i                   As Integer
    'Fps初始化
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
        '设定列类型
        .Col = FpsDetail.e_Choose   '选择
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeVAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        '设定列宽
        .ColWidth(-1) = 10
        .ColWidth(FpsDetail.e_Choose) = 4
        .RowHeight(-1) = 10
        '设定是否排序
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
    
    '点击把选择的单号都选上
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

'校验数据
Private Function CheckData() As Boolean
Dim i               As Long
Dim intCount        As Integer

    CheckData = False
    
    intCount = 0
    
    With Fps(0)
        If .MaxRows <= 0 Then
            MsgBox "没有任何资料,请先查询！", vbInformation, "提示"
            Exit Function
        End If
        '看是否有选择
        For i = 1 To .MaxRows
            .Row = i
            .Col = FpsDetail.e_Choose  '选择
            If .Value = 1 Then
                intCount = intCount + 1
            End If
        Next
    End With
    '--------------------------
    If intCount <= 0 Then
        MsgBox "没有选择任何资料！", vbInformation, "提示"
        Exit Function
    End If
    '校验是否有标签路径
    If Opt(0).Value = True Then '卷盘
        If Trim(txt(2).Text) = "" Then
            MsgBox "没有设定此客户的卷盘标签路径，请联系系统管理员！", vbInformation, "提示"
            Exit Function
        End If
    End If
    If Opt(1).Value = True Then '内盒
        If Trim(txt(3).Text) = "" Then
            MsgBox "没有设定此客户的内盒标签路径，请联系系统管理员！", vbInformation, "提示"
            Exit Function
        End If
    End If
    If Opt(2).Value = True Then '外箱
        If Trim(txt(4).Text) = "" Then
            MsgBox "没有设定此客户的外箱标签路径，请联系系统管理员！", vbInformation, "提示"
            Exit Function
        End If
    End If
    CheckData = True
End Function
'卷盘，内箱，外箱的变化
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
