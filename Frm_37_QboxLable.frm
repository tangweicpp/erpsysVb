VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_37_QboxLabel 
   Caption         =   "Semtech 内箱，外箱标签"
   ClientHeight    =   11070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14190
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
   ScaleHeight     =   11070
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   16325
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "内箱"
      TabPicture(0)   =   "Frm_37_QboxLable.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdReset"
      Tab(0).Control(1)=   "txtScan"
      Tab(0).Control(2)=   "TxtWaferIDIn"
      Tab(0).Control(3)=   "CmdOK"
      Tab(0).Control(4)=   "CmdExitIn"
      Tab(0).Control(5)=   "TxtDirInQbox"
      Tab(0).Control(6)=   "lbl"
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(8)=   "Label2"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "外箱"
      TabPicture(1)   =   "Frm_37_QboxLable.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label24"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "TxtDirOutQbox"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "CmdExitOut"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "CmdOKOut"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "TxtWaferIDOut"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "TxtDirOutHTQbox"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "ComDN"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox ComDN 
         Height          =   375
         Left            =   2400
         TabIndex        =   20
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "初始化"
         Height          =   480
         Left            =   -71040
         TabIndex        =   19
         Top             =   8400
         Width           =   1455
      End
      Begin VB.TextBox txtScan 
         Height          =   285
         Left            =   -73560
         TabIndex        =   18
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox TxtDirOutHTQbox 
         Height          =   375
         Left            =   8400
         TabIndex        =   13
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox TxtWaferIDOut 
         Height          =   6615
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1800
         Width           =   9255
      End
      Begin VB.CommandButton CmdOKOut 
         Caption         =   "确定"
         Height          =   480
         Left            =   2040
         TabIndex        =   9
         Top             =   8520
         Width           =   1455
      End
      Begin VB.CommandButton CmdExitOut 
         Caption         =   "取消"
         Height          =   480
         Left            =   4680
         TabIndex        =   8
         Top             =   8520
         Width           =   1575
      End
      Begin VB.TextBox TxtDirOutQbox 
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox TxtWaferIDIn 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   6135
         Left            =   -74520
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1920
         Width           =   9255
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "确定"
         Height          =   480
         Left            =   -73440
         TabIndex        =   3
         Top             =   8400
         Width           =   1455
      End
      Begin VB.CommandButton CmdExitIn 
         Caption         =   "退出"
         Height          =   480
         Left            =   -68640
         TabIndex        =   2
         Top             =   8400
         Width           =   1455
      End
      Begin VB.TextBox TxtDirInQbox 
         Height          =   285
         Left            =   -73560
         TabIndex        =   1
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描中："
         Height          =   195
         Left            =   -74400
         TabIndex        =   17
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN#："
         Height          =   195
         Left            =   1800
         TabIndex        =   15
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HT外箱Txt路径："
         Height          =   195
         Left            =   6960
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫入的内箱："
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semtech外箱Txt路径："
         Height          =   195
         Left            =   600
         TabIndex        =   11
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫入的卷盘："
         Height          =   315
         Left            =   -74640
         TabIndex        =   6
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Txt路径："
         Height          =   195
         Left            =   -74400
         TabIndex        =   5
         Top             =   720
         Width           =   780
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   195
      Left            =   6480
      TabIndex        =   16
      Top             =   4440
      Width           =   465
   End
End
Attribute VB_Name = "Frm_37_QboxLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim useridTemp As String
Dim txtFlag As String
Dim txtTimer As String
Dim sTrayNo(8) As String
Dim iPos As Integer


Private Sub cmdExit_Click()
txtWaferID.Text = ""
txtWaferID.SetFocus
End Sub

Private Sub cmdFlag_Click()
txtTextFlag.SetFocus
End Sub

Private Sub cmdReset_Click()
txtScan.SetFocus
InitData
iPos = 0
End Sub

Private Sub ComDN_DropDown()

Call InitCtrl2

End Sub


Private Sub InitCtrl2()
Dim i                   As Integer
Dim strSql              As String
Dim rs                  As New ADODB.Recordset
    
    ComDN.clear
    '加载单据类型
    strSql = "select distinct delivery  from   CUSTOMERSHIPPINGUPTBL  a where a.flag='Y' order by a.delivery desc "
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    ComDN.clear
    If Not rs.EOF Then
        Do While Not rs.EOF
            ComDN.AddItem Trim$("" & rs!Delivery)
            rs.MoveNext
        Loop
        'ComDN.ListIndex = 0
    End If
    rs.Close
   
End Sub


Private Sub CmdExitIn_Click()

Unload Me
End Sub

Private Sub CmdExitOut_Click()
TxtWaferIDOut.Text = ""
TxtWaferIDOut.SetFocus
End Sub

Private Sub CmdOK_Click()
txtScan.SetFocus
'把资料生成一个txt
Dim txtStr As String
Dim dirtemp As String
Dim cmdStr2 As String

Dim fileNameTemp As String
Dim msgTxtTemp As String
Dim msgTxtTemp2 As String
Dim qboxNoTemp As String
Dim qboxNoContainerTemp As String
Dim inBoxContainerTemp As String
Dim qboxNoSeqTemp As String
Dim qboxNoSeqTemp1 As String
Dim inboxnum As String
Dim stqtpj As String
Dim sqlDB As String
Dim sqlDBRS As New ADODB.Recordset

If iPos = 0 Then
    Exit Sub
End If

qboxNoSeqTemp = "-B00"
fileNameTemp = ""
msgTxtTemp = ""

txtStr = TxtWaferIDIn.Text
txtStr = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(txtStr, "卷盘1: ", ""), "卷盘2: ", ""), "卷盘3: ", ""), "卷盘4: ", ""), "卷盘5: ", ""), "卷盘6: ", ""), "卷盘7: ", ""), "卷盘8: ", ""), "卷盘9: ", "")
msgTxtTemp = Replace(txtStr, vbCrLf, "','")

''1234,'456,'789'
msgTxtTemp2 = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & "," & Right(msgTxtTemp, Len(msgTxtTemp) - InStr(msgTxtTemp, ","))

Dim bid

bid = sTrayNo
Dim lotStr As String

For i = 0 To iPos - 1
    lotStr = bid(i)
   
    If lotStr <> "" Then
    
        '先判断是否在仓库
        If Not Judge37TrayIn(lotStr) Then
            MsgBox "此卷：" & lotStr & " 不存在于ERP仓库中，不可以合内箱，请确认!", vbInformation, "友情提示"
            InitData
            iPos = 0
    
         '  Exit Sub
        Else

            '再判断是不是在60000仓与60001仓库
            If Judge37InvType(lotStr) Then
                MsgBox "此卷：" & lotStr & " 存在于ERP 6000或6001仓中，不可以合内箱，请确认!", vbInformation, "友情提示"
                InitData
                iPos = 0
          '      Exit Sub 'CCS ADD 20160720
            End If

        End If
    
        '先判断有没有装过
        If Judge37ExistInBox1(lotStr) Then
             MsgBox "此卷：" & lotStr & " 已装过内箱，不可以重复装，请确认!", vbInformation, "友情提示"
            InitData
            iPos = 0
          '  Exit Sub 'CCS ADD 20160720
        End If
     End If

Next i

'根据刘浩提出的问题，增加装箱出一个内箱箱号 箱号规则NH+年月日+4位流水码
Dim strRQ       As String
Dim strLSM      As String
Dim strSql      As String
Dim strqbnum As String
Dim strqbnum1 As String
Dim finame As String
Dim qbnum As String
Dim rs          As New ADODB.Recordset
Dim Rs1          As New ADODB.Recordset
    strRQ = "NH" + Format(Now(), "YYMMDD")
    strSql = "SELECT MAX(NHBox) NHBox FROM erpdata..TblTSV_INBOX_DETAILS WHERE NHBox LIKE '" & strRQ & "%'"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        If Trim("" & rs!NHBox) = "" Then    '表示当前还没合箱
            strLSM = strRQ + "0001"
        Else
            strLSM = strRQ + Right("0000" + Trim$(Val(Right(Trim$("" & rs!NHBox), 4)) + 1), 4)
        End If
    End If
    rs.Close

strqbnum = " select b.htlotid as firname   from TSV_Tray_details b  " & _
" where b.trayqboxnumber in ('" & msgTxtTemp2 & "')   group by b.htlotid  "
 
 If rs.State = adStateOpen Then rs.Close
    rs.Open strqbnum, Cnn, adOpenStatic, adLockReadOnly, adCmdText
   
  If Not rs.EOF Then
    Do While Not rs.EOF
    finame = rs!firname
    
    strqbnum1 = "select  '-B'|| substr('00'||(nvl(max(a.seqtxt),0)+1),-2)  from TSV_QBOXTBL_37SEQ a where a.firtname = '" & finame & "' group by a.firtname"
    If Rs1.State = adStateOpen Then Rs1.Close
    Rs1.Open strqbnum1, Cnn, adOpenStatic, adLockReadOnly, adCmdText
  If Not Rs1.EOF Then
    qbnum = Rs1.Fields(0).Value
    Else
    qbnum = "-B01"
    End If
     
    
sqlDB = " select replace(a.customerpt,'.P2','') +','+ replace(a.customerlotid,'M','') +','+'1T'+ replace(a.customerlotid,'M','') +','+ replace(a.customerpt,'.P2','') +','+'1P'+ replace(a.customerpt,'.P2','') +','+min(a.podatecode) +','+min(a.podatecode)+',' " & _
" +max(a.htlotid)+'" & qbnum & "'+','+'S'+max(a.htlotid)+'" & qbnum & "' +','+rtrim(sum(qty)) +','+'Q'+rtrim(sum(qty)) +','" & _
" +min(a.htdatecode) +','+min(a.htdatecode) " & _
" from  [erpdata].[dbo].[TblTSV_Tray_details] a where trayqboxnumber in ('" & msgTxtTemp2 & "')  and a.htlotid = '" & finame & "'  " & _
" group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htlotid,a.htdatecode"

  If sqlDBRS.State = adStateOpen Then sqlDBRS.Close
    sqlDBRS.Open sqlDB, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    
 Dim pp As Integer
 pp = 0
        pp = pp + 1
fileNameTemp = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & Format(Now(), "YYYYMMDDHHmmSS") & Trim(pp)
dirtemp = TxtDirInQbox.Text
Call addLabelTxt(fileNameTemp, sqlDBRS.Fields(0).Value, dirtemp)
 



'add 把 内箱号，几个Tray号，保存到一个表里
For i = 0 To iPos - 1
    lotStr = bid(i)
    
   stqtpj = Mid(qbnum, 3, 2)
   
    
    If Mid(lotStr, 2, InStr(lotStr, "-") - 2) = finame Then
    
'   If Judge37ExistInBox2(inboxnum) Then
'
'   MsgBox "箱号已存在"
'       Exit Sub
'
'    End If
    
      ' lotStr = Right(lotStr, Len(lotStr) - 1)
       
    
 cmdStr2 = " insert into [erpdata].[dbo].[TblTSV_INBOX_DETAILS](id,Containername,Subcontainername,Labeltype,customerpt,customerlotid,htlotid,podatecode,htdatecode,qty,flag,created_by,created_date,NHBox)" & _
 " select   1, 'S'+inboxno,trayboxno,typename,mpn_desc,jobnumber,htlotid,date_code,htdatecode,qty,'Y','" & useridTemp & "',getdate(),'" & strLSM & "' from ( " & _
 " select  max(a.htlotid) +'" & qbnum & "'  as inboxno,  '" & lotStr & "' as trayboxno , 'INQbox' as typename," & _
" a.customerpt as mpn_desc,a.customerlotid as jobnumber,max(a.htlotid) as htlotid ,min(a.podatecode) as date_code,min(a.htdatecode) as htdatecode,sum(qty) as qty" & _
" from  [erpdata].[dbo].[TblTSV_Tray_details] a where trayqboxnumber in ('" & bid(i) & "' )  " & _
" group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htdatecode ) X "
 
 cmdStpj = " insert into TSV_QBOXTBL_37SEQ(typename,createdate,seqtxt,containername,Firtname) " & _
          "  values ('INQbox',sysdate,'" & stqtpj & "','',substr( '" & lotStr & "',2,instr( '" & lotStr & "','-R')-2))"

    
  AddSql2 (cmdStr2)
  AddSql (cmdStpj)
  End If
  
    
Next i
 rs.MoveNext
  Loop



  
  
  
  Else
  qbnum = "-B01"
  
  sqlDB = " select replace(a.customerpt,'.P2','') +','+ replace(a.customerlotid,'M','') +','+'1T'+ replace(a.customerlotid,'M','') +','+ replace(a.customerpt,'.P2','') +','+'1P'+ replace(a.customerpt,'.P2','') +','+min(a.podatecode) +','+min(a.podatecode)+',' " & _
" +max(a.htlotid)+'" & qbnum & "'+','+'S'+max(a.htlotid)+'" & qbnum & "' +','+rtrim(sum(qty)) +','+'Q'+rtrim(sum(qty)) +','" & _
" +min(a.htdatecode) +','+min(a.htdatecode) " & _
" from  [erpdata].[dbo].[TblTSV_Tray_details] a where trayqboxnumber in ('" & msgTxtTemp2 & "') " & _
" group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htlotid,a.htdatecode"

  If sqlDBRS.State = adStateOpen Then sqlDBRS.Close
    sqlDBRS.Open sqlDB, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText



'add 把 内箱号，几个Tray号，保存到一个表里
For i = 0 To iPos - 1
    lotStr = bid(i)
    
   stqtpj = Mid(qbnum, 3, 2)
    
    If lotStr <> "" Then
    
 cmdStr2 = " insert into [erpdata].[dbo].[TblTSV_INBOX_DETAILS](id,Containername,Subcontainername,Labeltype,customerpt,customerlotid,htlotid,podatecode,htdatecode,qty,flag,created_by,created_date,NHBox)" & _
 " select   1, 'S'+inboxno,trayboxno,typename,mpn_desc,jobnumber,htlotid,date_code,htdatecode,qty,'Y','" & useridTemp & "',getdate(),'" & strLSM & "' from ( " & _
 " select  max(a.htlotid) +'" & qbnum & "'  as inboxno,  '" & lotStr & "' as trayboxno , 'INQbox' as typename," & _
" a.customerpt as mpn_desc,a.customerlotid as jobnumber,max(a.htlotid) as htlotid ,min(a.podatecode) as date_code,min(a.htdatecode) as htdatecode,sum(qty) as qty" & _
" from  [erpdata].[dbo].[TblTSV_Tray_details] a where trayqboxnumber in ('" & bid(i) & "' ) " & _
" group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htdatecode ) X "
 
 cmdStpj = " insert into TSV_QBOXTBL_37SEQ(typename,createdate,seqtxt,containername,Firtname) " & _
          "  values ('INQbox',sysdate,'" & stqtpj & "','" & lotStr & "',substr( '" & lotStr & "',2,instr( '" & lotStr & "','-R')-2))"

    
  AddSql2 (cmdStr2)
  AddSql (cmdStpj)
  End If
  
    
Next i



 Dim dd As Integer
 dd = 0
If Not sqlDBRS.EOF Then
        Do While Not sqlDBRS.EOF
       
        pp = pp + 1
fileNameTemp = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & Format(Now(), "YYYYMMDDHHmmSS") & Trim(dd)
dirtemp = TxtDirInQbox.Text
Call addLabelTxt(fileNameTemp, sqlDBRS.Fields(0).Value, dirtemp)
  sqlDBRS.MoveNext
  Loop
    End If
  
  
  End If
  
   
InitData
iPos = 0

End Sub

Private Sub Command1_Click()

Dim beginTime As String
Dim endTime As String
Dim woTemp As String
Dim productTemp As String
Dim sqlTemp As String
Dim cusPTTemp As String




beginTime = Format(DTP1.Value, "YYYY/MM/DD")
endTime = Format(DTP2.Value, "YYYY/MM/DD")
cusPTTemp = CusPT.Text

 
  sqlTemp = " select  row_number() over(order by 1) as ""No."" , X.SubName as ""Sub Name"",X.ShipTo as ""Ship To"",X.CustomerDevice as ""Customer Device"",X.GCVersion as ""GC Version"",X.CSTID as ""CST ID"",X.CSTQTY as ""CST QTY"",X.BondPro as ""Bond Pro."",X.FABLotID as ""FAB Lot ID"",X.WaferID as ""Wafer ID"",X.GrossDies as ""Gross Dies"",X.PONO as ""PO NO"",X.WO as ""WO"",X.InvoiceNO as ""Invoice NO"",X.FABDevice as ""FAB Device"",X.PacklotID as ""Pack lot ID"",X.FABOutDate as ""FAB-Out Date"", " & _
 " X.SamplingQty as ""Sampling Qty"",X.PassDies as ""Pass Dies"",X.Yield as ""Yield"",X.Remark as ""Remark""  from ( " & _
 " select distinct 'HTKS' as SubName, 'GC_LG' as ShipTo, replace(a.mpn_desc,'-3','-2.5') as CustomerDevice, a.imager_customer_rev as GCVersion, " & _
        "   Get_GCWLA_LotID(b.lotid,b.substrateid,to_date('" + beginTime + "','YYYY/MM/DD'),to_date('" + endTime + "','YYYY/MM/DD'),'" + cusPTTemp + "') as CSTID,   Get_GCWLA_Qty(b.lotid,b.substrateid,to_date('" + beginTime + "','YYYY/MM/DD'),to_date('" + endTime + "','YYYY/MM/DD'),'" + cusPTTemp + "') as CSTQTY, 'SH' as BondPro, b.lotid as FABLotID,  substr(b.substrateid,length(b.substrateid)-1,2) as WaferID, b.passbincount as GrossDies, " & _
        " a.po_num as PONO,a.mtrl_num as WO,  '' InvoiceNO, a.fab_conv_id as FABDevice, c.firstname as PacklotID,to_char(sysdate, 'YYYY-MM-DD') as FABOutDate, " & _
        " b.passbincount as SamplingQty,  '' as PassDies, '' as Yield, '' as Remark " & _
        " from  tsv_qboxnumber_GC d, mappingdatatest b, customeroitbl_test a, container c " & _
        " Where d.create_date >=to_date('" + beginTime + "','YYYY/MM/DD') and  d.create_date <to_date('" + endTime + "','YYYY/MM/DD') and b.customershortname = 'GC' and b.substrateid =d.waferscribenumber and b.filename = a.id " & _
        " and a.customershortname = 'GC' and c.containername = b.substrateid and a.mpn_desc='" + cusPTTemp + "'  " & _
        " order by   b.lotid,  9 ) X"

 
     ExporToExcel (sqlTemp)



End Sub

Private Sub Command2_Click()
'ERP的导出


Dim billNoTemp As String

 billNoTemp = Trim(TxtBillNoGC.Text)
  
      sqlTemp = "  SELECT row_number() OVER(ORDER BY a.工单号,a.流程卡编号) AS [No.],'HTKS' as [Sub Name],'GC_LG' as [Ship To], " & _
" replace(b.MPN_DESC,'-3','-2.5') as [Customer Device],b.IMAGER_CUSTOMER_REV as [GC Version], " & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotID_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST ID]," & _
" [erpdata].[dbo].[Get_TSV_GCWLA_LotIDQty_New](a.单据编号,rtrim(ltrim(a.工单号)),rtrim(ltrim(a.流程卡编号))) as [CST QTY]," & _
" 'SH' as [Bond Pro.],a.工单号 as [FAB Lot ID],right(rtrim(ltrim(a.流程卡编号)),2) as [Wafer ID]," & _
" a.数量 as [Gross Dies],b.PO_NUM as [PO NO],b.MTRL_NUM as [WO],'' as [Invoice NO],b.FAB_CONV_ID as [FAB Device], " & _
" c.FIRSTNAME as [Pack lot ID],convert(varchar(10), getdate(), 126) as [FAB-Out Date],a.数量 as [Sampling Qty]," & _
" ''as [Pass Dies],''as [Yield],''as [Remark] " & _
" FROM   erpdata.dbo.tblStockMovesub a ,[ERPBASE].[dbo].[tblCustomerOI] b ,erpdata.dbo.TblQBOXNUMBER_TSV   c WHERE a.单据编号='" + billNoTemp + "'" & _
" and b.SOURCE_BATCH_ID=a.工单号 and c.WAFERSCRIBENUMBER=a.流程卡编号 and c.WAFERNUMBER=a.工单号 "
        
        
        
     SqlServerExporToExcel (sqlTemp)


End Sub

Public Sub CmdOKOut_Click()
'37 外箱

'把资料生成一个txt
Dim txtStr As String
Dim dirtemp As String
Dim cmdStr2 As String

Dim fileNameTemp As String
Dim msgTxtTemp As String
Dim msgTxtTemp2 As String
Dim qboxNoTemp As String
Dim qboxNoContainerTemp As String
Dim inBoxContainerTemp As String
Dim dnTemp As String
Dim sqlDB As String
Dim sqlHTDB As String

'Dim qboxNoTemp As String
Dim idTemp As Long

Dim sqlDBjob As String
Dim sqlDBRSS As New ADODB.Recordset

'调拨
Dim tray

If ComDN.Text = "" Then
     MsgBox "请先选择DN#", vbInformation, "友情提示"
Exit Sub

Else
dnTemp = Trim(ComDN.Text)

End If

fileNameTemp = ""
msgTxtTemp = ""

txtStr = TxtWaferIDOut.Text

msgTxtTemp = Replace(txtStr, vbCrLf, "','")

''1234,'456,'789'
msgTxtTemp2 = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & "," & Right(msgTxtTemp, Len(msgTxtTemp) - InStr(msgTxtTemp, ","))
 'msgTxtTemp2 = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 2) & "," & Right(msgTxtTemp, Len(msgTxtTemp) - InStr(msgTxtTemp, ","))


'msgTxtTemp2 = Replace(msgTxtTemp2, "SSB", "SB")

Dim bid
bid = Split(Replace(msgTxtTemp2, "'", "") & ",", ",")

Dim lotStr As String

For i = 0 To UBound(bid) - 1
    lotStr = bid(i)
     If lotStr <> "" Then
    
   '先判断是否在内箱中
    If Not Judge37InBoxIn(lotStr) Then
         MsgBox "此内箱：" & lotStr & " 不存在于内箱库中，不可以合外箱，请确认!", vbInformation, "友情提示"
        Exit Sub
    End If


        '先判断有没有装过外箱
    If Judge37ExistInBox(lotStr) Then
         MsgBox "此内箱：" & lotStr & " 已装过外箱，不可以重复装，请确认!", vbInformation, "友情提示"
         Exit Sub

    End If
    
    
    If i = 0 Then
    '第一个内箱号作为主批号
     inBoxContainerTemp = lotStr
    End If
    
    End If

Next i



 If Judge37DnNom(msgTxtTemp2, dnTemp) = False Then
 
  MsgBox "请确认选择的DN#是否正确！", vbInformation, "友情提示"
  Exit Sub
 
 End If


'Semtech外箱sql

sqlDB = Get37OutQboxTxt(msgTxtTemp2, qboxNoTemp, inBoxContainerTemp, ComDN.Text)

'sqlDBjob = "select ship.shiptoname+','+ship.shiptostreet1+','+ship.shiptostreet2+','+ship.shiptostreet3+','+ship.city+' '+ship.state+' '+ship.postalcode +','+" & _
'" ship.countrykey+','+ship.contactname+','+ship.phone+','+ship.delivery+','+'I'+ship.delivery +','+ship.purchasingdocno+','+'K'+ship.purchasingdocno +','+ " & _
'" ship.customerpartnumber+','+'P'+ship.customerpartnumber +','+a.customerpt+','+'Z'+a.customerpt+','+rtrim(sum(c.qty))+','+'Q'+rtrim(sum(c.qty)) +','+ " & _
'" ship.freightforwarder +','+c.CUSTOMERLOTID +','+'' +','+'' +','+'COO:CHINA' +','+'CHINA'  " & _
'" from [ERPBASE].[dbo].[tblCustomerShippingUp] ship ,[erpdata].[dbo].[TblTSV_INBOX_DETAILS] a ,[erpdata].[dbo].[TblTSV_Tray_details]  c  " & _
'" where a.labeltype='INQbox' and a.containername in ('" & msgTxtTemp2 & "') and ship.batchnumber=c.customerlotid   and c.TRAYQBOXNUMBER=a.SUBCONTAINERNAME   and ship.delivery = '" & ComDN.Text & " '" & _
'" Group By ship.shiptoname,ship.shiptostreet1,ship.shiptostreet2,ship.shiptostreet3,ship.city,ship.state,ship.postalcode , " & _
'" ship.countrykey,ship.contactname,ship.phone,ship.delivery,'I'+ship.delivery ,ship.purchasingdocno, ship.customerpartnumber,c.CUSTOMERLOTID,a.customerpt,ship.freightforwarder  "
'

sqlDBjob = "select ship.delivery  + ',' +'I'+ship.delivery + ','+ left(ship.purchasingdocno,10) + ','+ 'K' + left(ship.purchasingdocno,10) + ','+ 'E2'+','+ left(ship.CustomerPartNumber,11)+','+'P' + left(ship.CustomerPartNumber,11)+','+ ship.MarketingPN +','+ 'Z'+ ship.MarketingPN +','+ rtrim(sum(c.qty))+ ','+ 'Q' + rtrim(sum(c.qty))+ ','+ ship.FreightForwarder+','+'CHINA'+','+ " & _
" left(SHIP.ShipToName,33) +','+ship.ShipToStreet1+','+ ship.ShipToStreet2+','+ ship.ShipToStreet3 +','+ ship.City+' '+ship.State + ' '+ship.PostalCode + ',' + ship.CountryKey+ ','+ship.Phone + ','+ replace(c.CUSTOMERLOTID,'M','') + ',' +'P'+ replace(c.CUSTOMERLOTID,'M','') + ',' + c.PODATECODE + ','+ '9D' + c.PODATECODE  from ERPBASE . dbo . tblCustomerShippingUp ship, erpdata . dbo . TblTSV_INBOX_DETAILS a, erpdata . dbo . TblTSV_Tray_details c " & _
" where a.labeltype = 'INQbox' and a.containername in ('" & msgTxtTemp2 & "') and ship.batchnumber = c.customerlotid and c.TRAYQBOXNUMBER = a.SUBCONTAINERNAME and ship.delivery = '" & ComDN.Text & " '  Group By ship.shiptoname, ship.MarketingPN , ship.shiptostreet1, ship.shiptostreet2, ship.shiptostreet3, ship.city, ship.state, ship.postalcode, " & _
 " ship.countrykey, ship.contactname,ship.phone,ship.delivery,'I' + ship.delivery, ship.purchasingdocno, ship.customerpartnumber,  a.customerpt, ship.freightforwarder, replace(c.CUSTOMERLOTID,'M',''),c.PODATECODE "


 If sqlDBRSS.State = adStateOpen Then sqlDBRS.Close
    sqlDBRSS.Open sqlDBjob, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    
Dim ppj As Integer
Dim sanxing As String
sanxing = "\\10.160.1.14\BarCode\37\37BoxOut\"
 ppj = 0
If Not sqlDBRSS.EOF Then
        Do While Not sqlDBRSS.EOF
       
        ppj = ppj + 1
fileNameTemp = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & Format(Now(), "YYYYMMDDHHmmSS") & Trim(ppj)
dirtemp = sanxing
Call addLabelTxt(fileNameTemp, sqlDBRSS.Fields(0).Value, dirtemp)
  sqlDBRSS.MoveNext
  Loop
    End If

'HT外箱sql
sqlHTDB = Get37OutQboxHTTxt(msgTxtTemp2, qboxNoTemp, inBoxContainerTemp)

'add 把 数据保存到外箱表里
For i = 0 To UBound(bid) - 1
    lotStr = bid(i)
    
    If lotStr <> "" Then
    
' cmdStr2 = " insert into TSV_InBox_details(id,Containername,Subcontainername,Labeltype,customerpt,customerlotid,htlotid,podatecode,htdatecode,qty,flag,created_by,created_date)" & _
' " select   InBox37_SEQ.nextval, inboxno,trayboxno,typename,mpn_desc,jobnumber,htlotid,date_code,htdatecode,qty,'Y','" & useridTemp & "',sysdate from ( " & _
' " select  max(a.htlotid) || get_37_LableID('INQbox','" & inBoxContainerTemp & "', max(a.htlotid))  as inboxno,  '" & lotStr & "' as trayboxno , 'INQbox' as typename," & _
'" a.customerpt as mpn_desc,a.customerlotid as jobnumber,max(a.htlotid) as htlotid ,min(a.podatecode) as date_code,min(a.htdatecode) as htdatecode,sum(qty) as qty" & _
'" from  TSV_Tray_details a where trayqboxnumber in ('" & msgTxtTemp2 & "' ) " & _
'" group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htdatecode ) "


 cmdStr2 = " insert into [erpdata].[dbo].[TblTSV_OutBOX_DETAILS]([ID],[CONTAINERNAME],[SUBCONTAINERNAME],[LABELTYPE],[TRAYQBOXNUMBER] " & _
" ,[ShipToName],[ShipToStreet1] ,[ShipToStreet2],[ShipToStreet3],[ShipToStreet4]" & _
" ,[CounTryKey],[ContactName],[Phone],[Invoice],[PONo] " & _
" ,[CustomerPT],[MFGPT],[Qty],[Forwarder] " & _
" ,[COO],[FLAG],[CREATED_BY],[CREATED_DATE]) " & _
" select  1,'" & sqlHTDB & "','" & lotStr & "','OutQbox','" & dnTemp & "', " & _
" ship.shiptoname,ship.shiptostreet1,ship.shiptostreet2,ship.shiptostreet3,ship.city+' '+ship.[state]+' '+ship.postalcode as shiptostreet4," & _
" ship.countrykey,ship.contactname,ship.phone,ship.delivery ,ship.purchasingdocno," & _
" ship.customerpartnumber,a.customerpt,sum( CAST(a.qty AS numeric(18,0))),ship.freightforwarder," & _
" 'CHINA','Y','" & useridTemp & "',getdate() " & _
" from [ERPBASE].[dbo].[tblCustomerShippingUp] ship ,[erpdata].[dbo].[TblTSV_INBOX_DETAILS]a ,erpdata . dbo . TblTSV_Tray_details c " & _
" where a.labeltype='INQbox' and a.containername in ('" & bid(i) & "') and  ship.batchnumber = c.CUSTOMERLOTID  and c.TRAYQBOXNUMBER = a.SUBCONTAINERNAME and ship.delivery = '" & ComDN.Text & "'" & _
" Group By ship.shiptoname,ship.shiptostreet1,ship.shiptostreet2,ship.shiptostreet3,ship.city,ship.[state],ship.postalcode ," & _
" ship.countrykey,ship.contactname,ship.phone,ship.delivery,'I'+ship.delivery ,ship.purchasingdocno,'K'+ship.purchasingdocno ," & _
" ship.customerpartnumber,'P'+ship.customerpartnumber ,a.customerpt,'Z'+a.customerpt,ship.freightforwarder"
 

  AddSql2 (cmdStr2)
  End If
  
    
Next i

'标签txt begnin------------------
'Semtech Qbox txt
fileNameTemp = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1)
dirtemp = TxtDirOutQbox.Text
Call addLabelTxt(fileNameTemp, sqlDB, dirtemp)

'Semtech HTQbox txt
fileNameTemp = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1)
dirtemp = TxtDirOutHTQbox.Text
Call addLabelTxt(fileNameTemp, sqlHTDB, dirtemp)

qboxNoTemp = sqlHTDB

'标签txt end------------------


'调拨   begin --------------
'抓出所有Tray盘号

'Set billLotTemp = GetDiaoBoList(msgTxtTemp2)
'If (billLotTemp.RecordCount > 0) Then
'    '循环有多少个Tray
'
'    Do While Not billLotTemp.EOF
'        lotIDTemp = billLotTemp.fields("waferlot").Value
'        productTemp = billLotTemp.fields("productname").Value
'        qtyTemp = CLng(billLotTemp.fields("qty").Value)
'
'        erpdate = Format(CDate(billLotTemp.fields("erpcreationdate").Value), "YYYY-MM-DD")
'
'        woDeptIDTemp = billLotTemp.fields("deptid").Value
'
'
'
'          '-----begin------
'
'         Set adoCmd = New ADODB.Command
'         Set adoCmd.ActiveConnection = INIadoCon2
'             adoCmd.CommandText = "uspPMC_XDInterface"
'             adoCmd.Parameters.Refresh
'             adoCmd.CommandType = adCmdStoredProc
'
'          Set adoprm1 = New ADODB.Parameter   '工单号
'          adoprm1.Type = adChar
'          adoprm1.Size = 20
'          adoprm1.Direction = adParamInput
'          adoprm1.Value = lotIDTemp
'          adoCmd.Parameters.Append adoprm1
'
'          Set adoprm2 = New ADODB.Parameter   '料号
'          adoprm2.Type = adChar
'          adoprm2.Size = 20
'          adoprm2.Direction = adParamInput
'          adoprm2.Value = productTemp
'          adoCmd.Parameters.Append adoprm2
'
'          Set adoprm3 = New ADODB.Parameter   '数量
'          adoprm3.Type = adInteger
'          adoprm3.Direction = adParamInput
'          adoprm3.Value = qtyTemp
'          adoCmd.Parameters.Append adoprm3
'
'            Set adoprm4 = New ADODB.Parameter   '时间
'
'          adoprm4.Type = adChar
'         adoprm4.Size = 20
'          adoprm4.Direction = adParamInput
'          adoprm4.Value = erpdate
'          adoCmd.Parameters.Append adoprm4
'
'
'            Set adoprm5 = New ADODB.Parameter   '线别
'          adoprm5.Type = adInteger
'          adoprm5.Direction = adParamInput
'          adoprm5.Value = 1
'          adoCmd.Parameters.Append adoprm5
'
'          Set adoprm6 = New ADODB.Parameter   '部门id
'          adoprm6.Type = adChar
'          adoprm6.Size = 30
'          adoprm6.Direction = adParamInput
'          adoprm6.Value = woDeptIDTemp
'          adoCmd.Parameters.Append adoprm6
'
'
'
'
'          adoCmd.Execute
'
'
'        billLotTemp.MoveNext
'
'    Loop
'
'End If

'调拨   end --------------

'2016-06-23 add 根据大箱号来更新Sql server

Dim qtyTemp As Long

qtyTemp = Get37BigQboxQty(qboxNoTemp)

 sqlTemp = "insert into [erpdata].[dbo].[tblPackMainInf](箱号,客户代码,数量,合格标记,装箱标记,产线标记) " & _
"  values('" & qboxNoTemp & "','37'," & qtyTemp & ",'0','1','1') "

AddSql2 (sqlTemp)

  '插入Sqlserver   tblPackTreeInf

sqlTemp = "insert into [erpdata].[dbo].[tblPackTreeInf](箱号,上级序号,基层标记 ,Memo) values('" & qboxNoTemp & "',0,1,'37')"
AddSql2 (sqlTemp)

'再更新小箱的上级序号

'把序号先查出来，再整体更新

idTemp = Get37BigQboxIDV1(qboxNoTemp)
boxdn = ComDN.Text

sqlTemp = "insert into [erpdata].[dbo].[tblStockNumTree](序号,箱号,上级序号,基层标记 ,尺寸,重量,Memo,DN) values( " & idTemp & ",'" & qboxNoTemp & "',0,1,'','','37','" & boxdn & "')"
AddSql2 (sqlTemp)

sqlTemp = " Update [erpdata].[dbo].[tblPackTreeInf] set 上级序号='" & idTemp & "',Memo='37' " & _
" where 箱号 in ( select c.箱号 from [erpdata].[dbo].[TblTSV_OutBOX_DETAILS] a ,[erpdata].[dbo].[TblTSV_INBOX_DETAILS] b ,[erpdata].[dbo].[tblPackTreeInf] c " & _
"Where b.CONTAINERNAME = a.SUBCONTAINERNAME and c.箱号=b.SUBCONTAINERNAME and a.CONTAINERNAME='" & qboxNoTemp & "') "

AddSql2 (sqlTemp)

sqlTemp = " Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & idTemp & "',Memo='37' " & _
" where 箱号 in ( select b.箱号 from [erpdata].[dbo].[tblPackTreeInf] a  , [erpdata].[dbo].[tblPackTreeInf] b " & _
" where a.箱号='" & qboxNoTemp & "' and b.上级序号=a.序号) "

AddSql2 (sqlTemp)

TxtWaferIDIn.Text = ""
TxtWaferIDIn.SetFocus

End Sub

Private Sub Form_Activate()
txtScan.SetFocus

End Sub

Private Sub Form_Load()

useridTemp = UCase(gUserName)

txtFlag = "ZHULIHUA"
TxtWaferIDIn.Text = ""
TxtWaferIDOut.Text = ""
TxtDirInQbox.Text = "\\10.160.1.14\BarCode\37\37内箱\"
TxtDirOutQbox.Text = "\\10.160.1.14\BarCode\37\37外箱\"
TxtDirOutHTQbox.Text = "\\10.160.1.14\BarCode\37\37BOX\"

InitData
iPos = 0

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If PreviousTab = 0 Then
TxtWaferIDOut.SetFocus

ElseIf PreviousTab = 1 Then
TxtWaferIDIn.SetFocus
End If

End Sub

Private Sub InitScanData()

'初始化扫描界面
TxtWaferIDIn.ForeColor = vbBlue
TxtWaferIDIn.Text = "卷盘1: " & sTrayNo(0) & vbCrLf & "卷盘2: " & sTrayNo(1) & vbCrLf & "卷盘3: " & sTrayNo(2) & vbCrLf & "卷盘4: " & sTrayNo(3) & vbCrLf & "卷盘5: " & sTrayNo(4) & vbCrLf & "卷盘6: " & sTrayNo(5) & vbCrLf & "卷盘7: " & sTrayNo(6) & vbCrLf & "卷盘8: " & sTrayNo(7) & vbCrLf & "卷盘9: " & sTrayNo(8) & vbCrLf

End Sub

Private Sub InitData()

For iPos = 0 To 8

    sTrayNo(iPos) = ""
    
Next

'初始化扫描界面
TxtWaferIDIn.ForeColor = vbRed
TxtWaferIDIn.Text = "卷盘1:准备中" & vbCrLf & "卷盘2:准备中" & vbCrLf & "卷盘3:准备中" & vbCrLf & "卷盘4:准备中" & vbCrLf & "卷盘5:准备中" & vbCrLf & "卷盘6:准备中" & vbCrLf & "卷盘7:准备中" & vbCrLf & "卷盘8:准备中" & vbCrLf & "卷盘9:准备中" & vbCrLf

End Sub

Private Sub CheckStatus()

If iPos = 9 Then
    
    CmdOK_Click
    InitData
    iPos = 0
    
End If

End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)

' 扫描结束触发
If KeyAscii <> 13 Then
    Exit Sub
End If

' 特殊符号则确定
If txtScan.Text = "ZHULIHUA" Then
    
    CmdOK_Click
    InitData
    iPos = 0
    Exit Sub
End If

'赋值

' 相同则不赋新值
If iPos > 0 And iPos < 9 Then
    For i = 0 To iPos
        If txtScan.Text = sTrayNo(i) Then
            
            If txtScan.Text <> "" Then
                MsgBox "请不要重复扫描同一lot: " & txtScan.Text, vbInformation
            End If
        
            txtScan.Text = ""
            Exit Sub
        End If
    Next
End If

' 赋值
If (Len(txtScan.Text) = 16 Or Len(txtScan.Text) = 17) And (Left$(txtScan.Text, 1) = "S") Then
    sTrayNo(iPos) = txtScan.Text
    iPos = iPos + 1
    
    InitScanData

End If

txtScan.Text = ""

CheckStatus

End Sub

Private Sub TxtWaferIDOut_KeyPress(KeyAscii As Integer)

' 扫描结束触发
If KeyAscii <> 13 Then
    Exit Sub
End If







End Sub
