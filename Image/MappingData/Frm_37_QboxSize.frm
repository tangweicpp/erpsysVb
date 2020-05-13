VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Frm_37_QboxSize 
   Caption         =   "Semtech箱子重量与尺寸维护"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
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
   ScaleHeight     =   8430
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox ComDN 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   480
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1200
      Width           =   3375
   End
   Begin VB.ComboBox CmbQbox 
      Height          =   315
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   480
      Left            =   5640
      TabIndex        =   10
      Top             =   2880
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "提交"
      Height          =   480
      Left            =   3480
      TabIndex        =   9
      Top             =   2880
      Width           =   990
   End
   Begin VB.TextBox TxtQboxTemp 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   1680
      Width           =   3375
   End
   Begin MSDataListLib.DataCombo DCbMainItem 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo CmbQboxType 
      Height          =   315
      Left            =   10320
      TabIndex        =   6
      Top             =   1080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "箱子重量(KG)："
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label LblTye 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户："
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "成品料号："
      Height          =   195
      Left            =   4800
      TabIndex        =   4
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "箱子尺寸(CM)："
      Height          =   195
      Left            =   9120
      TabIndex        =   3
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "箱号："
      Height          =   195
      Left            =   5040
      TabIndex        =   1
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DN#："
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   510
   End
End
Attribute VB_Name = "Frm_37_QboxSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim useridTemp As String
Dim DNBatch  As New ADODB.Recordset

Private Sub CmdExit_Click()
TxtWaferID.Text = ""
TxtWaferID.SetFocus
End Sub




Private Sub CmdExitIn_Click()
TxtWaferIDIn.Text = ""
TxtWaferIDIn.SetFocus

End Sub

Private Sub CmdExitOut_Click()
TxtWaferIDOut.Text = ""
TxtWaferIDOut.SetFocus
End Sub

Private Sub CmdOK_Click()
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


Dim sqlDB As String

fileNameTemp = ""
msgTxtTemp = ""

txtStr = TxtWaferIDIn.Text

msgTxtTemp = Replace(txtStr, vbCrLf, "','")

''1234,'456,'789'
msgTxtTemp2 = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & "," & Right(msgTxtTemp, Len(msgTxtTemp) - InStr(msgTxtTemp, ","))


Dim bid
bid = Split(Replace(msgTxtTemp2, "'", "") & ",", ",")

Dim lotStr As String

For i = 0 To UBound(bid) - 1
    lotStr = bid(i)
    
    If lotStr <> "" Then
    
    '先判断是否在仓库
    If Not Judge37TrayIn(lotStr) Then
         MsgBox "此卷：" & lotStr & " 不存在于ERP仓库中，不可以合内箱，请确认!", vbInformation, "友情提示"
    
    End If
    
    
        '先判断有没有装过
    If Judge37ExistInBox(lotStr) Then
         MsgBox "此卷：" & lotStr & " 已装过内箱，不可以重复装，请确认!", vbInformation, "友情提示"
    
    End If
    
    
    
    If i = 0 Then
     inBoxContainerTemp = lotStr
    End If
    
    End If


Next i


'add 先生成内箱箱号

qboxNoSeqTemp = Get37InQboxSeqTxt(msgTxtTemp2, qboxNoTemp, inBoxContainerTemp)


'sqlDB = Get37InQboxTxt(msgTxtTemp2, qboxNoTemp, inBoxContainerTemp, qboxNoSeqTemp)

'add 把 内箱号，几个Tray号，保存到一个表里

For i = 0 To UBound(bid) - 1
    lotStr = bid(i)
    
    If lotStr <> "" Then
    
 cmdStr2 = " insert into [erpdata].[dbo].[TblTSV_INBOX_DETAILS](id,Containername,Subcontainername,Labeltype,customerpt,customerlotid,htlotid,podatecode,htdatecode,qty,flag,created_by,created_date)" & _
 " select   1, inboxno,trayboxno,typename,mpn_desc,jobnumber,htlotid,date_code,htdatecode,qty,'Y','" & useridTemp & "',getdate() from ( " & _
 " select  max(a.htlotid) +'" & qboxNoSeqTemp & "'  as inboxno,  '" & lotStr & "' as trayboxno , 'INQbox' as typename," & _
" a.customerpt as mpn_desc,a.customerlotid as jobnumber,max(a.htlotid) as htlotid ,min(a.podatecode) as date_code,min(a.htdatecode) as htdatecode,sum(qty) as qty" & _
" from  [erpdata].[dbo].[TblTSV_Tray_details] a where trayqboxnumber in ('" & msgTxtTemp2 & "' ) " & _
" group by a.customerpt ,a.customerlotid ,a.customerlotid ,a.customerpt ,a.podatecode ,a.htdatecode ) X "
    
  AddSql2 (cmdStr2)
  End If
  
    
Next i


fileNameTemp = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1)

dirtemp = TxtDirInQbox.Text


Call addLabelTxt(fileNameTemp, sqlDB, dirtemp)
TxtWaferIDIn.Text = ""
TxtWaferIDIn.SetFocus

End Sub






Private Sub Command1_Click()
Dim dnTemp As String
Dim qboxTemp As String
Dim customerTemp As String
Dim ptTemp  As String
Dim qboxSizeTemp As String
Dim weightTemp As String
Dim sqlTemp As String
Dim idTemp As Long
Dim qtyTemp As Long
Dim DNNTemp As String


qboxTemp = CmbQbox.Text


DNNTemp = ComDN.Text


qboxSizeTemp = CmbQboxType.Text
weightTemp = TxtQboxTemp.Text

If qboxTemp = "" Or qboxSizeTemp = "" Or weightTemp = "" Then
    MsgBox "箱号，尺寸，重量不可以为空，请确认!", vbInformation, "友情提示"
    Exit Sub
End If

'插入Sqlserver



'
''插入Sqlserver   装箱记录
'
'qtyTemp = Get37BigQboxQty(qboxTemp)
'
'
' sqlTemp = "insert into [erpdata].[dbo].[tblPackMainInf](箱号,客户代码,数量,合格标记,装箱标记,产线标记) " & _
'"  values('" & qboxTemp & "','37'," & qtyTemp & ",'0','1','1') "
'
'AddSql2 (sqlTemp)
'
'
'
'  '插入Sqlserver   tblPackTreeInf
'
'sqlTemp = "insert into [erpdata].[dbo].[tblPackTreeInf](箱号,上级序号,基层标记 ,Memo) values('" & qboxTemp & "',0,1,'37')"
'AddSql2 (sqlTemp)
'
''再更新小箱的上级序号
'
''把序号先查出来，再整体更新
'
idTemp = Get37BigQboxIDV1(qboxTemp)
'
'
'sqlTemp = " Update [erpdata].[dbo].[tblPackTreeInf] set 上级序号='" & idTemp & "',Memo='37' " & _
'" where 箱号 in ( select c.箱号 from [erpdata].[dbo].[TblTSV_OutBOX_DETAILS] a ,[erpdata].[dbo].[TblTSV_INBOX_DETAILS] b ,[erpdata].[dbo].[tblPackTreeInf] c " & _
'"Where b.CONTAINERNAME = a.SUBCONTAINERNAME and c.箱号=b.SUBCONTAINERNAME and a.CONTAINERNAME='" & qboxTemp & "') "
'
'AddSql2 (sqlTemp)
'
'
'

'插入Sqlserver   [tblStockNumTree]

sqlTemp = "update  [erpdata].[dbo].[tblStockNumTree]  set 尺寸 ='" & qboxSizeTemp & "',重量= '" & weightTemp & "',DN ='" & DNNTemp & "' where  箱号='" & qboxTemp & "'"
AddSql2 (sqlTemp)

'再更新小箱的上级序号

'把序号先查出来，再整体更新


'sqlTemp = " Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & idTemp & "',Memo='37' " & _
'" where 箱号 in ( select c.箱号 from [erpdata].[dbo].[TblTSV_OutBOX_DETAILS] a ,[erpdata].[dbo].[TblTSV_INBOX_DETAILS] b ,[erpdata].[dbo].[tblStockNumTree] c " & _
'"Where b.CONTAINERNAME = a.SUBCONTAINERNAME and c.箱号=b.SUBCONTAINERNAME and a.CONTAINERNAME='" & qboxTemp & "') "

'sqlTemp = " Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & idTemp & "',Memo='37' " & _
'" where 箱号 in ( select c.箱号 from [erpdata].[dbo].[TblTSV_OutBOX_DETAILS] a ,[erpdata].[dbo].[TblTSV_INBOX_DETAILS] b ,[erpdata].[dbo].[tblStockNumTree] c " & _
'"Where b.CONTAINERNAME = a.SUBCONTAINERNAME and c.箱号=b.SUBCONTAINERNAME and a.CONTAINERNAME='" & qboxTemp & "') "

'
'sqlTemp = " Update [erpdata].[dbo].[tblStockNumTree] set 上级序号='" & idTemp & "',Memo='37' " & _
'" where 箱号 in ( select b.箱号 from [erpdata].[dbo].[tblPackTreeInf] a  , [erpdata].[dbo].[tblPackTreeInf] b " & _
'" where a.箱号='" & qboxTemp & "' and b.上级序号=a.序号) "


'AddSql2 (sqlTemp)

MsgBox "提交成功!", vbInformation, "友情提示"

End Sub

Private Sub Command2_Click()
Unload Me


End Sub

Private Sub CmdOKOut_Click()
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


'调拨
Dim tray


If cmbDN.Text = "" Then
     MsgBox "请先选择DN#", vbInformation, "友情提示"
Exit Sub

Else
dnTemp = Trim(cmbDN.Text)

End If




fileNameTemp = ""
msgTxtTemp = ""

txtStr = TxtWaferIDOut.Text

msgTxtTemp = Replace(txtStr, vbCrLf, "','")

''1234,'456,'789'
msgTxtTemp2 = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & "," & Right(msgTxtTemp, Len(msgTxtTemp) - InStr(msgTxtTemp, ","))


Dim bid
bid = Split(Replace(msgTxtTemp2, "'", "") & ",", ",")

Dim lotStr As String

For i = 0 To UBound(bid) - 1
    lotStr = bid(i)
     If lotStr <> "" Then
    
   '先判断是否在内箱中
    If Not Judge37InBoxIn(lotStr) Then
         MsgBox "此内箱：" & lotStr & " 不存在于内箱库中，不可以合外箱，请确认!", vbInformation, "友情提示"
    
    End If
    
    
        '先判断有没有装过外箱
    If Judge37ExistInBox(lotStr) Then
         MsgBox "此内箱：" & lotStr & " 已装过外箱，不可以重复装，请确认!", vbInformation, "友情提示"
    
    End If
    
    
    If i = 0 Then
    '第一个内箱号作为主批号
     inBoxContainerTemp = lotStr
    End If
    
    End If

Next i

'Semtech外箱sql

sqlDB = Get37OutQboxTxt(msgTxtTemp2, qboxNoTemp, inBoxContainerTemp, ComDN.Text)

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
" from [ERPBASE].[dbo].[tblCustomerShippingUp] ship ,[erpdata].[dbo].[TblTSV_INBOX_DETAILS]a " & _
" where a.labeltype='INQbox' and a.containername in ('" & msgTxtTemp2 & "') and ship.batchnumber=a.customerlotid " & _
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



TxtWaferIDOut.Text = ""
TxtWaferIDOut.SetFocus



End Sub

'Private Sub IniDNList()
'Set mainItemRS = Get37Dn()
'Set CmbDN.RowSource = mainItemRS
'CmbDN.ListField = mainItemRS("DNName").Name
'CmbDN.BoundColumn = mainItemRS("DNID").Name

'End Sub


'Private Sub IniQboxList()
'Set mainItemRS = Get37QboxList()
'Set CmbQbox.RowSource = mainItemRS
'CmbQbox.ListField = mainItemRS("DNName").Name
'CmbQbox.BoundColumn = mainItemRS("DNID").Name




'End Sub



Private Sub IniCustomerList()
Set mainItemRS = Get37CustomerList()
Set DCbMainItem.RowSource = mainItemRS
DCbMainItem.ListField = mainItemRS("DNName").Name
DCbMainItem.BoundColumn = mainItemRS("DNID").Name

End Sub



'Private Sub IniPTList()
'Set mainItemRS = Get37PTList()
'Set DCbChildItem.RowSource = mainItemRS
'DCbChildItem.ListField = mainItemRS("DNName").Name
'DCbChildItem.BoundColumn = mainItemRS("DNID").Name

'End Sub

Private Sub IniQboxTypeList()
Set mainItemRS = Get37QboxTypeList()
Set CmbQboxType.RowSource = mainItemRS
CmbQboxType.ListField = mainItemRS("DNName").Name
CmbQboxType.BoundColumn = mainItemRS("DNID").Name


End Sub



'Private Sub DCbChildItem_Click(Area As Integer)
'If Len(CmbQbox.Text) = 0 Then
'MsgBox "请选择箱号"
'Exit Sub
'Else
'Set mainItemRS = Get37PTList(CmbDN.Text)
'Set DCbChildItem.RowSource = mainItemRS
'DCbChildItem.ListField = mainItemRS("DNName").Name
'DCbChildItem.BoundColumn = mainItemRS("DNID").Name
'End If


'End Sub

Private Sub ComDN_DropDown()

Call InitCtrl2

End Sub

Private Sub CmbQbox_DropDown()

Call InitCtrl(ComDN.Text)

End Sub

Private Sub Combo2_DropDown()

Call InitCtrl1(ComDN.Text)

End Sub

Private Sub InitCtrl1(tempDN As String)
Dim i                   As Integer
Dim strSql              As String
Dim Rs                  As New ADODB.Recordset
    
    '加载单据类型
    strSql = "SELECT DISTINCT RTRIM(c.料号) PRODUCT " & _
             " FROM erpbase..tblCustomerShippingUp a " & _
             " INNER JOIN erpdata..TblTSV_Tray_details b ON a.BatchNumber=b.CUSTOMERLOTID " & _
             " INNER JOIN erpdata..tblPackMainInfSub c ON b.TRAYQBOXNUMBER=c.箱号 AND b.LOTNUM=c.工单号 " & _
             " WHERE a.Delivery='" & tempDN & "' "
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    Combo2.Clear
    If Not Rs.EOF Then
        Do While Not Rs.EOF
            Combo2.AddItem Trim$("" & Rs!product)
            Rs.MoveNext
        Loop
        Combo2.ListIndex = 0
    End If
    Rs.Close
   
End Sub

Private Sub InitCtrl(tempDN As String)
Dim i                   As Integer
Dim strSql              As String
Dim Rs                  As New ADODB.Recordset
    
    Combo2.Clear
    '加载单据类型
    strSql = "SELECT DISTINCT erpdata.dbo.f_getbigparent(b.TRAYQBOXNUMBER) 箱号 " & _
             " FROM erpbase..tblCustomerShippingUp a " & _
             " INNER JOIN erpdata..TblTSV_Tray_details b ON a.BatchNumber=b.CUSTOMERLOTID " & _
             " WHERE a.Delivery='" & tempDN & "' " & _
             " AND erpdata.dbo.f_getbigparent(b.TRAYQBOXNUMBER)<>'' "
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, INIadoCon, adOpenStatic, adLockReadOnly, adCmdText
    CmbQbox.Clear
    If Not Rs.EOF Then
        Do While Not Rs.EOF
            CmbQbox.AddItem Trim$("" & Rs!箱号)
            Rs.MoveNext
        Loop
        CmbQbox.ListIndex = 0
    End If
    Rs.Close
   
End Sub

Private Sub InitCtrl2()
Dim i                   As Integer
Dim strSql              As String
Dim Rs                  As New ADODB.Recordset
    
    ComDN.Clear
    '加载单据类型
    strSql = "select distinct delivery  from   CUSTOMERSHIPPINGUPTBL  a where a.flag='Y' order by a.delivery desc "
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    CmbQbox.Clear
    If Not Rs.EOF Then
        Do While Not Rs.EOF
            ComDN.AddItem Trim$("" & Rs!Delivery)
            Rs.MoveNext
        Loop
        ComDN.ListIndex = 0
    End If
    Rs.Close
   
End Sub

'Private Sub CmbDN_Click(Area As Integer)
'CmbQbox.Clear
'Combo2.Clear
'Set mainItemRS = Get37Dn()
'Set CmbDN.RowSource = mainItemRS
'CmbDN.ListField = mainItemRS("DNName").Name
'CmbDN.BoundColumn = mainItemRS("DNID").Name
'
'
'End Sub

Private Sub Form_Load()

useridTemp = UCase(gUserName)


'IniDNList

'IniQboxList

IniCustomerList

'IniPTList

IniQboxTypeList


End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If PreviousTab = 0 Then
TxtWaferIDOut.SetFocus

ElseIf PreviousTab = 1 Then
TxtWaferIDIn.SetFocus
End If



End Sub



Private Sub TxtQboxTemp_KeyPress(KeyAscii As Integer)
Dim Numbers As String
Numbers = "1234567890" + Chr(8) + Chr(46) + Chr(13)
If InStr(Numbers, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If
End Sub
