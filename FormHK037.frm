VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FormHK037 
   Caption         =   "Form1"
   ClientHeight    =   9315
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14415
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
   ScaleHeight     =   9315
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SST 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   16325
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "内箱"
      TabPicture(0)   =   "FormHK037.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TxtDirInQbox"
      Tab(0).Control(1)=   "cmd(3)"
      Tab(0).Control(2)=   "cmd(2)"
      Tab(0).Control(3)=   "txtText1(1)"
      Tab(0).Control(4)=   "lblTxt"
      Tab(0).Control(5)=   "lbl(1)"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "外箱"
      TabPicture(1)   =   "FormHK037.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lbl(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblTxt11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtText2(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtText3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdout1(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmd(5)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FormHK037.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmd 
         Caption         =   "取消"
         Height          =   480
         Index           =   5
         Left            =   4800
         TabIndex        =   20
         Top             =   7920
         Width           =   1455
      End
      Begin VB.CommandButton cmdout1 
         Caption         =   "确定"
         Height          =   480
         Index           =   4
         Left            =   840
         TabIndex        =   19
         Top             =   7920
         Width           =   1455
      End
      Begin VB.TextBox txtText3 
         Height          =   405
         Left            =   1560
         TabIndex        =   18
         Top             =   840
         Width           =   7095
      End
      Begin VB.TextBox txtText2 
         Height          =   5775
         Index           =   2
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "FormHK037.frx":0054
         Top             =   1680
         Width           =   9255
      End
      Begin VB.TextBox TxtDirInQbox 
         Height          =   375
         Left            =   -73560
         TabIndex        =   8
         Top             =   600
         Width           =   6855
      End
      Begin VB.CommandButton cmd 
         Caption         =   "取消"
         Height          =   480
         Index           =   3
         Left            =   -70320
         TabIndex        =   7
         Top             =   8400
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         Caption         =   "确定"
         Height          =   480
         Index           =   2
         Left            =   -72960
         TabIndex        =   6
         Top             =   8400
         Width           =   1455
      End
      Begin VB.TextBox txtText1 
         Height          =   6615
         Index           =   1
         Left            =   -74520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "FormHK037.frx":005C
         Top             =   1440
         Width           =   9255
      End
      Begin VB.CommandButton cmd 
         Caption         =   "取消"
         Height          =   480
         Index           =   1
         Left            =   -70320
         TabIndex        =   4
         Top             =   8520
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         Caption         =   "确定"
         Height          =   480
         Index           =   0
         Left            =   -72960
         TabIndex        =   3
         Top             =   8520
         Width           =   1455
      End
      Begin VB.TextBox txtText1 
         Height          =   6615
         Index           =   0
         Left            =   -74520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "FormHK037.frx":0062
         Top             =   1800
         Width           =   9255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -72600
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label lblTxt11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Txt路径："
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫入的小箱："
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label lblTxt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Txt路径："
         Height          =   195
         Left            =   -74400
         TabIndex        =   14
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫入的卷盘："
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   13
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lblSemtechTxt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semtech外箱Txt路径："
         Height          =   195
         Left            =   -74400
         TabIndex        =   12
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫入的内箱："
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   11
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label lblHTTxt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HT外箱Txt路径："
         Height          =   195
         Left            =   -68040
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblDN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN#："
         Height          =   195
         Left            =   -73200
         TabIndex        =   9
         Top             =   1200
         Width           =   510
      End
   End
End
Attribute VB_Name = "FormHK037"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
useridTemp = UCase(gUserName)



txtText1(1).Text = ""
txtText2(2).Text = ""
'TxtWaferIDOut.Text = ""
TxtDirInQbox.Text = "\\10.160.1.14\BarCode\HK037\HK037INBOX\"
txtText3.Text = "\\10.160.1.14\BarCode\HK037\HK037OUTBOX\"



End Sub

Private Sub cmd_Click(Index As Integer)

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
Dim sqlDBRS As New adodb.Recordset


fileNameTemp = ""
msgTxtTemp = ""

txtStr = txtText1(1).Text

msgTxtTemp = Replace(txtStr, vbCrLf, "','")
newmsgTxtTemp = Replace(txtStr, vbCrLf, ",") & ","

msgTxtTemp2 = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & "," & Right(msgTxtTemp, Len(msgTxtTemp) - InStr(msgTxtTemp, ","))



Dim bid
Dim bid1
bid1 = Replace(msgTxtTemp2, "'", "")
bid = Split(bid1, ",")

Dim lotStr As String

'For i = 0 To UBound(bid) - 1
   ' lotStr = bid(i)
   
    'If lotStr <> "" Then
    
       'If Not HK037tray(lotStr) Then
      '  MsgBox "此卷：" & lotStr & " 此卷不存在，请确认!", vbInformation, "友情提示"
      '   Exit Sub
     '  End If
    '
    
     '  If HK037inboxid(lotStr) Then
      '  MsgBox "此卷：" & lotStr & " 此卷已经装箱，请确认!", vbInformation, "友情提示"
      '   Exit Sub
      ' End If
    
    
   ' End If



'Next i



Dim strid           As String
Dim strcode           As String
Dim pj1
Dim strqb
Dim Rs          As New adodb.Recordset
Dim Rs1          As New adodb.Recordset
Dim Rs2          As New adodb.Recordset
Dim Rs3          As New adodb.Recordset
Dim Rs4          As New adodb.Recordset
Dim Rs5          As New adodb.Recordset

 strqb = Split(newmsgTxtTemp, ",")
 
   strid = " select HK037_LableID.get_inboxid('" & newmsgTxtTemp & "') from dual"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strid, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    seq = Rs.fields(0).Value
    
    strcode = " select HK037_LableID.get_inboxcode('" & newmsgTxtTemp & "') from dual"
   If Rs1.State = adStateOpen Then Rs1.Close
    Rs1.open strcode, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    Qbox = Rs1.fields(0).Value

For j = 0 To UBound(strqb) - 1

If Len(strqb(j)) > 2 Then

strqdetail = "  select gg.container,gg.subcontainer,gg.customerpt,gg.qty,gg.lot,gg.firstname,gg.pkg  from pj_hk037_traydetails gg  where gg.qboxnum in ('" & strqb(j) & "') "

 If Rs2.State = adStateOpen Then Rs2.Close
    Rs2.open strqdetail, Cnn, adOpenStatic, adLockReadOnly, adCmdText

strinbox = "  insert into pj_hk037_traydetails values " & _
" ('" & Rs2.fields(0).Value & "','" & Rs2.fields(1).Value & "','INBOX','" & seq & "','" & Qbox & "','" & Rs2.fields(2).Value & "','" & Rs2.fields(3).Value & "', " & _
"  '" & Rs2.fields(4).Value & "','" & Rs2.fields(5).Value & "','" & Rs2.fields(6).Value & "',sysdate) "
 AddSql (strinbox)

End If
Next j


stribtext = " select tr.customerpt || ',' || tr.lot || ',' || to_char(sysdate, 'YYWW') || ',' || " & _
      " sum(tr.qty) || ',' || tr.pkg || ',' || 'SUB011' || ',' || tr.seq || ',' || " & _
      " tr.qboxnum || ',' || to_char(sysdate, 'YYYY-MM-DD') || ',' || '3000' " & _
"  from pj_hk037_traydetails tr  where tr.qboxnum = '" & Qbox & "' " & _
" group by tr.customerpt, tr.lot,to_char(sysdate, 'YYWW'),tr.pkg, " & _
"   'SUB011', tr.seq, tr.qboxnum, to_char(sysdate, 'YYYY-MM-DD'), '3000' "


 If Rs3.State = adStateOpen Then Rs3.Close
    Rs3.open stribtext, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
   fileNameTemp = Qbox & Format(Now(), "YYYYMMDDHHmmSS")
   dirtemp = TxtDirInQbox.Text
   Call addLabelTxt(fileNameTemp, Rs3.fields(0).Value, dirtemp)

 
txtText1(1).Text = ""
txtText1(1).SetFocus
End Sub


Private Sub cmdout1_Click(Index As Integer)
Dim fileNameTemp As String
Dim dirtemp As String
Dim msgTxtTemp As String
Dim msgTxtTemp2 As String
Dim newmsgTxtTemp As String



fileNameTemp = ""
msgTxtTemp = ""

txtStr = txtText2(2).Text

msgTxtTemp = Replace(txtStr, vbCrLf, "','")
newmsgTxtTemp = Replace(txtStr, vbCrLf, ",") & ","
msgTxtTemp2 = Mid(msgTxtTemp, 1, InStr(msgTxtTemp, ",") - 1) & "," & Right(msgTxtTemp, Len(msgTxtTemp) - InStr(msgTxtTemp, ","))

Dim bid
Dim bid1
bid1 = Replace(msgTxtTemp2, "'", "")
bid = Split(bid1, ",")

Dim lotStr As String

'For i = 0 To UBound(bid) - 1
   ' lotStr = bid(i)
   
    'If lotStr <> "" Then
    
       'If Not HK037tray(lotStr) Then
      '  MsgBox "此卷：" & lotStr & " 此卷不存在，请确认!", vbInformation, "友情提示"
      '   Exit Sub
     '  End If
    '
    
     '  If HK037inboxid(lotStr) Then
      '  MsgBox "此卷：" & lotStr & " 此卷已经装箱，请确认!", vbInformation, "友情提示"
      '   Exit Sub
      ' End If
    
    
   ' End If



'Next i

  
 Dim strid           As String
Dim strcode           As String
Dim pj1
Dim strqb
Dim Rs          As New adodb.Recordset
Dim Rs1          As New adodb.Recordset
Dim Rs2          As New adodb.Recordset
Dim Rs3          As New adodb.Recordset
Dim Rs4         As New adodb.Recordset
Dim Rs5          As New adodb.Recordset
Dim Rs6          As New adodb.Recordset
  
 strqb = Split(newmsgTxtTemp, ",")

strid = " select HK037_LableID.get_outboxid('" & newmsgTxtTemp & "') from dual"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.open strid, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    seq = Rs.fields(0).Value
    
    strcode = " select HK037_LableID.get_outboxcode('" & newmsgTxtTemp & "') from dual"
   If Rs1.State = adStateOpen Then Rs1.Close
    Rs1.open strcode, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    Qbox = Rs1.fields(0).Value

For j = 0 To UBound(strqb) - 1

If Len(strqb(j)) > 2 Then

strqdetail = "  select gg.container,gg.subcontainer,gg.customerpt,gg.qty,gg.lot,gg.firstname,gg.pkg  from pj_hk037_traydetails gg  where gg.qboxnum in ('" & strqb(j) & "') "

 If Rs2.State = adStateOpen Then Rs2.Close
    Rs2.open strqdetail, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
   If Not Rs2.EOF Then
    Do While Not Rs2.EOF

strinbox = "  insert into pj_hk037_traydetails values " & _
" ('" & Rs2.fields(0).Value & "','" & Rs2.fields(1).Value & "','OUTBOX','" & seq & "','" & Qbox & "','" & Rs2.fields(2).Value & "','" & Rs2.fields(3).Value & "', " & _
"  '" & Rs2.fields(4).Value & "','" & Rs2.fields(5).Value & "','" & Rs2.fields(6).Value & "',sysdate) "
 AddSql (strinbox)
 Rs2.MoveNext
 Loop

End If
End If
Next j


stribtext = " select tr.customerpt || ',' || tr.lot || ',' || to_char(sysdate, 'YYWW') || ',' || " & _
      " sum(tr.qty) || ',' || tr.pkg || ',' || 'SUB011' || ',' || tr.seq || ',' || " & _
      " tr.qboxnum || ',' || to_char(sysdate, 'YYYY-MM-DD') || ',' || '3000' " & _
"  from pj_hk037_traydetails tr  where tr.qboxnum = '" & Qbox & "' " & _
" group by tr.customerpt, tr.lot,to_char(sysdate, 'YYWW'),tr.pkg, " & _
"   'SUB011', tr.seq, tr.qboxnum, to_char(sysdate, 'YYYY-MM-DD'), '3000' "


 If Rs3.State = adStateOpen Then Rs3.Close
    Rs3.open stribtext, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
   fileNameTemp = Qbox & Format(Now(), "YYYYMMDDHHmmSS")
   dirtemp = txtText3.Text
   Call addLabelTxt(fileNameTemp, Rs3.fields(0).Value, dirtemp)



strsubcon = " select distinct a.subcontainer from pj_hk037_traydetails  a where a.qboxnum in ('" & msgTxtTemp2 & "') "

    If Rs4.State = adStateOpen Then Rs4.Close
    Rs4.open strsubcon, Cnn, adOpenStatic, adLockReadOnly, adCmdText
  If Not Rs4.EOF Then
    Do While Not Rs4.EOF

strqty = " select gg.qty from pj_hk037_traydetails gg where gg.subcontainer  =  '" & Rs4.fields(0).Value & "'"
 
   If Rs5.State = adStateOpen Then Rs5.Close
    Rs5.open strqty, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
strcon = " select distinct a.container from pj_hk037_traydetails  a where a.subcontainer  =  '" & Rs4.fields(0).Value & "' "
  
     If Rs6.State = adStateOpen Then Rs6.Close
    Rs6.open strcon, Cnn, adOpenStatic, adLockReadOnly, adCmdText

 
     
Strqbdt = "insert into TSV_QboxNumber_Details(Pdata1,Productname, WAFERNUMBER, NDPW, LAST_UPDATE_DATE, CONTAINERID, QBOXNUMBER," & _
"  CONTAINERNAME, WAFERSCRIBENUMBER, WORKORDERNAME, FIRSTNAME, CUSTOMERNAME, SpecName) " & _
" select distinct d.alternatename || '/' || f.workorderattr3 Pdata1, e.productname, a.wafernumber, '" & Rs5.fields(0).Value & "', sysdate Pdate, b.containerid,'" & Qbox & "'," & _
" '" & Rs4.fields(0).Value & "', a.waferscribenumber, a.workordername,b.firstname, 'HK037',5275  from a_lotwafers a,  container b,  a_lotattributes c, product d," & _
" productbase  e, mfgorder f where a.containerid = b.containerid and b.containerid = c.containerid and d.productbaseid = e.productbaseid " & _
" and f.mfgordername = a.workordername and b.productid = d.productid and b.containername = '" & Rs6.fields(0).Value & "'   AND C.WAFERBIN IS NOT NULL"
 AddSql (Strqbdt)
 
  Rs4.MoveNext
 Loop
 
End If

       
  
txtText2(2).Text = ""
txtText2(2).SetFocus
End Sub

                    



