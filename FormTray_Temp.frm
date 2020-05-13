VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_Tray_tmp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12330
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
   ScaleHeight     =   7800
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTTab0 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "批号标签打印"
      TabPicture(0)   =   "FormTray_Temp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabel3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtText1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdPrint"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtText2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCommand1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtText3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCommand2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCommand3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdpr"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "标签核对"
      TabPicture(1)   =   "FormTray_Temp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtTmpTray"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtThisTray"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtStatus"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdpr 
         Caption         =   "重置"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   6720
         TabIndex        =   16
         Top             =   4800
         Width           =   990
      End
      Begin VB.CommandButton cmdCommand3 
         Caption         =   "打印"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         TabIndex        =   15
         Top             =   3600
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdCommand2 
         Caption         =   "确认"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         TabIndex        =   14
         Top             =   2640
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox txtText3 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   2640
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CommandButton cmdCommand1 
         Caption         =   "补打"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   4440
         TabIndex        =   11
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox txtStatus 
         Height          =   4335
         Left            =   -73680
         TabIndex        =   10
         Top             =   2760
         Width           =   8655
      End
      Begin VB.TextBox txtThisTray 
         Height          =   285
         Left            =   -72600
         TabIndex        =   9
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtTmpTray 
         Height          =   285
         Left            =   -72600
         TabIndex        =   8
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtText2 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   3600
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H8000000D&
         Caption         =   "打印"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   960
         MaskColor       =   &H0000FFFF&
         TabIndex        =   3
         Top             =   4800
         Width           =   2895
      End
      Begin VB.TextBox txtText1 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   1740
         Width           =   3735
      End
      Begin VB.Label lblLabel3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入密码:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   13
         Top             =   2760
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正式标签:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73680
         TabIndex        =   7
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "临时标签:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73680
         TabIndex        =   6
         Top             =   1582
         Width           =   975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷盘编号："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Top             =   3600
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工单号："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   1800
         Width           =   900
      End
   End
End
Attribute VB_Name = "Frm_Tray_tmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCommand1_Click()


MsgBox "请输入补打密码", vbInformation, "提示"

lblLabel3.Visible = True
txtText3.Visible = True
txtText3.Text = ""
txtText1.Visible = False
lbl(0).Visible = False
cmdPrint.Visible = False
cmdCommand2.Visible = True


End Sub

Private Sub cmdCommand2_Click()

If Trim(txtText3.Text) = "14526" Then

lbl(1).Visible = True
txtText2.Visible = True
txtText2.Text = ""
cmdCommand3.Visible = True
lblLabel3.Visible = False
txtText3.Visible = False
cmdCommand2.Visible = False
cmdCommand1.Visible = False

MsgBox "请输入单个临时卷盘ID进行补打", vbInformation, "提示"

Else
MsgBox "请输入补打密码", vbInformation, "提示"

Exit Sub

End If

End Sub

Private Sub cmdCommand3_Click()

 Dim sqlTray   As String
 
 Dim sqlTrayRS As New ADODB.Recordset
 
 Dim reprint As String
 
 Dim op As String
 
 op = gUserName
 
 sqlTray = " select w.tray_id,nvl(w.remark2,'ERROR') from TRAY_TEMP W where w.tray_id = '" & Trim(txtText2.Text) & "'"
 
  If sqlTrayRS.State = adStateOpen Then sqlTrayRS.Close
    sqlTrayRS.Open sqlTray, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
  If Not sqlTrayRS.EOF Then
  
  Call addLabelTxt(sqlTrayRS.Fields(0).Value, sqlTrayRS.Fields(1).Value, "\\10.160.1.14\BarCode\37\37TRAY_TEMP\")
   MsgBox "打印完成", vbInformation, "!"
   
   reprint = "update TRAY_TEMP set last_reprint_by = '" & gUserName & "' ,last_reprint_date = sysdate ,print = print + 1 where tray_id = '" & sqlTrayRS.Fields(0).Value & "'"
   AddSql (reprint)
  
  Else
      
     MsgBox "卷盘ID不存在", vbInformation, "提示"
     Exit Sub
      
  End If
 



End Sub

Private Sub cmdpr_Click()

txtText1.Visible = True
lbl(0).Visible = True
lbl(1).Visible = False
txtText2.Visible = False
lblLabel3.Visible = False
txtText3.Visible = False
cmdCommand2.Visible = False
cmdCommand3.Visible = False
cmdCommand1.Visible = True
cmdPrint.Visible = True

End Sub

Private Sub cmdPrint_Click()

    Dim sqlDB   As String

    Dim sqlDBRS As New ADODB.Recordset
    
     Dim sqlshop   As String

    Dim sqlshopRS As New ADODB.Recordset

    Dim QTY     As Long

    Dim label   As String

    Dim J       As Integer

    Dim i       As Integer
    
    Dim num As String
    
    Dim tray_id As String
    
    Dim tray_id_log As String
    
    Dim User As String
    
    
    Dim sqlDC  As String

    Dim sqlDCRS As New ADODB.Recordset
    
    Dim labelall As String
   
    
    User = gUserName
     
     labelall = ""
     
     sqlshop = "select w.* from TRAY_TEMP W where w.remark1 = '" & Trim(txtText1.Text) & "'  "
     
      If sqlshopRS.State = adStateOpen Then sqlshopRS.Close
    sqlshopRS.Open sqlshop, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not sqlshopRS.EOF Then  '表示有数据了
       
       MsgBox "工单已打印过临时标签", vbInformation, "提示"
       Exit Sub
     
    Else

   sqlDB = "select a.ht_device || ',' || b.propertyvalue || ',' ||e.test_mtrl_desc || ',' ||replace(replace(replace(a.bonded,d.lotid||',',''),'.1','') ,';','') ,sum(c.gross_die_qty) ,replace(replace(replace(a.bonded,d.lotid||',',''),'.1','') ,';','') " & " from shop_order a,shop_order_property b,shop_order_detail c,mappingdatatest d,customeroitbl_test e " & " where a.shop_order = '" & txtText1.Text & "' and b.shop_order = a.shop_order and b.propertyname = 'CUST_PART_NUM1' " & " and c.shop_order = b.shop_order and d.substrateid = c.wafer_id and to_char(e.id) = d.filename " & " group by  a.ht_device,b.propertyvalue,e.test_mtrl_desc,replace(replace(replace(a.bonded,d.lotid||',',''),'.1','') ,';','') "

    If sqlDBRS.State = adStateOpen Then sqlDBRS.Close
    sqlDBRS.Open sqlDB, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
    
    If Not sqlDBRS.EOF Then  '表示有数据了

    QTY = sqlDBRS.Fields(1).Value / 15000
    
    For J = 0 To QTY + 1
    
        tray_id = ""
  
        i = J + 1
        If Len(Trim(Str(i))) = 1 Then
        num = "0" + Trim(Str(i))
        Else
        num = Trim(Str(i))
        End If
        
        If J = 0 Then
   
            label = sqlDBRS.Fields(2).Value + "M-R" + num + "," + sqlDBRS.Fields(0).Value
            tray_id = sqlDBRS.Fields(2).Value + "M-R" + num
  
        Else
  
            label = sqlDBRS.Fields(2).Value + "-R" + num + "," + sqlDBRS.Fields(0).Value
            tray_id = sqlDBRS.Fields(2).Value + "-R" + num
  
        End If
        
       
        
        
        tray_id_log = "insert into TRAY_TEMP(tray_id,print,flag,Create_by,create_date,remark1,remark2 ) values ('" & tray_id & "',1,0,'" & User & "',SYSDATE,'" & Trim(txtText1.Text) & "','" & label & "' )"
        
        AddSql (tray_id_log)
        
      labelall = labelall + label + vbCrLf
      
    Next J
    
    
     Call addLabelTxt(sqlDBRS.Fields(2).Value + "-R" + num, labelall, "\\10.160.1.14\BarCode\37\37TRAY_TEMP\")
    
    
'    sqlDC = " select distinct replace(replace(replace(a.bonded, b.cust_lot_id || ',', ''),  '.1', ''), ';', '') || ',' || 'D'||  decode(e.status,'0','0',  DATE_CODE_CONVERT.DC_CONVERT(to_char(a.erp_create_date, 'YYYY-MM-DD'),1)) || 'B' || e.BLINE || 'C' || nvl(e.CODE,0)  from shop_order a  left join shop_order_detail b " & _
'            " on b.shop_order = a.shop_order left join mappingdatatest c on c.substrateid = replace( b.wafer_id,'+','') left join customeroitbl_test d " & _
'            "  on to_char(d.id) = c.filename and d.source_batch_id = c.lotid  left join code37 e  ON e.DEVICE = d.mpn_desc  where a.shop_order = '" & Trim(txtText1.Text) & "' "


     sqlDC = "select distinct replace(replace(replace(a.bonded, b.cust_lot_id || ',', ''), '.1', ''), ';', '') || ',' || 'D' || decode(e.status, " & _
            " '0', '0',f.dc) || 'B' || e.BLINE || 'C' || nvl(e.CODE, 0)  from shop_order a  left join shop_order_detail b   on b.shop_order = a.shop_order " & _
            " left join mappingdatatest c  on c.substrateid = replace(b.wafer_id, '+', '')  left join customeroitbl_test d  on to_char(d.id) = c.filename " & _
            " and d.source_batch_id = c.lotid left join mappingdatatest cc  on cc.substrateid = b.wafer_id   left join customeroitbl_test dd " & _
            " on to_char(dd.id) = cc.filename  and dd.source_batch_id = cc.lotid  left join code37 e  ON e.DEVICE = d.mpn_desc left join  TBL37TESTDC f " & _
            " on f.jobid = dd.test_mtrl_desc  where a.shop_order = '" & Trim(txtText1.Text) & "' "
     
     
    
    If sqlDCRS.State = adStateOpen Then sqlDCRS.Close
    sqlDCRS.Open sqlDC, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not sqlDCRS.EOF Then
    
     Call addLabelTxt(sqlDBRS.Fields(2).Value + "DC", sqlDCRS.Fields(0).Value, "\\10.160.1.14\BarCode\37\37TRAY_TEMP_DC\")
    
    Else
    
      MsgBox "缺少打标信息", vbInformation, "!"
    
    End If
    
    

    MsgBox "打印完成", vbInformation, "!"
    
    Else
    
        MsgBox "没有数据", vbInformation, "提示"
        Exit Sub

    End If
    
    End If
  
End Sub

Private Sub Form_Activate()
    txtText1.SetFocus
End Sub

Private Sub SSTTab0_Click(PreviousTab As Integer)

    Select Case PreviousTab

        Case 0
            txtTmpTray.SetFocus
        
        Case 1
    
            txtText1.SetFocus
    End Select

End Sub

Private Sub txtThisTray_KeyPress(KeyAscii As Integer)
Dim strCode As String
Dim strTmpCode As String
Dim sMatch As String
Dim strSql As String

If KeyAscii <> vbKeyReturn Then
    Exit Sub
End If

txtStatus.BackColor = vbWhite
If txtTmpTray.Text = "" Then
    MsgBox "临时卷盘没有数据, 请先扫描临时卷盘", vbCritical, "警告"
    txtTmpTray.SetFocus
    Exit Sub
Else
    strTmpCode = Trim$(txtTmpTray.Text)
End If

If txtThisTray.Text = "" Then
    MsgBox "实际卷盘没有数据,请重新扫描", vbCritical, "警告"
    Exit Sub
Else
    strCode = Trim$(txtThisTray.Text)
End If

If (Left$(strCode, 1) <> "S") Or (strCode = strTmpCode) Then
    txtThisTray.Text = ""
    Exit Sub
End If


If Get_OracleCnt("select * from CHECK_TRAY_HISTORY where real_Taryid= '" & strCode & "'") > 0 Then
    txtStatus.BackColor = vbRed
    MsgBox "该卷盘已经核对过, 请确认是否有问题", vbExclamation, "提示"
    Exit Sub
End If


'sMatch = Mid(strTmpCode, 1, InStr(strTmpCode, "-") - 3)

sMatch = Left(Replace$(strTmpCode, ".1", ""), Len(Replace$(strTmpCode, ".1", "")) - 4)

If InStr(strCode, sMatch) > 0 Then
    ' 正确
    txtStatus.BackColor = vbBlue
    strSql = "insert into CHECK_TRAY_HISTORY values('" & strTmpCode & "', '" & strCode & "', 'Y',sysdate, '" & gUserName & "')"
    AddSql (strSql)
    
    Clear
Else
    ' 错误
    txtStatus.BackColor = vbRed
    MsgBox "出错", vbCritical, "错误提示:"
    
    strSql = "insert into CHECK_TRAY_HISTORY values('" & strTmpCode & "', '" & strCode & "', 'N',sysdate, '" & gUserName & "')"
    AddSql (strSql)
    
End If

End Sub

Private Sub Clear()
txtTmpTray.Text = ""
txtThisTray.Text = ""
txtTmpTray.SetFocus

End Sub

Private Sub txtTmpTray_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then
    Exit Sub
End If

If InStr(Trim$(txtTmpTray.Text), "-") = 0 Then
    txtTmpTray.Text = ""
Else
    txtThisTray.SetFocus
End If

End Sub














