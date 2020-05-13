VERSION 5.00
Begin VB.Form Form_NEWBOX 
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14340
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
   ScaleHeight     =   8160
   ScaleWidth      =   14340
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      Begin VB.CommandButton cmd 
         Caption         =   "生成数据"
         Height          =   360
         Left            =   9360
         TabIndex        =   9
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox txtText4 
         Height          =   495
         Left            =   8400
         TabIndex        =   8
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox txtText3 
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   3960
         Width           =   3975
      End
      Begin VB.TextBox txtText2 
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox txtText1 
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "箱号"
         Height          =   315
         Left            =   7200
         TabIndex        =   7
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label lblF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-F批号"
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lblE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-E批号"
         Height          =   315
         Left            =   480
         TabIndex        =   3
         Top             =   2520
         Width           =   1485
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-A批号"
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   1080
         Width           =   1485
      End
   End
End
Attribute VB_Name = "Form_NEWBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Click()
On Error GoTo DealError
Dim qboxtemp As String
Dim conn As String
Dim delstr As String
Dim qboxcount As String
Dim qbcount As String
Dim custcount As String
Dim concust As String
Dim concust1 As String
Dim nEWBOX As String
Dim cust As String
Dim yzqbox As String
Dim yzqboxnum As String
Dim yzcount As String
Dim yzqboxold As String
Dim yzqboxnew As String
Dim yzspec As String
Dim yzqbinser As String
Dim spec As String
Dim custspec As String
Dim upspec As String
Dim spcon
Dim Rs   As New adodb.Recordset
Dim Rs1   As New adodb.Recordset
Dim Rs2   As New adodb.Recordset
Dim Rs3   As New adodb.Recordset
Dim Rs4   As New adodb.Recordset
Dim Rsyz1   As New adodb.Recordset
Dim Rsyz2   As New adodb.Recordset
Dim Rsyz3   As New adodb.Recordset
Dim Rsyz4   As New adodb.Recordset

If txtText4.Text = "" Then
   
    MsgBox "箱号不能为空", vbInformation, "友情提示"
    Exit Sub
End If

qboxtemp = UCase(Trim(txtText4.Text))



If Len(Trim(txtText1.Text)) < 2 And Len(Trim(txtText2.Text)) < 2 And Len(Trim(txtText3.Text)) < 2 Then
      MsgBox "批号不能为空", vbInformation, "友情提示"
 Exit Sub
 End If
 
 If Len(Trim(txtText1.Text)) > 0 And InStr(txtText1.Text, "-A") = 0 Then
  MsgBox "请输入-A批号", vbInformation, "友情提示"
 Exit Sub
 
 End If
 
  If Len(Trim(txtText2.Text)) > 0 And InStr(txtText2.Text, "-E") = 0 Then
  MsgBox "请输入-E批号", vbInformation, "友情提示"
 Exit Sub
 
 End If
 
  If Len(Trim(txtText3.Text)) > 0 And InStr(txtText3.Text, "-F") = 0 Then
  MsgBox "请输入-F批号", vbInformation, "友情提示"
 Exit Sub
 
 End If
 
Call DelERPQboxData(qboxtemp)

 

conn = txtText1.Text & "'" & "," & "'" & txtText2.Text & "'" & "," & "'" & txtText3.Text
spcon = Split(Replace(conn, "'", ""), ",")
concust = "select count(distinct c.customershortname) from container a, a_lotwafers b, mappingdatatest c where b.containerid = a.containerid" & _
           " and c.substrateid = b.waferscribenumber  and a.containername in( '" & conn & "') "
 If Rs.State = adStateOpen Then Rs.Close
    Rs.open concust, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    custcount = Rs.fields(0).Value
    
 If custcount > 1 Then
       MsgBox "批号属于不同客户", vbInformation, "友情提示"
    Exit Sub
 
 Else
 
 concust1 = "select distinct c.customershortname from container a, a_lotwafers b, mappingdatatest c where b.containerid = a.containerid" & _
           " and c.substrateid = b.waferscribenumber  and a.containername in( '" & conn & "') "
 If Rs1.State = adStateOpen Then Rs1.Close
    Rs1.open concust1, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    cust = Rs1.fields(0).Value
 
 End If
 
 If cust = "YZ22" And (Mid(qboxtemp, 1, 1) = "T" Or Mid(qboxtemp, 1, 1) = "K") Then
 
 yzspec = "  select  count(a.specname)  from a_wiplothistory a  Where a.containername = '" & Trim(txtText1.Text) & "' " & _
        "  and a.creationtimestamp = (select max(b.creationtimestamp) from a_wiplothistory b  where b.containername = '" & Trim(txtText1.Text) & "') "
        
    If Rsyz4.State = adStateOpen Then Rsyz4.Close
    Rsyz4.open yzspec, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    spec = Rsyz4.fields(0).Value
        
 If spec = 0 Then
 
       MsgBox "请先过站", vbInformation, "友情提示"
    Exit Sub
 
 End If
 
 yzqbox = " select count(*) from tsv_qboxnumber_details a  where a.containername in ('" & conn & "')"
 
  If Rsyz1.State = adStateOpen Then Rsyz1.Close
    Rsyz1.open yzqbox, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    yzcount = Rsyz1.fields(0).Value
    
   
  If yzcount = 0 Then
  
     yzqbinser = "  SELECT B.CONTAINERNAME,trglabelseq.QTSeqNormal('" & cust & "','" & Trim(txtText1.Text) & "') FROM" & _
         " container b,a_lotattributes c  WHERE B.CONTAINERID=C.CONTAINERID " & _
         " AND B.CONTAINERNAME='" & Trim(txtText1.Text) & "'  AND c.customername='" & cust & "' "
         
    AddSql (yzqbinser)
 
End If
         
       ' MsgBox "请先过站", vbInformation, "友情提示"
  '  Exit Sub
 
 
 yzqboxnum = "select distinct a.qboxnumber from tsv_qboxnumber_details a  where a.containername in ('" & conn & "')"
 
   If Rsyz2.State = adStateOpen Then Rsyz2.Close
    Rsyz2.open yzqboxnum, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    yzqboxold = Rsyz2.fields(0).Value
 
 yzqboxnew = " update tsv_qboxnumber_details set qboxnumber = '" & qboxtemp & "' where containername in ('" & conn & "')"
 
  AddSql (yzqboxnew)
 

 Else
 
    
  
  delstr = " delete from tsv_qboxnumber_details where containername in ( '" & conn & "') "
  
   AddSql (delstr)
   
  qboxcount = "select count(distinct b.qboxnumber) from l_sequence b where b.containername in ( '" & conn & "')"
  
   If Rs2.State = adStateOpen Then Rs2.Close
    Rs2.open qboxcount, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    qbcount = Rs2.fields(0).Value
    
  If qbcount > 1 Then
    MsgBox "存在不同箱号，请确认", vbInformation, "友情提示"
    
    Exit Sub
    
   Else
   
   For i = 0 To UBound(spcon)
   If Len(spcon(i)) > 3 Then
   
 qbinser = "  SELECT trglabelseq.QTSeqNormal('" & cust & "','" & spcon(i) & "') FROM" & _
         " container b,a_lotattributes c  WHERE B.CONTAINERID=C.CONTAINERID " & _
         " AND B.CONTAINERNAME='" & spcon(i) & "'  AND c.customername='" & cust & "' "
         
   If Rs3.State = adStateOpen Then Rs3.Close
    Rs3.open qbinser, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    nEWBOX = Rs3.fields(0).Value
    
  
  End If
  Next i

 

  
  End If
  
   
End If

  
  custspec = " select top 1 b.SpecName from TblQBoxCustomer b where b.Customer = '" & cust & "'  "
  
  If Rs4.State = adStateOpen Then Rs4.Close
    Rs4.open custspec, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText
    spec = Rs4.fields(0).Value
        
  upspec = " update  tsv_qboxnumber_details  set  specname = '" & spec & "' where qboxnumber = '" & nEWBOX & "'  "
  AddSql (upspec)

txtText1.Text = ""
txtText2.Text = ""
txtText3.Text = ""
txtText4.Text = ""

Exit Sub

DealError:
MsgBox "箱号已建立入库单 "
    

End Sub



