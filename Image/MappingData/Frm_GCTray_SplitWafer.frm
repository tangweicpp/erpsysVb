VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#3.5#0"; "fpSpr35.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Frm_GCTray_SplitWafer 
   Caption         =   "GC ��WLA��Normal�趨"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   16140
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
      Height          =   480
      Left            =   6120
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "�˳�"
      Height          =   480
      Left            =   8640
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "�޸�"
      Height          =   480
      Left            =   3600
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "����"
      Height          =   480
      Left            =   1080
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox TxtPeace 
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox TxtLotID 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
      Width           =   2535
   End
   Begin MSDataListLib.DataCombo DCbMainItem 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DCbChildItem 
      Height          =   315
      Left            =   7200
      TabIndex        =   3
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin FPSpreadADO.fpSpread fps 
      Height          =   5055
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   3120
      Width           =   12855
      _Version        =   196613
      _ExtentX        =   22675
      _ExtentY        =   8916
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
      SpreadDesigner  =   "Frm_GCTray_SplitWafer.frx":0000
      TextTip         =   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϊ�趨GC GC0310��GC0312��GC6123���������棬Normal��Ƭ��"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1440
      TabIndex        =   8
      Top             =   240
      Width           =   5520
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ƭ����"
      Height          =   195
      Left            =   6360
      TabIndex        =   6
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LotID��"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ʒ�Ϻţ�"
      Height          =   195
      Left            =   6120
      TabIndex        =   2
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ����֣�"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "Frm_GCTray_SplitWafer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oiRS    As New ADODB.Recordset

Private Enum E_FPS0          'Detail�֭��
    E_SeqId = 1                '���
    E_CustPT                  '�ͻ�����
    E_QtechPT                 '��Ʒ�Ϻ�
    E_LOTID                   'LotID
    E_Qty                    '����
    E_Date                    '����

    
    E_End
    
End Enum



Private Sub Command1_Click()
''����
'Dim lotIDTemp As String
'Dim dtTemp As Date
'Dim sqlTemp As String
'Dim remarkTemp As String
'
'
'
'If Trim(Text1.Text) <> "" Then
'
'    lotIDTemp = Trim(Text1.Text)
'    dtTemp = DTPicker1.Value
'    remarkTemp = ""
'
'    '�ж������Lot�ţ��Ƿ���ȷ
'
'    If JudgeLot2(lotIDTemp) Then
'
'
'        '�ж��Ƿ���� ��������ʾ��Ϣ
'        If Not (JudgeLot(lotIDTemp)) Then
'        sqlTemp = "insert into WipreportDate(lotid,lotdate,remark) values ( '" & lotIDTemp & "',to_date('" & dtTemp & "','yyyy-mm-dd'),'" & remarkTemp & "' ) "
'        AddSql (sqlTemp)
'        MsgBox "��ӳɹ�!"
'
'        Else
'
'        MsgBox "LotId:" & lotIDTemp & "�Ѵ��ڣ�"
'        End If
'
'    Else
'         MsgBox "LotId:" & lotIDTemp & "��Mesϵͳ�в����ڣ���ȷ��Lot�ţ�"
'
'    End If
'
'
'Else
'MsgBox "��������LotId!"
'End If



 ExporToExcel ("  select id,CustomerPT,productname,lotid,qty,to_char(createddate,'YYYY-MM-DD') InDate  from  TSV_GCTRAY_SetWLA where   flag='Y' order by id desc  ")
 

End Sub

Private Sub Command2_Click()
'�޸�
Dim lotIDTemp As String
Dim dtTemp As Date
Dim sqlTemp As String
Dim remarkTemp As String


If Trim(Text2.Text) <> "" Then

    lotIDTemp = Trim(Text2.Text)
    dtTemp = DTPicker3.Value
    remarkTemp = ""
    
    '�ж��Ƿ���� �������޸ģ���������ʾ
     If JudgeLot(lotIDTemp) Then
     
        sqlTemp = "update WipreportDate set lotdate=to_date('" & dtTemp & "','yyyy-mm-dd'), remark='" & remarkTemp & "'    where lotid='" & lotIDTemp & "' "
        AddSql (sqlTemp)
        MsgBox "�޸ĳɹ�!"
        
    Else
        
          MsgBox "LotId:" & lotIDTemp & "�����ڣ�"
     End If
    

    

Else
MsgBox "��������LotId!"
End If


End Sub

Public Function JudgeLot(lotIDTemp As String) As Boolean

Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "  select * from WipreportDate where lotid='" + lotIDTemp + "' "
         
slectResult = QueryStr(cmdStr)
JudgeLot = slectResult
End Function


Public Function JudgeWafer(lotIDTemp As String) As Boolean

Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "  select * from  TSV_GCTRAY_SetWLA where lotid='" + lotIDTemp + "' and flag='Y' "
         
slectResult = QueryStr(cmdStr)
JudgeWafer = slectResult
End Function


Public Function JudgeLot2(lotIDTemp As String) As Boolean

Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "  select * from A_Lotwafers where  wafernumber='" + lotIDTemp + "' "
         
         
slectResult = QueryStr(cmdStr)
JudgeLot2 = slectResult
End Function

Public Function JudgeWafer2(lotIDTemp As String) As Boolean

Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
'cmdStr = "  select * from  TSV_GCTRAY_SetWLA where lotid='" + lotIDtemp + "' and flag='Y' "

cmdStr = "  select * from mappingdatatest where lotid='" + lotIDTemp + "' "

    
slectResult = QueryStr(cmdStr)
JudgeWafer2 = slectResult
End Function


Public Function JudgeWaferProduct(lotIDTemp As String, productTemp As String) As Boolean

Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
'cmdStr = "  select * from  TSV_GCTRAY_SetWLA where lotid='" + lotIDtemp + "' and flag='Y' "

cmdStr = " select a.* from ib_waferlist a, ib_wohistory b  where a.waferlot='" + lotIDTemp + "' and b.ordername=a.ordername and b.product='" + productTemp + "' "

    
slectResult = QueryStr(cmdStr)
JudgeWaferProduct = slectResult
End Function



Private Sub Command3_Click()
 ExporToExcel ("select lotid,lotdate,remark,CreateDate from WipreportDate order by CreateDate desc ")
End Sub

Private Sub Command4_Click()


'���� Remark  2012-06-18
Dim lotIDTemp As String
Dim sqlTemp As String
Dim remarkTemp As String



If Trim(TxtWafer.Text) <> "" Then

    lotIDTemp = Trim(TxtWafer.Text)
    remarkTemp = Trim(TxtRemark.Text)
    
    '�ж������Lot�ţ��Ƿ���ȷ
    
    If JudgeWafer2(lotIDTemp) Then
    
    
        '�ж��Ƿ���� ��������ʾ��Ϣ
        If Not (JudgeWafer(lotIDTemp)) Then
        sqlTemp = "insert into WipreportDateRemark(lotid,remark) values ( '" & lotIDTemp & "','" & remarkTemp & "' ) "
        AddSql (sqlTemp)
        MsgBox "��ӳɹ�!"
        
        Else
        
        MsgBox "WaferId:" & lotIDTemp & "�Ѵ��ڣ�"
        End If
        
    Else
         MsgBox "WaferId:" & lotIDTemp & "��Mesϵͳ�в����ڣ���ȷ��Wafer�ţ�"
    
    End If
    

Else
MsgBox "��������WaferId!"
End If






End Sub

Private Sub Command5_Click()

'�޸� Remark 2012-06-18
Dim lotIDTemp As String
Dim sqlTemp As String
Dim remarkTemp As String


If Trim(TxtWafer2.Text) <> "" Then

    lotIDTemp = Trim(TxtWafer2.Text)

    remarkTemp = Trim(TxtRemark2.Text)
    
    '�ж��Ƿ���� �������޸ģ���������ʾ
     If JudgeWafer(lotIDTemp) Then
     
        sqlTemp = "update WipreportDateRemark set  remark='" & remarkTemp & "'    where lotid='" & lotIDTemp & "' "
        AddSql (sqlTemp)
        MsgBox "�޸ĳɹ�!"
        
    Else
        
          MsgBox "WaferId:" & lotIDTemp & "�����ڣ�"
     End If
    

    

Else
MsgBox "��������WaferId!"
End If





End Sub

Private Sub Command6_Click()
 ExporToExcel ("select lotid as WaferId,remark,CreateDate from WipreportDateRemark order by CreateDate desc ")
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdModify_Click()




'���� Remark  2012-06-18
Dim cusPTTemp As String
Dim htPTTemp As String
Dim lotIDTemp As String
Dim pcsQty As Integer
Dim userNameTemp As String

Dim sqlTemp As String
Dim remarkTemp As String



If TxtLotID.Text <> "" And TxtPeace.Text <> "" Then
'    cusPTtemp = DCbMainItem.Text
'    htPTtemp = DCbChildItem.Text
    lotIDTemp = UCase(Trim(TxtLotID.Text))
    pcsQty = CInt(UCase(TxtPeace.Text))
    
    userNameTemp = UCase(gUserName)
    
    '�ж������Lot�ţ��Ƿ���ȷ
   
            
        '�ж��Ƿ���� ��������ʾ��Ϣ
        If JudgeWafer(lotIDTemp) Then
        sqlTemp = " Update TSV_GCTRAY_SetWLA set qty=" & pcsQty & " ,lastupdateby='" & userNameTemp & "',lastupdatedate=sysdate where lotid='" & lotIDTemp & "'"
 
        AddSql (sqlTemp)

       MsgBox "�޸ĳɹ�!", vbInformation, "������ʾ"
            
        
        ShowData_Where
        
        Else
        

         MsgBox "LotID:" & lotIDTemp & "�����ڣ��޷��޸ģ�", vbInformation, "������ʾ"
         
         
        End If

    
   
Else

 MsgBox "������������Ϣ!", vbInformation, "������ʾ"
    

Exit Sub

End If






End Sub

Private Sub CmdSave_Click()

'���� Remark  2012-06-18
Dim cusPTTemp As String
Dim htPTTemp As String
Dim lotIDTemp As String
Dim pcsQty As Integer
Dim userNameTemp As String

Dim sqlTemp As String
Dim remarkTemp As String



If DCbMainItem.Text <> "" And DCbChildItem.Text <> "" And TxtLotID.Text <> "" And TxtPeace.Text <> "" Then
    cusPTTemp = DCbMainItem.Text
    htPTTemp = DCbChildItem.Text
    lotIDTemp = UCase(Trim(TxtLotID.Text))
    pcsQty = CInt(UCase(TxtPeace.Text))
    
    userNameTemp = UCase(gUserName)
    
    '�ж������Lot�ţ��Ƿ���ȷ
    
    If JudgeWafer2(lotIDTemp) Then
        '�ж������lot���Ϻ��Ƿ���ȷ
        
          If JudgeWaferProduct(lotIDTemp, htPTTemp) Then
            
                '�ж��Ƿ���� ��������ʾ��Ϣ
                If Not (JudgeWafer(lotIDTemp)) Then
                sqlTemp = "insert into TSV_GCTRAY_SetWLA(id,CustomerPT,productname,lotid,qty,flag,createdby,createddate) values (GCTray_WLA_SEQ.Nextval,'" & cusPTTemp & "','" & htPTTemp & "','" & lotIDTemp & "'," & pcsQty & ",'Y','" & userNameTemp & "',sysdate)  "
                
                
                AddSql (sqlTemp)

               MsgBox "��ӳɹ�!", vbInformation, "������ʾ"
                    
                
                ShowData_Where
                
                Else
                
        
                 MsgBox "LotID:" & lotIDTemp & "��¼�룬�벻Ҫ���¼�룡", vbInformation, "������ʾ"
                 
                 
                End If
          Else
          
   
             
           MsgBox "��ȷ�ϴ�LotID��ѡ��ĳ�Ʒ�Ϻ��Ƿ���ȷ ��", vbInformation, "������ʾ"
              
                
        End If
        
    Else

         
          MsgBox "��ȷ�ϴ�LotID�Ƿ���ȷ��", vbInformation, "������ʾ"
          
    
    End If
    

Else

 MsgBox "������������Ϣ!", vbInformation, "������ʾ"
    

Exit Sub

End If




End Sub


Private Sub ShowData_Where()
Set reportRS = GetGCTrayRptWla()

With fps(0)
        .MaxRows = 0
        If reportRS.RecordCount > 0 Then
            Set .DataSource = reportRS
       
        End If
End With

End Sub


Private Sub DCbChildItem_Change()
TxtLotID.SetFocus
End Sub

Private Sub DCbMainItem_Change()

Dim mainitem_id As String
DCbChildItem.Text = ""

mainitem_id = DCbMainItem.BoundText
'��ѯС��
IniChildItem mainitem_id



End Sub


Private Sub IniChildItem(main_id As String)
Set childItemRS = GetChildItem(main_id)
Set DCbChildItem.RowSource = childItemRS
DCbChildItem.ListField = childItemRS("smallname").Name
DCbChildItem.BoundColumn = childItemRS("id").Name

End Sub


Public Function GetChildItem(mainitemIdTemp As String) As ADODB.Recordset
'��ѯС����
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset


cmdStr = "select productname as id, productname as smallname from  TSV_GCTrayRptSet  where flag='Y'  and gcpt='" & mainitemIdTemp & "' order by  productname  "
      
Set RSResult = getStr(cmdStr)
Set GetChildItem = RSResult
End Function




Private Sub Form_Activate()
'Text1.SetFocus
End Sub

Private Sub Form_Load()
'DTPicker1.Value = DateTime.Date
'DTPicker3.Value = DateTime.Date

IniFpsHeader

IniMainItem

ShowData_Where

End Sub

Private Sub IniMainItem()
Set mainItemRS = GetMainItem()
Set DCbMainItem.RowSource = mainItemRS
DCbMainItem.ListField = mainItemRS("bigname").Name
DCbMainItem.BoundColumn = mainItemRS("id").Name

End Sub

Public Function GetMainItem() As ADODB.Recordset
'��ѯ������
Dim cmdStr As String
Dim RSResult As New ADODB.Recordset

cmdStr = "select distinct gcpt as id  , gcpt as  BIGNAME  from  TSV_GCTrayRptSet  where flag='Y' and gcpt in ('GC0310','GC0312','GC6123') order by gcpt  "


Set RSResult = getStr(cmdStr)
Set GetMainItem = RSResult


End Function


Private Sub IniFpsHeader()
    With fps(0)
        .ReDraw = False
        .MaxCols = E_FPS0.E_End - 1
        .MaxRows = 0
        
        '�]�m�榡
        .DAutoHeadings = False
        .DAutoCellTypes = False
        .DAutoSizeCols = DAutoSizeColsNone
        
        .Col = -1
        .Row = -1
        .Lock = True
        .OperationMode = OperationModeNormal
        .TypeVAlign = TypeVAlignCenter
        .SelForeColor = &HFF8080
        

    
        
        .SetText E_FPS0.E_SeqId, 0, "��¼��"
        .SetText E_FPS0.E_CustPT, 0, "�ͻ�����"
        .SetText E_FPS0.E_QtechPT, 0, "��Ʒ�Ϻ�"
        .SetText E_FPS0.E_LOTID, 0, "LotID"
        .SetText E_FPS0.E_Qty, 0, "Ƭ��"
        .SetText E_FPS0.E_Date, 0, "����"
        
        
        
        .ColWidth(E_FPS0.E_SeqId) = 8
        .ColWidth(E_FPS0.E_CustPT) = 12
        .ColWidth(E_FPS0.E_QtechPT) = 12
        .ColWidth(E_FPS0.E_LOTID) = 12
        .ColWidth(E_FPS0.E_Qty) = 10
        .ColWidth(E_FPS0.E_Date) = 10
        
        

        .RowHeight(0) = 20
        .RowHeight(-1) = 15
        
        
        .ReDraw = True
    End With
    
    
    

End Sub


Private Sub Text1_Change()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim lotIDTemp As String
lotIDTemp = Trim(Text2.Text)

 If KeyAscii = 13 Then
    
    
    Set oiRS = GetWipSetData(lotIDTemp)
    If (oiRS.RecordCount > 0) Then
    
    DTPicker3.Value = CDate(oiRS.fields("lotdate").Value)
    Text3.Text = IIf(IsNull(oiRS.fields("remark").Value), "", oiRS.fields("remark").Value)

    End If
    
    
    
    
 End If
 

End Sub
