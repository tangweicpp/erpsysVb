VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmSetTime 
   Caption         =   "�Զ���ʱ����Remark �趨"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   12780
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "ʱ���趨"
      TabPicture(0)   =   "Form4.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Remark�趨"
      TabPicture(1)   =   "Form4.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Command6"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame5 
         Caption         =   "�����깤ʱ��"
         Height          =   1095
         Left            =   720
         TabIndex        =   28
         Top             =   3600
         Width           =   9615
         Begin VB.CommandButton cmd 
            Caption         =   "�޸�"
            Height          =   360
            Left            =   6960
            TabIndex        =   34
            Top             =   360
            Width           =   990
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   4800
            TabIndex        =   33
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   133562369
            CurrentDate     =   42612
         End
         Begin VB.TextBox txtText3 
            Height          =   405
            Left            =   1200
            TabIndex        =   31
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   195
            Index           =   1
            Left            =   4080
            TabIndex        =   32
            Top             =   480
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ţ�"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   30
            Top             =   480
            Width           =   720
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "��������"
         Height          =   600
         Left            =   -74280
         TabIndex        =   27
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "�޸�"
         Height          =   1815
         Left            =   -74280
         TabIndex        =   21
         Top             =   2400
         Width           =   9615
         Begin VB.TextBox TxtRemark2 
            Height          =   375
            Left            =   1440
            TabIndex        =   24
            Top             =   1080
            Width           =   5415
         End
         Begin VB.CommandButton Command5 
            Caption         =   "�޸�"
            Height          =   360
            Left            =   6960
            TabIndex        =   23
            Top             =   480
            Width           =   990
         End
         Begin VB.TextBox TxtWafer2 
            Height          =   375
            Left            =   1440
            TabIndex        =   22
            Top             =   480
            Width           =   5415
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remark��"
            Height          =   195
            Left            =   600
            TabIndex        =   26
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WaferID��"
            Height          =   195
            Left            =   600
            TabIndex        =   25
            Top             =   600
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "����"
         Height          =   1695
         Left            =   -74280
         TabIndex        =   14
         Top             =   480
         Width           =   9615
         Begin VB.TextBox TxtRemark 
            Height          =   375
            Left            =   1440
            TabIndex        =   17
            Top             =   960
            Width           =   5415
         End
         Begin VB.CommandButton Command4 
            Caption         =   "���"
            Height          =   360
            Left            =   6960
            TabIndex        =   16
            Top             =   360
            Width           =   990
         End
         Begin VB.TextBox TxtWafer 
            Height          =   375
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   5415
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remark��"
            Height          =   195
            Left            =   600
            TabIndex        =   20
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   4080
            TabIndex        =   19
            Top             =   480
            Width           =   45
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WaferID��"
            Height          =   195
            Left            =   600
            TabIndex        =   18
            Top             =   480
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "����"
         Height          =   975
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   9615
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1200
            TabIndex        =   10
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton Command1 
            Caption         =   "���"
            Height          =   360
            Left            =   6960
            TabIndex        =   9
            Top             =   360
            Width           =   990
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4680
            TabIndex        =   11
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   133562369
            CurrentDate     =   40947
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LOTID��"
            Height          =   195
            Left            =   600
            TabIndex        =   13
            Top             =   480
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ڣ�"
            Height          =   195
            Left            =   4080
            TabIndex        =   12
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�޸�"
         Height          =   1335
         Left            =   720
         TabIndex        =   2
         Top             =   1920
         Width           =   9615
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1200
            TabIndex        =   4
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton Command2 
            Caption         =   "�޸�"
            Height          =   360
            Left            =   6960
            TabIndex        =   3
            Top             =   480
            Width           =   990
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   4680
            TabIndex        =   5
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   133562369
            CurrentDate     =   40947
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LOTID��"
            Height          =   195
            Left            =   600
            TabIndex        =   7
            Top             =   600
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ڣ�"
            Height          =   195
            Left            =   3960
            TabIndex        =   6
            Top             =   600
            Width           =   720
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "��������"
         Height          =   600
         Left            =   600
         TabIndex        =   1
         Top             =   4800
         Width           =   1335
      End
   End
   Begin VB.Label lblLOTID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOTID��"
      Height          =   195
      Left            =   1320
      TabIndex        =   29
      Top             =   4440
      Width           =   630
   End
End
Attribute VB_Name = "FrmSetTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oiRS        As New ADODB.Recordset

Private Sub cmd_Click()
Dim strsql As String
Dim workName As String
Dim RS As New Recordset
Dim cmd As New ADODB.Command
Dim dtTemp As Date

dtTemp = DTPicker2.Value
workName = Trim(txtText3.Text)
If workName = "" Then
MsgBox "�����빤����"
Exit Sub
End If

strsql = "select ORDERNAME from ib_wohistory where ORDERNAME='" & workName & "'" '���жϹ����Ƿ����
 If Cnn.State = 0 Then
    ConOracle
 End If
 
   RS.open strsql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
   If RS.RecordCount <= 0 Then
   MsgBox "Ҫ�޸ĵĹ�����������ȷ�ϣ�"
   Exit Sub
   End If
   
 strsql = "update ib_wohistory set planenddate=to_date('" & dtTemp & "','yyyy-mm-dd') where ORDERNAME='" & workName & "'" '�޸Ĺ����깤������ҪΪMPSWIP�������
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strsql
   cmd.CommandType = adCmdText
   cmd.Execute
   MsgBox "�޸ĳɹ���"

End Sub

Private Sub Command1_Click()
'����
Dim lotIDTemp As String
Dim dtTemp As Date
Dim sqlTemp As String
Dim remarkTemp As String



If Trim(Text1.Text) <> "" Then

    lotIDTemp = Trim(Text1.Text)
    dtTemp = DTPicker1.Value
    remarkTemp = ""
    
    '�ж������Lot�ţ��Ƿ���ȷ
    
    If JudgeLot2(lotIDTemp) Then
    
    
        '�ж��Ƿ���� ��������ʾ��Ϣ
        If Not (JudgeLot(lotIDTemp)) Then
        sqlTemp = "insert into WipreportDate(lotid,lotdate,remark) values ( '" & lotIDTemp & "',to_date('" & dtTemp & "','yyyy-mm-dd'),'" & remarkTemp & "' ) "
        AddSql (sqlTemp)
        MsgBox "��ӳɹ�!"
        
        Else
        
        MsgBox "LotId:" & lotIDTemp & "�Ѵ��ڣ�"
        End If
        
    Else
         MsgBox "LotId:" & lotIDTemp & "��Mesϵͳ�в����ڣ���ȷ��Lot�ţ�"
    
    End If
    

Else
MsgBox "��������LotId!"
End If


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
cmdStr = "  select * from WipreportDateRemark where lotid='" + lotIDTemp + "' "
         
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
cmdStr = "  select * from A_Lotwafers where  waferscribenumber='" + lotIDTemp + "' "
         
         
slectResult = QueryStr(cmdStr)
JudgeWafer2 = slectResult
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

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
DTPicker1.Value = DateTime.Date
DTPicker3.Value = DateTime.Date

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
