VERSION 5.00
Begin VB.Form FrmDelERPQbox 
   Caption         =   "ɾ����ERP���������"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
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
   ScaleHeight     =   6765
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "AA����Ʒ��Ŵ���"
      Height          =   1695
      Left            =   360
      TabIndex        =   13
      Top             =   4320
      Width           =   6255
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2880
         TabIndex        =   19
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1080
         TabIndex        =   18
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton CmdAANG 
         BackColor       =   &H000000FF&
         Caption         =   "ɾ��"
         Height          =   360
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtAANG 
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ɾ����ţ�"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ţ�"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "���"
      Height          =   480
      Left            =   4680
      TabIndex        =   12
      Top             =   3360
      Width           =   990
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "�޸�"
      Height          =   480
      Left            =   2280
      TabIndex        =   11
      Top             =   3360
      Width           =   990
   End
   Begin VB.TextBox TxtQboxNew 
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox TxtQboxOld 
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox TxtContainername 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "ɾ��"
      Height          =   480
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��С��ţ�"
      Height          =   195
      Left            =   6840
      TabIndex        =   9
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "С��ţ�"
      Height          =   195
      Left            =   3600
      TabIndex        =   7
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ţ�"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "36�ͻ�С��ű����"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   1620
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   6480
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ERPС�������ͬ����"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "С��ţ�"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   720
   End
End
Attribute VB_Name = "FrmDelERPQbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAANG_Click()
Dim contanerName As String
contanerName = UCase(Trim(TxtAANG.Text))


Text2.Text = GetAADelQboxNo1(contanerName)

Text3.Text = GetAADelQboxNo2(contanerName)





        
         Set adoCmd = New ADODB.Command
         Set adoCmd.ActiveConnection = Cnn
             adoCmd.CommandText = "del_ON_NGQbox"
             adoCmd.CommandType = adCmdStoredProc
        
          Set adoprm1 = New ADODB.Parameter   '������
          adoprm1.Type = adChar
          adoprm1.Size = 20
          adoprm1.Direction = adParamInput
          adoprm1.Value = contanerName
          adoCmd.Parameters.Append adoprm1
        
          adoCmd.Execute
          
      MsgBox "�ٰ������������ŷֱ��Ƶ����棬��ERP������ɾ��������ͬ��", vbInformation, "������ʾ"

End Sub

Private Sub cmdClear_Click()
TxtContainername.Text = ""
TxtQboxOld.Text = ""
TxtQboxNew.Text = ""



End Sub

Private Sub cmdModify_Click()
Dim containerTemp As String
Dim qboxtemp As String
Dim qboxNewTemp As String

'�ж��������Ϣ�Բ���
containerTemp = UCase(Trim(TxtContainername.Text))
qboxtemp = UCase(Trim(TxtQboxOld.Text))
qboxNewTemp = UCase(Trim(TxtQboxNew.Text))

If JudgeMofidyQboxStatus(containerTemp, qboxtemp) = True Then

'���ж���һ�䣬�Ƿ����ɾ��





Else
     MsgBox "�������������ԭ��Ų���ȷ����ȷ�ϣ�"

End If




End Sub

Public Sub Command1_Click()
Dim qboxtemp As String

qboxtemp = UCase(Trim(Text1.Text))

If qboxtemp <> "" Then


Call DelERPQboxData(qboxtemp)

  MsgBox "�����ɾ����"

Else

     MsgBox "��Ų�����Ϊ�գ�"
     Exit Sub

End If



End Sub

