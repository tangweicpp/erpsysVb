VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_CGVH 
   Caption         =   "�ɹ�ά��"
   ClientHeight    =   11595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15195
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11595
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   10935
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   15255
      Begin VB.CommandButton Command7 
         Caption         =   "��ԭ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "��ԭ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6000
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "��ѯ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5160
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9840
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin FPSpreadADO.fpSpread fpss 
         Height          =   3855
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   6960
         Width           =   15255
         _Version        =   524288
         _ExtentX        =   26908
         _ExtentY        =   6800
         _StockProps     =   64
         DAutoCellTypes  =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   0
         MaxRows         =   0
         SpreadDesigner  =   "For4.frx":0000
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   6000
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ȷ��"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9000
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "For4.frx":0404
         Left            =   1200
         List            =   "For4.frx":040E
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   6000
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   2895
      End
      Begin FPSpreadADO.fpSpread fps 
         Height          =   4455
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   1800
         Width           =   15255
         _Version        =   524288
         _ExtentX        =   26908
         _ExtentY        =   7858
         _StockProps     =   64
         DAutoCellTypes  =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   0
         MaxRows         =   0
         SpreadDesigner  =   "For4.frx":0426
         TextTip         =   2
         AppearanceStyle =   0
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label lb5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ɹ��޸ļ�¼"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lb6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�빺�޸ļ�¼"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   15
         Top             =   6480
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lb4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӧ��"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lb3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ϻ�"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lb2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ����"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lb1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޸ķ�ʽ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   12360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":082A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":147C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":20CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":2D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":3972
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":3CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":4916
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":5568
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":61BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "For4.frx":6E0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1535
      ButtonWidth     =   1455
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ѯ"
            Key             =   "QUE"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "�����޸�"
            Key             =   "MOD"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�����µ�"
            Key             =   "DEL"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�"
            Key             =   "EXIT"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�޸ļ�¼"
            Key             =   "QUERY"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ӧ��"
            Key             =   "Modify"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�빺����"
            Key             =   "Part"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "δ����"
            Key             =   "Delay"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "Frm_CGVH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_click()

    Select Case Combo1.Text

        Case "�빺����"
            lb2 = "�빺����"

        Case "�ɹ�����"
            lb2 = "�ɹ�����"
            
    End Select

End Sub

Private Sub Command1_Click()
    
    ForDe1

End Sub

Private Sub Command2_Click()
    
    ForMod2

End Sub
Private Sub Command3_Click()

    PartN

End Sub

Private Sub Command4_Click()

    ForQuery

End Sub

Private Sub Command5_Click()

    PartN1

End Sub

Private Sub Command6_Click()

    PartN2

End Sub

Private Sub Command7_Click()

    PartN3

End Sub


'��ʼ��
Private Sub Initial()

    fps(0).MaxRows = 0
    fps(0).MaxCols = 0
    fpss(0).MaxRows = 0
    fpss(0).MaxCols = 0
    
    Combo1.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""

End Sub



Private Sub Form_Load()

    With fps(0)

        .Col = -1
        .Row = -1
        .Lock = True
        
        .DAutoSizeCols = DAutoSizeColsBest

    End With
    
    With fpss(0)

        .Col = -1
        .Row = -1
        .Lock = True
        
        .DAutoSizeCols = DAutoSizeColsBest

    End With

End Sub

Private Sub Datapro(strsql As String)

    Dim rs As New ADODB.Recordset
      
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType1(rs)
    Else
        
        MsgBox "��ѯ�����òɹ���Ϣ", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

Private Sub Datapro1(strsql As String)

    Dim rs As New ADODB.Recordset
      
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType3(rs)
    Else
        
        MsgBox "��ѯ�����òɹ���Ϣ", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

'�Ϻż��
Private Sub DataDetect(strPartno1 As String)

     If Get_SqlserverCnt("select b.�Ϻ� from erpdata..tblSmainM2 b  WHERE b.�Ϻ� = '" & strPartno1 & "'") = 0 Then
        MsgBox "û�д��Ϻ�,����������", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

'�빺�����
Private Sub DataDetect1(strBuy1 As String)

    If Get_SqlserverCnt("select a.�빺����� from erpbase..tblCRequest a  WHERE a.�빺����� = '" & strBuy1 & "'") = 0 Then
        MsgBox "û�д��빺����,����������", vbInformation, "��ʾ"
        Exit Sub
    End If

End Sub

'�ɹ������
Private Sub DataDetect2(strBuy2 As String)

    If Get_SqlserverCnt("select a.�ɹ������ from erpbase..tblCPurDataSub a  WHERE a.�ɹ������ = '" & strBuy2 & "'") = 0 Then
        MsgBox "û�д˲ɹ�����,����������", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key

        Case "QUE"
            Initial
            
            Toolbar1.Buttons(3).Enabled = True
            Toolbar1.Buttons(3).Caption = "�����޸�"
            Toolbar1.Buttons(3).Image = 6
            Toolbar1.Buttons(3).Visible = True
            
            Combo1.Visible = True
            
            lb1.Visible = True
            lb2.Visible = True
            lb3.Visible = True
            lb4.Visible = False
            lb5.Visible = False
            lb6.Visible = False
            
            Text1.Visible = True
            Text2.Visible = True
            Text3.Visible = False
            
            Command1.Visible = False
            Command2.Visible = False
            Command3.Visible = False
            '��ѯ
            Command4.Visible = True
            Command5.Visible = False
            Command6.Visible = False
            Command7.Visible = False
        
        Case "MOD"
            
            Combo1.Visible = False
            
            lb1.Visible = False
            
            Command4.Visible = False
            
            ForMod1
               
        Case "DEL"
           
            Initial
            
            Toolbar1.Buttons(3).Enabled = False
            
            Combo1.Visible = True
            
            lb1.Visible = True
            lb2.Visible = True
            lb3.Visible = True
            lb4.Visible = False
            lb5.Visible = False
            lb6.Visible = False
            
            Text1.Visible = True
            Text2.Visible = True
            Text3.Visible = False
            
            '�����µ�
            Command1.Visible = True
            Command2.Visible = False
            Command3.Visible = False
            Command4.Visible = False
            Command5.Visible = False
            Command6.Visible = False
            Command7.Visible = False

        Case "EXIT"
            Unload Me
            
        Case "QUERY"
            Initial
            
            Toolbar1.Buttons(3).Enabled = False
           
            Combo1.Visible = False
            
            lb5 = "�ɹ��޸ļ�¼"
            lb6 = "�빺�޸ļ�¼"
            lb1.Visible = False
            lb2.Visible = False
            lb3.Visible = False
            lb4.Visible = False
            lb5.Visible = True
            lb6.Visible = True
            
            Text1.Visible = False
            Text2.Visible = False
            Text3.Visible = False
            
            Command1.Visible = False
            Command2.Visible = False
            Command3.Visible = False
            Command4.Visible = False
            Command5.Visible = False
            Command6.Visible = False
            Command7.Visible = False
        
            Query1
            
        Case "Modify"
            Initial
            
            Toolbar1.Buttons(3).Enabled = False
            
            Combo1.Visible = False
            lb2 = "�ɹ�����"
            lb1.Visible = False
            lb2.Visible = True
            lb3.Visible = False
            lb4.Visible = True
            lb5.Visible = False
            lb6.Visible = False
          
            Text1.Visible = True
            Text2.Visible = False
            Text3.Visible = True
            
            Command1.Visible = False
            '��Ӧ���޸�
            Command2.Visible = True
            Command3.Visible = False
            Command4.Visible = False
            Command5.Visible = False
            Command6.Visible = False
            Command7.Visible = False

        Case "Part"
            Initial
            
            Toolbar1.Buttons(3).Enabled = False
           
            Combo1.Visible = False
            
            lb2 = "�빺����"
            lb1.Visible = False
            lb2.Visible = True
            lb3.Visible = True
            lb4.Visible = False
            lb5.Visible = False
            lb6.Visible = False
            
            Text1.Visible = True
            Text2.Visible = True
            Text3.Visible = False
            
            Command1.Visible = False
            Command2.Visible = False
            '�빺������
            Command3.Visible = True
            Command4.Visible = False
            '��ԭ
            Command5.Visible = True
            Command6.Visible = False
            Command7.Visible = False
        
        Case "Delay"
        
            Initial
            
            Toolbar1.Buttons(3).Enabled = False
           
            Combo1.Visible = True
        
            lb1.Visible = True
            lb2.Visible = True
            lb3.Visible = True
            lb4.Visible = False
            lb5.Visible = False
            lb6.Visible = False
            
            Text1.Visible = True
            Text2.Visible = True
            Text3.Visible = False
            
            Command1.Visible = False
            Command2.Visible = False
            Command3.Visible = False
            Command4.Visible = False
            Command5.Visible = False
            'δ����
            Command6.Visible = True
            '��ԭ
            Command7.Visible = True

    End Select

End Sub

Private Sub ForQuery()

    Dim rs        As New ADODB.Recordset

    Dim strBuy    As String
    
    Dim strPartno As String

    Dim strsql    As String
    
    If Combo1.Text = "" Then
    
        MsgBox "��ѡ���޸ķ�ʽ", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    If Text1.Text = "" Then
    
        MsgBox "����������", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    If Text2.Text = "" Then
    
        MsgBox "�������Ϻ�", vbInformation, "��ʾ"
        Exit Sub
    Else
        strBuy = Trim$(Text1.Text)
     
        strPartno = Trim$(Text2.Text)

        If lb2 = "�빺����" Then
            
            Call DataDetect1(strBuy)
            
            Call DataDetect(strPartno)
            
            strsql = "select '' as  '��',a.�ɹ������,a.�ɹ������,a.�빺�����,a.�빺�����,a.���ϱ��,b.�Ϻ�,a.�ɹ�����,a.��׼�ɹ�����,a.����,a.���,a.�ұ� from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0'"
        Else
            
            Call DataDetect2(strBuy)
            
            Call DataDetect(strPartno)
            
            strsql = "select '' as  '��',a.�ɹ������,a.�ɹ������,a.�빺�����,a.�빺�����,a.���ϱ��,b.�Ϻ�,a.�ɹ�����,a.��׼�ɹ�����,a.����,a.���,a.�ұ� from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�ɹ������ = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0'"
    
        End If

    End If
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        lb5 = "��ѯ���"
        Call ListDataType2(rs)
    Else
        
        MsgBox "��ѯ��������", vbInformation, "��ʾ"
        Exit Sub

    End If

End Sub

'�Ϻ��޸ĺ����ϳ���
Private Sub ForQuery1(strBuy As String, strPartno As String)

    Dim rs        As New ADODB.Recordset
    
    Dim strsql    As String
            
    strsql = "select '' as  '��',a.�ɹ������,a.�ɹ������,a.�빺�����,a.�빺�����,a.���ϱ��,b.�Ϻ�,a.�ɹ�����,a.��׼�ɹ�����,a.����,a.���,a.�ұ� from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�ɹ������ = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0'"
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        lb5 = "��ѯ���"
        Call ListDataType2(rs)
    Else
        
        MsgBox "��ѯ��������", vbInformation, "��ʾ"
        Exit Sub

    End If
    
End Sub

'�޸ļ�¼
Private Sub Query1()

    Dim rs     As New ADODB.Recordset

    Dim strsql As String
    
    strsql = "select a.�޸���,a.�޸ķ�ʽ,a.�޸�״̬,a.�޸�ʱ��,a.�ɹ������,a.�ɹ������,a.�ɹ������,a.�빺�����,a.���ϱ��,a.Լ����������,a.�빺����,a.�ɹ�����,a.��׼�ɹ�����,a.����,a.���,a.�ұ� from erptemp.dbo.ksrequisition a order by a.�޸�ʱ��"
    
    fps(0).MaxRows = 0
    
    If rs.State = adStateOpen Then rs.Close
    
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType1(rs)
    Else
        
        MsgBox "��ѯ�����޸ļ�¼", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strsql = "select b.�޸���,b.�޸ķ�ʽ,b.�޸�״̬,b.�޸�ʱ��,b.�빺�����,b.�빺�����,b.���ϱ��,b.�빺��,b.�빺����,b.��������,b.�빺����,b.�������,b.�Ƿ����,b.��׼����,b.��������,b.�ɹ�Ա from erptemp.dbo.ksbuy b order by b.�޸�ʱ��"
    
    fpss(0).MaxRows = 0
    
    If rs.State = adStateOpen Then rs.Close
    
    rs.Open strsql, INIadoCon2, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then  '��ʾ��������
        Call ListDataType3(rs)
    Else
        
        MsgBox "��ѯ�����޸ļ�¼", vbInformation, "��ʾ"
        Exit Sub

    End If
      
    
End Sub

Private Sub ListDataType1(rs As ADODB.Recordset)
   
    With fps(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With

End Sub

Private Sub ListDataType2(rs As ADODB.Recordset)

    Dim i As Long
   
    With fps(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With
    
    With fps(0)

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            .ColWidth(1) = 2
            .CellType = CellTypeCheckBox
        Next

    End With

End Sub

Private Sub ListDataType3(rs As ADODB.Recordset)
   
    With fpss(0)
        
        .MaxRows = 0

        Set .DataSource = rs

    End With

End Sub

'�ɹ���ϸ����
Private Sub Databackup1(strstyle As String, _
                        strstyle1 As String, _
                        strBuy As String, _
                        strPartno As String, _
                        strno As Integer)

    Dim strid      As Integer
    
    Dim strid1     As Integer
    
    Dim Userrecord As String
    
    Userrecord = gUserName '��ȡ��¼��

    strid = Get_SqlStr("select isnull(max(id),0) from erptemp.dbo.ksrequisition where �޸ķ�ʽ = '" & strstyle & "' ")
    
    strid1 = strid + 1
    
    AddSql2 (" insert into erptemp.dbo.ksrequisition(id,�޸���,�޸ķ�ʽ,�޸�״̬,�޸�ʱ��,�Ƿ����,�ɹ������,�ɹ������,�빺�����,�빺�����,���ϱ��,Լ����������,�빺����,�ɹ�����,��׼�ɹ�����,����,���,�ұ�) select '" & strid1 & "','" & Userrecord & "','" & strstyle & "','" & strstyle1 & "',GetDate(),a.�Ƿ����,a.�ɹ������,a.�ɹ������,a.�빺�����,a.�빺�����,a.���ϱ��,a.Լ����������,a.�빺����,a.�ɹ�����,a.��׼�ɹ�����,a.����,a.���,a.�ұ�  from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�ɹ������ = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '" & strno & "'")

End Sub

'�빺��ϸ����
Private Sub Databackup2(strstyle As String, _
                        strstyle1 As String, _
                        strBuy As String, _
                        strPartno As String, _
                        strno As Integer)

    Dim strid      As Integer
    
    Dim strid1     As Integer

    Dim Userrecord As String
    
    Userrecord = gUserName '��ȡ��¼��
    
    strid = Get_SqlStr("select isnull(max(id),0) from erptemp.dbo.ksbuy where �޸ķ�ʽ = '" & strstyle & "' ")
        
    strid1 = strid + 1
        
    AddSql2 (" insert into erptemp.dbo.ksbuy(id,�޸���,�޸ķ�ʽ,�޸�״̬,�޸�ʱ��,�빺�����,�빺�����,���ϱ��,�빺��,�빺����,��������,�빺����,�������,�Ƿ����,��׼����,��������,�ɹ�Ա) select '" & strid1 & "','" & Userrecord & "','" & strstyle & "','" & strstyle1 & "',GetDate(),a.�빺�����,a.�빺�����,a.���ϱ��,a.�빺��,a.�빺����,a.��������,a.�빺����,a.�������,a.�Ƿ����,a.��׼����,a.��������,a.�ɹ�Ա from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺�����  = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '" & strno & "' ")

End Sub

Private Sub ForMod1()

    Dim i         As Integer

    Dim m         As Integer

    Dim J         As Integer

    Dim strstyle  As String

    Dim strstyle1 As String

    Dim strstyle2 As String
    
    Dim strpu1    As String

    Dim strpu2    As String

    Dim strpu3    As String

    Dim strpu4    As String

    Dim strpu5    As String

    Dim strpu6    As String

    Dim strpu7    As String

    Dim strpu8    As String

    Dim strpu9    As String

    Dim strpu10   As String

    Dim strpu11   As String

    Dim strpua    As String
    
    Dim strPartno As String
    
    Dim bFlag     As Boolean
    
    Dim strsql    As String
    
    strPartno = Trim$(Text2.Text)

    If Toolbar1.Buttons(3).Caption <> "ȷ���޸�" Then

        With fps(0)

            For i = 1 To .MaxRows
                .Row = i
                .Col = 1
                .Lock = False
      
                For m = 7 To 10
            
                    .Col = m
                    .Lock = False
      
                Next
    
            Next
        
        End With
    
        Toolbar1.Buttons(3).Caption = "ȷ���޸�"
        Toolbar1.Buttons(3).Image = 10
        
        Exit Sub

    End If

    bFlag = False
    
    With fps(0)

        If .MaxRows = 0 Then
            MsgBox "û������", vbInformation, "��ʾ"
            Exit Sub

        End If

        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
    
            J = 0

            If .Text = "1" Then
            
                J = J + 1
                bFlag = True
    
                .Col = 2
                strpu1 = Trim$(.Text)
                
                .Col = 3
                strpu2 = Trim$(.Text)
                
                .Col = 4
                strpu3 = Trim$(.Text)
                
                .Col = 5
                strpu4 = Trim$(.Text)
                
                .Col = 6
                strpu5 = Trim$(.Text)
                
                '�Ϻ�modify
                .Col = 7
                strpu6 = Trim$(.Text)
                
                If Trim$(strpu6) <> Trim$(strPartno) Then
                    
                    If Get_SqlserverCnt("select distinct �Ϻ� from erpdata..tblSmainM2 where �Ϻ� = '" & strpu6 & "'") = 0 Then
                                                
                        MsgBox "û�д��Ϻ���ȷ��", vbInformation, "��ʾ"
                        Exit Sub

                    End If
                    
                    '��ȡ�µ����ϱ��
                    strpua = Get_SqlStr("select distinct ���ϱ�� from erpdata..tblSmainM2 where �Ϻ� = '" & strpu6 & "'")
                    Else
                 
                    strpua = Get_SqlStr("select distinct ���ϱ�� from erpdata..tblSmainM2 where �Ϻ� = '" & strPartno & "'")

                    
                End If
                
                '�ɹ�����modify
                
                .Col = 8
                strpu7 = Trim$(.Text)
                
                '��׼�ɹ�����modify
                .Col = 9
                strpu8 = Trim$(.Text)
                
                '����modify
                .Col = 10
                strpu9 = Trim$(.Text)
                
                .Col = 11
                strpu10 = Trim$(.Text)
                
                .Col = 12
                strpu11 = Trim$(.Text)
                
                strstyle = "�����޸�"
                
                strstyle1 = "�޸�ǰ"
                
                strstyle2 = "�޸ĺ�"
                
                Call Databackup1(strstyle, strstyle1, strpu1, strPartno, 0)
                
                AddSql2 (" update a set a.���ϱ�� = '" & strpua & "',a.�ɹ����� = '" & strpu7 & "',a.��׼�ɹ����� = '" & strpu8 & "',a.���� = '" & strpu9 & "' from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�ɹ������ = '" & strpu1 & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0'")
                
                Call Databackup1(strstyle, strstyle2, strpu1, strpu6, 0)
                
                Call Databackup2(strstyle, strstyle1, strpu3, strPartno, 0)

                AddSql2 ("update a set a.���ϱ�� = '" & strpua & "',a.��׼���� = '" & strpu8 & "',a.�빺���� = '" & strpu8 & "',a.�������� = '" & strpu8 & "' from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strpu3 & "' and  b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0' ")
                
                Call Databackup2(strstyle, strstyle2, strpu3, strpu6, 0)
            
            End If
            
        Next
        
        If bFlag = False And J = 0 Then
            MsgBox "��ѡ��Ҫ�޸ĵ���", vbInformation, "��ʾ"
            Exit Sub
            
        End If
    
    End With
    
    MsgBox "�޸ĳɹ�", vbInformation, "��ʾ"

    Toolbar1.Buttons(3).Caption = "�����޸�"
    Toolbar1.Buttons(3).Image = 6
    
    If Trim$(strpu6) = Trim$(strPartno) Then
    
        lb5 = "�ɹ���ϸ����"
        lb5.Visible = True
        
        ForQuery
        
        strsql = "select a.�Ƿ����,a.�빺�����,a.�빺�����,a.���ϱ��,a.�빺��,a.�빺����,a.��������,a.�빺����,a.�������,a.��׼����,a.��������,a.�ɹ�Ա from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺�����  = '" & strpu3 & "' and  a.�Ƿ���� = '0' and b.�Ϻ� = '" & strPartno & "'"
        
        lb6 = "�빺��ϸ����"
        lb6.Visible = True
        
        Call Datapro1(strsql)
    Else
        lb5 = "�ɹ���ϸ����"
        lb5.Visible = True
        
        Call ForQuery1(strpu1, strpu6)
        
        strsql = "select a.�Ƿ����,a.�빺�����,a.�빺�����,a.���ϱ��,a.�빺��,a.�빺����,a.��������,a.�빺����,a.�������,a.��׼����,a.��������,a.�ɹ�Ա from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺�����  = '" & strpu3 & "' and  a.�Ƿ���� = '0' and b.�Ϻ� = '" & strpu6 & "'"
        
        lb6 = "�빺��ϸ����"
        lb6.Visible = True
        Call Datapro1(strsql)

    End If

End Sub

Private Sub ForDe1()
    
    Dim strBuy    As String
    
    Dim strBuy1   As String

    Dim strPartno As String
    
    Dim strstyle  As String

    Dim strstyle1 As String
    
    Dim strstyle2 As String
 
    Dim strsql    As String
    
    If Combo1.Text = "" Then
        MsgBox "��ѡ���޸ķ�ʽ", vbInformation, "��ʾ"
        Exit Sub
            
    End If
            
    If Combo1.Text <> "�ɹ�����" And Combo1.Text <> "�빺����" Then
            
        MsgBox "��ѡ����ȷ���޸ķ�ʽ", vbInformation, "��ʾ"
            
        Exit Sub
            
    End If
    
    If Text1.Text = "" Then
        MsgBox "�����뵥��", vbInformation, "��ʾ"
        Exit Sub

    End If
       
    strBuy = Trim$(Text1.Text)
    
    If lb2 = "�ɹ�����" Then
    
        Call DataDetect2(strBuy)
    
    Else

        Call DataDetect1(strBuy)

    End If

    If Text2.Text = "" Then
        MsgBox "�������Ϻ�", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strPartno = Trim$(Text2.Text)
    
    Call DataDetect(strPartno)

    If lb2 = "�ɹ�����" Then
        If Get_SqlserverCnt("select a.�ɹ������,a.�빺�����,a.���ϱ�� from erpbase..tblCPurDataSub a inner join erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� WHERE a.�ɹ������ = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0'") = 0 Then
            MsgBox "û�д˱�����,����������", vbInformation, "��ʾ"
            Exit Sub

        End If
    
        strsql = "select a.�ɹ������,a.�빺�����,a.���ϱ�� from erpbase..tblCPurDataSub a inner join erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� WHERE a.�ɹ������ = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0'"
    Else

        If Get_SqlserverCnt("select a.�ɹ������,a.�빺�����,a.���ϱ�� from erpbase..tblCPurDataSub a inner join erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� WHERE a.�빺����� = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0'") = 0 Then
            MsgBox "û�д˱�����,����������", vbInformation, "��ʾ"
            Exit Sub

        End If

        strsql = "select a.�ɹ������,a.�빺�����,a.���ϱ�� from erpbase..tblCPurDataSub a inner join erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� WHERE a.�빺����� = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0'"

    End If
    
    lb5 = "�޸��е�����"
    lb5.Visible = True
    
    Call Datapro(strsql)
    
    strstyle = "�����µ�"
    
    strstyle1 = "�޸�ǰ"
        
    strstyle2 = "�޸ĺ�"
    
    If lb2 = "�ɹ�����" Then
        
        '��ȡ�빺�����
        strBuy1 = Get_SqlStr("select distinct c.�빺����� from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.���ϱ�� = d.���ϱ�� WHERE c.�ɹ������ = '" & strBuy & "' and d.�Ϻ� = '" & strPartno & "' and c.�Ƿ���� = '0'")
                
        Call Databackup1(strstyle, strstyle1, strBuy, strPartno, 0)
        '�ɹ���ϸ����Ϣ����'
        AddSql2 (" update a set a.�Ƿ���� = '1'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�ɹ������ = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0'")
    
        '�޸ĺ������backup
             
        Call Databackup1(strstyle, strstyle2, strBuy, strPartno, 1)
    
        '�빺��ϸ������backup
                
        Call Databackup2(strstyle, strstyle1, strBuy1, strPartno, 0)
        
        '�빺��ϸ����Ϣ����
        AddSql2 ("update a set a.�������� = '0',a.������� = '0' from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy1 & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0'")
        
        '�޸ĺ������backup

        Call Databackup2(strstyle, strstyle2, strBuy1, strPartno, 0)
        
        strsql = "select a.�빺�����,a.�빺�����,a.���ϱ��,a.�������,a.�������� from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where b.�Ϻ� = '" & strPartno & "' and a.�빺����� = '" & strBuy1 & "' and a.�Ƿ���� = '0'"
    Else
        
        '��ȡ�ɹ������
        strBuy1 = Get_SqlStr("select distinct c.�ɹ������ from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.���ϱ�� = d.���ϱ�� WHERE c.�빺����� = '" & strBuy & "' and d.�Ϻ� = '" & strPartno & "' and c.�Ƿ���� = '0'")
        
        Call Databackup1(strstyle, strstyle1, strBuy1, strPartno, 0)
        
        '�ɹ���ϸ����Ϣ����'
        AddSql2 (" update a set a.�Ƿ���� = '1'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�ɹ������ = '" & strBuy1 & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0' ")
    
        '�޸ĺ������backup
        
        Call Databackup1(strstyle, strstyle2, strBuy1, strPartno, 1)
        
        '�빺��ϸ������backup
                
        Call Databackup2(strstyle, strstyle1, strBuy, strPartno, 0)
        
        '�빺��ϸ����Ϣ����
        AddSql2 ("update a set a.�������� = '0',a.������� = '0' from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0' ")
        
        '�޸ĺ������backup
                
        Call Databackup2(strstyle, strstyle2, strBuy, strPartno, 0)
        
        strsql = "select a.�빺�����,a.�빺�����,a.���ϱ��,a.�������,a.�������� from erpbase..tblCRequest a inner join erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0' "

    End If
    
    lb6 = "�޸ĺ���빺����"
    lb6.Visible = True
    
    Call Datapro1(strsql)
    
    MsgBox "�����µ��ɹ�", vbInformation, "��ʾ"

End Sub

Private Sub ForMod2()

    Dim strBuy      As String
    
    Dim strprovider As String
    
    Dim strsupply   As String

    Dim strsupply1  As String

    Dim strsql      As String
    
    Dim Userrecord  As String
    
    
    
    Userrecord = gUserName '��ȡ��¼��
    
    If Text1.Text = "" Then
        MsgBox "������ɹ�����", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strBuy = Trim$(Text1.Text)
    
    Call DataDetect2(strBuy)
    
    If Text3.Text = "" Then
        MsgBox "�����빩Ӧ������", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strprovider = Trim$(Text3.Text)
    
    If Get_SqlserverCnt("SELECT distinct ��Ӧ�̱�� FROM ERPBASE..tblSupplierData where ��Ӧ������ = '" & strprovider & "'") = 0 Then
        MsgBox "û�д˹�Ӧ����Ϣ,����������", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strsupply = Get_SqlStr("SELECT distinct ��Ӧ�̱�� FROM ERPBASE..tblSupplierData where ��Ӧ������ = '" & strprovider & "'")
    
    strsupply1 = Get_SqlStr("SELECT distinct ��Ӧ�̱�� FROM erpbase..tblcpurdata where �ɹ������ = '" & strBuy & "' and �Ƿ���� = '0'")

    If Trim$(strsupply) = Trim$(strsupply1) Then
        MsgBox "��Ӧ�̱���Ѿ�����һ�������޸ģ�", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    '����backup
    AddSql2 ("insert into erptemp.dbo.kspur(�޸���,�޸ķ�ʽ,�޸�״̬,�޸�ʱ��,�ɹ�����,��Ӧ�̱��) select '" & Userrecord & "','��Ӧ�̸���','�޸�ǰ',GetDate(),�ɹ������,��Ӧ�̱�� from erpbase..tblcpurdata where �ɹ������ = '" & strBuy & "' and �Ƿ���� = '0'")
    
    AddSql2 ("update erpbase..tblcpurdata set ��Ӧ�̱�� = '" & strsupply & "' where �ɹ������ = '" & strBuy & "' and �Ƿ���� = '0'")
    
    AddSql2 ("insert into erptemp.dbo.kspur(�޸���,�޸ķ�ʽ,�޸�״̬,�޸�ʱ��,�ɹ�����,��Ӧ�̱��) select '" & Userrecord & "','��Ӧ�̸���','�޸ĺ�',GetDate(),�ɹ������,��Ӧ�̱�� from erpbase..tblcpurdata where �ɹ������ = '" & strBuy & "' and �Ƿ���� = '0'")
    '�޸ĺ����ϳ���
    
    strsql = "select distinct m.�ɹ�����,m.��Ӧ�̱�� as �޸�ǰ��Ӧ�̱��,h.��Ӧ������ as �޸�ǰ��Ӧ������,n.��Ӧ�̱�� as �޸ĺ�Ӧ�̱��,'" & strprovider & "' as  �޸ĺ�Ӧ������ from erptemp.dbo.kspur m inner join erpbase..tblcpurdata n on m.�ɹ����� = n.�ɹ������ left join  ERPBASE..tblSupplierData h on h.��Ӧ�̱�� = m.��Ӧ�̱�� where m.�ɹ����� = '" & strBuy & "' and m.�޸�״̬ = '�޸�ǰ' and n.�Ƿ���� = '0' "

    '�޸�ǰ���ϳ���
    lb5 = "�޸�״̬����"
    lb5.Visible = True
    
    Call Datapro(strsql)

    '�޸ļ�¼
    strsql = "select * from erptemp.dbo.kspur where 1 = 1  "
    
    lb6 = "��ʷ�޸ļ�¼"
    lb6.Visible = True
    Call Datapro1(strsql)
    
    MsgBox "��Ӧ���޸ĳɹ�", vbInformation, "��ʾ"
 
End Sub

Private Sub PartN()

    Dim strBuy    As String
    
    Dim strPartno As String

    Dim strstyle  As String

    Dim strstyle1 As String
    
    Dim strstyle2 As String
    
    Dim strsql    As String
    
    Dim strid1 As String

    strstyle = "�빺����"
    
    strstyle1 = "�޸�ǰ"
    
    strstyle2 = "�޸ĺ�"
    
    If Text1.Text = "" Then
        MsgBox "�������빺����", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strBuy = Trim$(Text1.Text)
    
    Call DataDetect1(strBuy)

    If Text2.Text = "" Then
        MsgBox "�������Ϻ�", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strPartno = Trim$(Text2.Text)
    
    Call DataDetect(strPartno)
    
    If Get_SqlserverCnt("select a.�빺����� from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� WHERE a.�빺����� = '" & strBuy & "' and  a.�Ƿ���� = '0' and b.�Ϻ� = '" & strPartno & "'") = 0 Then
        MsgBox "�˱��빺���Ѿ�����,��ȷ��", vbInformation, "��ʾ"
        Exit Sub
        
    End If
    
    If Get_SqlserverCnt("select a.�ɹ������ from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ��  WHERE a.�빺����� = '" & strBuy & "' and  a.�Ƿ���� = '0' and b.�Ϻ� = '" & strPartno & "' ") <> 0 Then
        MsgBox "�˱��빺���Ѿ����вɹ�����,��ȷ��", vbInformation, "��ʾ"
        Exit Sub
        
    End If
    
    Call Databackup2(strstyle, strstyle1, strBuy, strPartno, 0)
    
    strid1 = Get_SqlStr("select isnull(max(id),0) from erptemp.dbo.ksbuy where �޸ķ�ʽ = '" & strstyle & "' ")
    
    strsql = "select �Ƿ����,�빺�����,�빺�����,���ϱ��,�빺��,�빺����,��������,�빺����,�������,��׼����,��������,�ɹ�Ա from erptemp.dbo.ksbuy where �빺�����  = '" & strBuy & "'and �޸�״̬ = '�޸�ǰ' and id = '" & strid1 & "'  order by �빺����� "
    
    lb5 = "�޸�ǰ������"
    lb5.Visible = True
    Call Datapro(strsql)
    
    AddSql2 (" update a set a.�Ƿ���� = '1'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0' ")
    
    Call Databackup2(strstyle, strstyle2, strBuy, strPartno, 1)
    
    strsql = "select a.�Ƿ����,a.�빺�����,a.�빺�����,a.���ϱ��,a.�빺��,a.�빺����,a.��������,a.�빺����,a.�������,a.��׼����,a.��������,a.�ɹ�Ա from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺�����  = '" & strBuy & "' and  a.�Ƿ���� = '1' and b.�Ϻ� = '" & strPartno & "' order by a.�빺����� "
        
    lb6 = "�޸ĺ������"
    lb6.Visible = True
     
    Call Datapro1(strsql)
     
    MsgBox "�Ѿ�����", vbInformation, "��ʾ"
    
End Sub

'��ԭ

Private Sub PartN1()

    Dim strBuy    As String
    
    Dim strPartno As String
    
    Dim strstyle  As String

    Dim strstyle1 As String
    
    Dim strstyle2 As String
    
    Dim strsql    As String
    
    strstyle = "�빺���ϻ�ԭ"
    
    strstyle1 = "�޸�ǰ"
    
    strstyle2 = "�޸ĺ�"
    
    If Text1.Text = "" Then
        MsgBox "�������빺����", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strBuy = Trim$(Text1.Text)
    
    Call DataDetect1(strBuy)

    If Text2.Text = "" Then
        MsgBox "�������Ϻ�", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strPartno = Trim$(Text2.Text)
    
    Call DataDetect(strPartno)
    
    If Get_SqlserverCnt("select a.�빺����� from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� WHERE a.�빺����� = '" & strBuy & "' and  a.�Ƿ���� = '1' and b.�Ϻ� = '" & strPartno & "'") = 0 Then
        MsgBox "�˱��빺��û�����ϼ�¼,��ȷ��", vbInformation, "��ʾ"
        Exit Sub
        
    End If
    
    If Get_SqlserverCnt("select a.�ɹ������ from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ��  WHERE a.�빺����� = '" & strBuy & "' and  a.�Ƿ���� = '0' and b.�Ϻ� = '" & strPartno & "' ") <> 0 Then
        MsgBox "�˱��빺���Ѿ����вɹ�����,��ȷ��", vbInformation, "��ʾ"
        Exit Sub
        
    End If
    
    Call Databackup2(strstyle, strstyle1, strBuy, strPartno, 1)
    
    Dim strid1    As String
strid1 = Get_SqlStr("select isnull(max(id),0) from erptemp.dbo.ksbuy where �޸ķ�ʽ = '" & strstyle & "' ")
    
    strsql = "select �Ƿ����,�빺�����,�빺�����,���ϱ��,�빺��,�빺����,��������,�빺����,�������,��׼����,��������,�ɹ�Ա from erptemp.dbo.ksbuy where �빺�����  = '" & strBuy & "'and �޸�״̬ = '�޸�ǰ' and id = '" & strid1 & "'  order by �빺����� "
    
    lb5 = "�޸�ǰ������"
    lb5.Visible = True
    Call Datapro(strsql)
    
    AddSql2 (" update a set a.�Ƿ���� = '0'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '1' ")
     
    Call Databackup2(strstyle, strstyle2, strBuy, strPartno, 0)
     
    strsql = "select a.�Ƿ����,a.�빺�����,a.�빺�����,a.���ϱ��,a.�빺��,a.�빺����,a.��������,a.�빺����,a.�������,a.��׼����,a.��������,a.�ɹ�Ա from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺�����  = '" & strBuy & "' and  a.�Ƿ���� = '0' and b.�Ϻ� = '" & strPartno & "' order by a.�빺����� "
        
    lb6 = "�޸ĺ������"
    lb6.Visible = True
     
    Call Datapro1(strsql)
        
    MsgBox "��ԭ�ɹ�", vbInformation, "��ʾ"

End Sub

Private Sub PartN2()
    
    Dim strBuy    As String
    
    Dim strBuy1   As String
    
    Dim strPartno As String
    
    Dim strstyle  As String

    Dim strstyle1 As String
    
    Dim strstyle2 As String
     
    strstyle = "δ��������"
    
    strstyle1 = "�޸�ǰ"
    
    strstyle2 = "�޸ĺ�"
     
    If Combo1.Text = "" Then
        MsgBox "��ѡ���޸ķ�ʽ", vbInformation, "��ʾ"
        Exit Sub
            
    End If
            
    If Combo1.Text <> "�ɹ�����" And Combo1.Text <> "�빺����" Then
            
        MsgBox "��ѡ����ȷ���޸ķ�ʽ", vbInformation, "��ʾ"
            
        Exit Sub
            
    End If
    
    If Text1.Text = "" Then
        MsgBox "�����뵥��", vbInformation, "��ʾ"
        Exit Sub

    End If
       
    strBuy = Trim$(Text1.Text)
    
    If lb2 = "�ɹ�����" Then
    
        Call DataDetect2(strBuy)
    
    Else

        Call DataDetect1(strBuy)

    End If

    If Text2.Text = "" Then
        MsgBox "�������Ϻ�", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strPartno = Trim$(Text2.Text)
    
    Call DataDetect(strPartno)
    
    If lb2 = "�빺����" Then
    
        If Get_SqlserverCnt("select a.�빺����� from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� WHERE a.�빺����� = '" & strBuy & "' and  a.�Ƿ���� = '0' and b.�Ϻ� = '" & strPartno & "'") = 0 Then
            MsgBox "�˱��빺���Ѿ�����,��ȷ��", vbInformation, "��ʾ"
            Exit Sub
        
        End If
    
        If Get_SqlserverCnt("select a.�ɹ������ from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ��  WHERE a.�빺����� = '" & strBuy & "' and  a.�Ƿ���� = '0' and b.�Ϻ� = '" & strPartno & "' ") = 0 Then
            MsgBox "�˱ʲɹ����Ѿ�����,��ȷ��", vbInformation, "��ʾ"
            Exit Sub
    
        End If
                
        '��ȡ�ɹ�����
        strBuy1 = Get_SqlStr("select distinct c.�ɹ������ from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.���ϱ�� = d.���ϱ�� WHERE c.�빺����� = '" & strBuy & "' and d.�Ϻ� = '" & strPartno & "'")
        
        Call Databackup2(strstyle, strstyle1, strBuy, strPartno, 0)
        
        AddSql2 (" update a set a.�Ƿ���� = '1'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0' ")
        
        Call Databackup2(strstyle, strstyle2, strBuy, strPartno, 1)
        
        Call Databackup1(strstyle, strstyle1, strBuy1, strPartno, 0)
        
        AddSql2 (" update a set a.�Ƿ���� = '1'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0' ")
           
        Call Databackup1(strstyle, strstyle2, strBuy1, strPartno, 1)
    
    Else
    
        '��ȡ�빺����
        strBuy1 = Get_SqlStr("select distinct c.�빺����� from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.���ϱ�� = d.���ϱ�� WHERE c.�ɹ������ = '" & strBuy & "' and d.�Ϻ� = '" & strPartno & "'")

        If Get_SqlserverCnt("select a.�빺����� from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� WHERE a.�빺����� = '" & strBuy1 & "' and  a.�Ƿ���� = '0' and b.�Ϻ� = '" & strPartno & "'") = 0 Then
            MsgBox "�˱��빺���Ѿ�����,��ȷ��", vbInformation, "��ʾ"
            Exit Sub
        
        End If
    
        If Get_SqlserverCnt("select a.�ɹ������ from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ��  WHERE a.�빺����� = '" & strBuy1 & "' and  a.�Ƿ���� = '0' and b.�Ϻ� = '" & strPartno & "' ") = 0 Then
            MsgBox "�˱ʲɹ����Ѿ�����,��ȷ��", vbInformation, "��ʾ"
            Exit Sub
    
        End If

        Call Databackup2(strstyle, strstyle1, strBuy1, strPartno, 0)
        
        AddSql2 (" update a set a.�Ƿ���� = '1'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy1 & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0' ")
      
        Call Databackup2(strstyle, strstyle2, strBuy1, strPartno, 1)
        
        Call Databackup1(strstyle, strstyle1, strBuy, strPartno, 0)
        
        AddSql2 (" update a set a.�Ƿ���� = '1'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy1 & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '0' ")
        
        Call Databackup1(strstyle, strstyle2, strBuy, strPartno, 1)

    End If
    
    MsgBox "�Ѿ�����", vbInformation, "��ʾ"

End Sub

Private Sub PartN3()

    Dim strBuy    As String
    
    Dim strBuy1   As String
    
    Dim strPartno As String
    
    Dim strstyle  As String

    Dim strstyle1 As String
    
    Dim strstyle2 As String
     
    strstyle = "δ������ԭ"
    
    strstyle1 = "�޸�ǰ"
    
    strstyle2 = "�޸ĺ�"
     
    If Combo1.Text = "" Then
        MsgBox "��ѡ���޸ķ�ʽ", vbInformation, "��ʾ"
        Exit Sub
            
    End If
            
    If Combo1.Text <> "�ɹ�����" And Combo1.Text <> "�빺����" Then
            
        MsgBox "��ѡ����ȷ���޸ķ�ʽ", vbInformation, "��ʾ"
            
        Exit Sub
            
    End If
    
    If Text1.Text = "" Then
        MsgBox "�����뵥��", vbInformation, "��ʾ"
        Exit Sub

    End If
       
    strBuy = Trim$(Text1.Text)
    
    If lb2 = "�ɹ�����" Then
    
        Call DataDetect2(strBuy)
    
    Else

        Call DataDetect1(strBuy)

    End If

    If Text2.Text = "" Then
        MsgBox "�������Ϻ�", vbInformation, "��ʾ"
        Exit Sub

    End If
    
    strPartno = Trim$(Text2.Text)
    
    Call DataDetect(strPartno)
    
    If lb2 = "�빺����" Then
        
        '��ȡ�ɹ�����
        strBuy1 = Get_SqlStr("select distinct c.�ɹ������ from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.���ϱ�� = d.���ϱ�� WHERE c.�빺����� = '" & strBuy & "' and d.�Ϻ� = '" & strPartno & "'")
   
        If Get_SqlserverCnt("select a.�빺����� from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� WHERE a.�빺����� = '" & strBuy & "' and  a.�Ƿ���� = '1' and b.�Ϻ� = '" & strPartno & "'") = 0 Then
            MsgBox "�˱��빺��δ�������軹ԭ,��ȷ��", vbInformation, "��ʾ"
            Exit Sub
        
        End If
    
        If Get_SqlserverCnt("select a.�ɹ������ from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ��  WHERE a.�빺����� = '" & strBuy & "' and  a.�Ƿ���� = '1' and b.�Ϻ� = '" & strPartno & "' ") = 0 Then
            MsgBox "�˱ʲɹ���δ�������軹ԭ,��ȷ��", vbInformation, "��ʾ"
            Exit Sub
    
        End If
        
        Call Databackup2(strstyle, strstyle1, strBuy, strPartno, 1)
        
        AddSql2 (" update a set a.�Ƿ���� = '0'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '1' ")
      
        Call Databackup2(strstyle, strstyle2, strBuy, strPartno, 0)
        
        Call Databackup1(strstyle, strstyle1, strBuy1, strPartno, 1)
      
        AddSql2 (" update a set a.�Ƿ���� = '0'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '1' ")
        
        Call Databackup1(strstyle, strstyle2, strBuy1, strPartno, 0)
    Else
    
        '��ȡ�빺����
        strBuy1 = Get_SqlStr("select distinct c.�빺����� from erpbase..tblCPurDataSub c inner join erpdata..tblSmainM2 d on c.���ϱ�� = d.���ϱ�� WHERE c.�ɹ������ = '" & strBuy & "' and d.�Ϻ� = '" & strPartno & "'")

        If Get_SqlserverCnt("select a.�빺����� from erpbase..tblCRequest a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� WHERE a.�빺����� = '" & strBuy1 & "' and  a.�Ƿ���� = '1' and b.�Ϻ� = '" & strPartno & "'") = 0 Then
            MsgBox "�˱��빺��δ�������軹ԭ,��ȷ��", vbInformation, "��ʾ"
            Exit Sub
        
        End If
    
        If Get_SqlserverCnt("select a.�ɹ������ from erpbase..tblCPurDataSub a inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ��  WHERE a.�빺����� = '" & strBuy1 & "' and  a.�Ƿ���� = '1' and b.�Ϻ� = '" & strPartno & "' ") = 0 Then
            MsgBox "�˱ʲɹ���δ�������軹ԭ,��ȷ��", vbInformation, "��ʾ"
            Exit Sub
    
        End If

        Call Databackup2(strstyle, strstyle1, strBuy1, strPartno, 1)
         
        AddSql2 (" update a set a.�Ƿ���� = '0'  from erpbase..tblCRequest a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy1 & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '1' ")
        
        Call Databackup2(strstyle, strstyle2, strBuy1, strPartno, 0)
        
        Call Databackup1(strstyle, strstyle1, strBuy, strPartno, 1)
      
        AddSql2 (" update a set a.�Ƿ���� = '0'  from erpbase..tblCPurDataSub a  inner join  erpdata..tblSmainM2 b on a.���ϱ�� = b.���ϱ�� where a.�빺����� = '" & strBuy1 & "' and b.�Ϻ� = '" & strPartno & "' and a.�Ƿ���� = '1' ")

        Call Databackup1(strstyle, strstyle2, strBuy, strPartno, 0)
        
    End If
    
    MsgBox "��ԭ�ɹ�", vbInformation, "��ʾ"

End Sub
