VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_ProductionPlanNew 
   Caption         =   "��������"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7545
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
   MinButton       =   0   'False
   ScaleHeight     =   9452.914
   ScaleMode       =   0  'User
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlanNew.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlanNew.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlanNew.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlanNew.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlanNew.frx":2368
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_ProductionPlanNew.frx":2C42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   1535
      ButtonWidth     =   2408
      ButtonHeight    =   1482
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "       ����        "
            Key             =   "SEARCH"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Ԥ��"
            Key             =   "PREVIEW"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ˢ��"
            Key             =   "INIT"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�"
            Key             =   "EXIT"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "LOT��ϸ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6975
      Left            =   4680
      TabIndex        =   3
      Top             =   960
      Width           =   2775
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����LOTID"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtSel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "ȫѡ/��ѡ"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5460
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   1440
         Width           =   2295
      End
   End
   Begin VB.Frame Fra 
      Caption         =   "����ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   4575
      Begin VB.ComboBox cb37Pri2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Frm_ProductionPlanNew.frx":351C
         Left            =   2640
         List            =   "Frm_ProductionPlanNew.frx":3526
         TabIndex        =   29
         Text            =   "N"
         Top             =   4965
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   26
         Top             =   3285
         Width           =   1695
      End
      Begin VB.ComboBox cbWO 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   25
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox cbHTPN 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2400
         Width           =   2535
      End
      Begin VB.ComboBox cbCusCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   21
         Top             =   1185
         Width           =   1455
      End
      Begin VB.TextBox txtCusPN 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   1598
         Width           =   3135
      End
      Begin VB.ComboBox cbPN 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1995
         Width           =   2535
      End
      Begin VB.ComboBox cb37Pri 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Frm_ProductionPlanNew.frx":3530
         Left            =   1080
         List            =   "Frm_ProductionPlanNew.frx":353D
         TabIndex        =   17
         Text            =   "Normal Lot"
         Top             =   4965
         Width           =   1455
      End
      Begin VB.ComboBox cbLotType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Frm_ProductionPlanNew.frx":3565
         Left            =   2760
         List            =   "Frm_ProductionPlanNew.frx":3575
         TabIndex        =   16
         Text            =   "������(M)"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtWODept 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3645
         Width           =   3255
      End
      Begin VB.ComboBox cbWOType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Frm_ProductionPlanNew.frx":35A5
         Left            =   1080
         List            =   "Frm_ProductionPlanNew.frx":35BB
         TabIndex        =   14
         Text            =   "��ͨ����"
         Top             =   360
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   20
         Top             =   5865
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   65280
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16777215
         Format          =   291110913
         CurrentDate     =   43271
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   22
         Top             =   6285
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   16744576
         CalendarTitleBackColor=   16744703
         CalendarTitleForeColor=   8438015
         CalendarTrailingForeColor=   16777215
         Format          =   291110913
         CurrentDate     =   43271
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   12
         Top             =   3720
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ǰ׺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   8
         Left            =   240
         TabIndex        =   11
         Top             =   3300
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   10
         Top             =   420
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "37_PRI"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   5025
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڻ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   2445
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�깤����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   6360
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   5955
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʒ�Ϻ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   2055
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   1650
         Width           =   840
      End
   End
End
Attribute VB_Name = "Frm_ProductionPlanNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SetParent _
                Lib "user32.dll" (ByVal hWndChild As Long, _
                                  ByVal hWndNewParent As Long) As Long

Private Sub Form_Load()
    InitDate
    InitCustomerCode
    InitLotWO

End Sub

Private Sub InitDate()
DTPicker1(0).Value = Format(Now(), "yyyy-MM-dd")
DTPicker1(1).Value = Format(Year(Now()) & "-" & Month(Now()) & "-" & "28", "yyyy-MM-dd")
End Sub

Private Sub InitCustomerCode()

    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = SqlConnect
    rs.Source = "SELECT �ͻ����� as PID,�ͻ����� as productname FROM erpdata.dbo.tblXCustomer " & " union  select 'JX117' as PID,'JX117' as productname " & " union  select 'AA(ON)' as PID,'AA(ON)' as productname " & " union  select '37(ICI)' as PID,'37(ICI)' as productname " & " union  select 'AB18(2)' as PID,'AB18(2)' as productname " & " union  select 'BUMPINGDM' as PID,'BUMPINGDM' as productname " & " union select 'YZ22(2)' as PID,'YZ22(2)' as productname order by �ͻ����� "

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    cbCusCode.Clear

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount
            cbCusCode.AddItem Trim(rs("productname"))
            rs.MoveNext
        Next i

    End If
  
    rs.Close
    Set rs = Nothing

End Sub

Private Sub InitLotWO()

    Dim strSql As String

    Dim rs     As New ADODB.Recordset
    
    strSql = "select distinct substr(trim(ordername),1,3) as prefix from ib_wohistory order by prefix "

    If rs.State = adStateOpen Then rs.Close
    rs.Open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
    cbWO.Clear
    cbWO.AddItem ("")
    
    If Not rs.EOF Then
        
        Do While Not rs.EOF
            cbWO.AddItem Trim$("" & rs!Prefix)
            rs.MoveNext
        Loop

    End If

End Sub

Private Sub ClearData()

    cbPN.Clear
    cbHTPN.Text = ""
    cbWO.Text = ""
    txtWODept.Text = ""
    List1.Clear

End Sub

Private Sub cbPN_Click()

    ' �������ڻ���
    Dim rs       As New ADODB.Recordset

    Dim strCusPN As String

    Dim strPN    As String
  
    cbHTPN.Text = ""
      
    strCusPN = Trim(txtCusPN.Text)
    strPN = Trim$(cbPN.Text)
    
    Set rs.ActiveConnection = OraConnect
    rs.Source = "select distinct qtechptno from tbltsvnpiproduct where customerptno1 = '" & strCusPN & "' and qtechptno2 = '" & strPN & "' "

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        If rs.RecordCount > 1 Then
            MsgBox "�ϺŴ����˶�����ڻ���, ��NPIȷ��", vbCritical, "����"

            Exit Sub

        End If

        rs.MoveFirst

        For i = 1 To rs.RecordCount
            cbHTPN.Text = Trim(rs("qtechptno"))
            rs.MoveNext
        Next i

    Else
        MsgBox "�û���:" & strCusPN & "��ѯ�������ڻ���, ��NPIȷ��", vbCritical, "����"

        Exit Sub

    End If

    rs.Close
    Set rs = Nothing
    
    ' 37�жϳ��ڻ��ֺ��ϺŹ�ϵ
   
    If Trim(cbCusCode.Text) = "37" And cbHTPN.Text = "X37B" Then
        If Left(Right(strPN, 2), 1) <> "B" Then
            MsgBox "NPIά������, X37B��Ӧ�Ϻŵ����ڶ�λ������B, ��NPIȷ��", vbCritical, "����"
            
            cbHTPN.Text = ""

            Exit Sub

        End If

    End If
        
    ' ������������
    Dim sProductDept As String
    
    Dim sProductCode As String
    
    txtWODept.Text = ""

    sProductDept = GetWoDept(cbPN.Text)
    sProductCode = GetGWoDeptID(sProductDept)

    If sProductDept <> "" And sProductCode <> "" Then
        txtWODept.Text = sProductDept & "_" & sProductCode

    End If

End Sub

Private Sub cbWO_Change()

    If Mid$(Trim(cbWO.Text), 2, 1) = "P" Or Mid$(Trim$(cbWO.Text), 2, 1) = "T" Then
        cbLotType.Text = "������(M)"

    End If

    If Mid$(Trim(cbWO.Text), 2, 1) = "S" Then
        cbLotType.Text = "������(E)"

    End If

End Sub

Private Sub cbWO_Click()

    If Mid$(Trim(cbWO.Text), 2, 1) = "P" Or Mid$(Trim$(cbWO.Text), 2, 1) = "T" Then
        cbLotType.Text = "������(M)"

    End If

    If Mid$(Trim(cbWO.Text), 2, 1) = "S" Then
        cbLotType.Text = "������(E)"

    End If

End Sub

Private Sub cbWOType_Click()
   
    Select Case cbWOType.Text

        Case "�ع�����"
            Unload Frm_ReWONew
            SetParent Frm_ReWONew.hWnd, Me.hWnd
            Frm_ReWONew.Show
            
        Case "Dummy����"
            Unload Frm_ReWONew
            SetParent Frm_ReWONew.hWnd, Me.hWnd
            Frm_ReWONew.Show
            
        Case "��������"
            Unload Frm_ReWONew
            SetParent Frm_ReWONew.hWnd, Me.hWnd
            Frm_ReWONew.Show
            
        Case "�������"
            Unload Frm_ReWONew
            SetParent Frm_ReWONew.hWnd, Me.hWnd
            Frm_ReWONew.Show

    End Select
    
End Sub

Private Sub Check1_Click()

    Dim i As Integer

    If Check1.Value = 1 Then
    
        With List1

            For i = 0 To .ListCount - 1
                    
                .Selected(i) = True
                    
            Next
                
        End With
        
    ElseIf Check1.Value = 0 Then

        With List1

            For i = 0 To .ListCount - 1
                    
                .Selected(i) = False
                    
            Next
                
        End With
        
    End If

End Sub

Private Sub cmdQuery_Click()

    Dim strKey As String

    strKey = Trim$(txtSel)

    If strKey = "" Then
        MsgBox "������LOT ID", vbInformation, "��ʾ:"

        Exit Sub

    End If

    With List1

        For i = 0 To .ListCount - 1

            If strKey = .List(i) Then
            
                .Selected(i) = True

            End If
        
        Next

    End With

End Sub

Public Sub ForSearch()

    Dim strCusCode As String

    Dim strCusPN   As String

    Dim strdevice  As String
    
    Dim rs8        As New ADODB.Recordset
        
    If cbWOType.Text = "��������" Then
        If InStr(txtCusPN.Text, "-CV") = 0 Then
            txtCusPN.Text = txtCusPN.Text & "-CV"

        End If

    End If
    
    If cbWOType.Text = "�������" Then
        If InStr(txtCusPN.Text, "-FO") = 0 Then
            MsgBox "��������Ŀͻ����ֺ�׺����Ϊ'-FO'", vbCritical, "����"
            Exit Sub

        End If

    End If
    
    If cbWOType.Text = "FO_CSP����" Then
        If InStr(txtCusPN.Text, "-FO") > 0 Then
            MsgBox "FO_CSP�����Ŀͻ����ֺ�׺�����԰���'-FO'", vbCritical, "����"
            Exit Sub

        End If

    End If
   
    strCusCode = cbCusCode.Text
    strCusPN = Trim$(txtCusPN.Text)

    If strCusCode = "" Then
        MsgBox "�ͻ����벻��Ϊ��", vbCritical, "����"

        Exit Sub

    End If
    
    If strCusPN = "" Then
        MsgBox "�ͻ����ֲ���Ϊ��", vbCritical, "����"

        Exit Sub
        
    End If
    
    strdevice = "select * from tbltsvnpiproduct a ,ib_wohistory b where a.customerptno1 = '" & strCusPN & "' and a.customershortname = '" & strCusCode & "' and b.product = a.qtechptno2 and TO_CHAR(B.ERPCREATEDATE,'YYYY-MM-DD') > to_char( sysdate -180,'YYYY-MM-DD')  "
    
    If rs8.State = adStateOpen Then rs8.Close
    rs8.Open strdevice, Cnn, adOpenStatic, adLockReadOnly, adCmdText
     
    If rs8.RecordCount < 1 Then
        MsgBox "�ÿͻ�����:" & strCusCode & "����:" & strCusPN & ":������û�������� ", vbCritical, "����"
        MsgBox "�ÿͻ�����:" & strCusCode & "����:" & strCusPN & ":������û�������� ", vbCritical, "����"
    
    End If

    rs8.Close
    
    Call SearchByCPN(strCusCode, strCusPN)
  
End Sub

Private Sub SearchByCPN(strCusCode As String, strCusPN As String)

    Dim rs  As New ADODB.Recordset

    Dim rs1 As New ADODB.Recordset

    Dim Rs2 As New ADODB.Recordset
    
    ' ����У��
    If (cbCusCode.Text = "37" Or cbCusCode.Text = "EU010" Or cbCusCode.Text = "HK075") And cbWOType.Text <> "�ع�����" Then
        Set rs1.ActiveConnection = OraConnect

        rs1.Source = " select * from tbltsvnpiproduct a where a.customershortname in ( '37','EU010','HK075')  and   instr(a.struckstr1,'ASSY') >0  and a.customerptno1 = '" & Trim$(txtCusPN.Text) & "'"

        rs1.Open , , adOpenStatic, adLockReadOnly, adCmdText

        If rs1.RecordCount > 0 Then
            rs1.Close
            Set rs1 = Nothing
            Set Rs2.ActiveConnection = OraConnect
            Rs2.Source = "   select * from code37 d where d.device = '" & Trim$(txtCusPN.Text) & "' "
            Rs2.Open , , adOpenStatic, adLockReadOnly, adCmdText

            If Rs2.RecordCount < 1 Then

                MsgBox "�ÿͻ�����:" & strCusCode & "����:" & strCusPN & ": û��ά��������", vbCritical, "����"

                Exit Sub

            End If

        End If

    End If

    ' �����Ϻ�
    Set rs.ActiveConnection = OraConnect
    
    If cbWOType.Text = "Dummy����" Then
    
        rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & strCusCode & "' and  customerptno1 = '" & strCusPN & "' and substr(qtechptno2, -3, 1) = 'W' "
    
    Else
    
        rs.Source = "select distinct qtechptno2 from tbltsvnpiproduct where customershortname = '" & strCusCode & "' and customerptno1 = '" & strCusPN & "' and substr(qtechptno2, -3, 1) <> 'W' "

    End If

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText
  
    cbPN.Clear

    If rs.RecordCount > 0 Then
        If rs.RecordCount > 1 Then
            MsgBox "��ע��,�ÿͻ����ְ��������Ʒ�Ϻ�, ��ȷ���Ƿ�����", vbInformation, "��ʾ"

        End If
    
        rs.MoveFirst

        For i = 1 To rs.RecordCount
            cbPN.AddItem Trim(rs("qtechptno2"))
            cbPN.Text = Trim(rs("qtechptno2"))
            rs.MoveNext
        Next i

    Else
        MsgBox "�ÿͻ�����:" & strCusCode & "����:" & strCusPN & ": NPIδά����Ӧ��ϵ, ��ѯ�����Ϻ�", vbCritical, "����"

        Exit Sub

    End If
  
    rs.Close
    Set rs = Nothing

    ' ��ѯ�˻���,����LotID
    If strCusCode = "AA" And cbWOType.Text <> "Dummy����" And cbWOType.Text <> "��������" Then
        Call GetAALotID(rs, strCusCode, strCusPN)
    Else
        Call GetLotID(rs, strCusCode, strCusPN)

    End If
  
    List1.Clear

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount
            List1.AddItem Trim(rs("source_batch_id"))
            rs.MoveNext
        Next i

    Else
        MsgBox "�û���:" & strCusPN & "��ѯ����������Ϣ, ��ȷ��" & vbCrLf & "���,����,dummy�������ֶ�ά������", vbCritical, "����"

        Exit Sub

    End If
  
    rs.Close
    Set rs = Nothing
    
    tb1.Buttons("SEARCH").Enabled = False
    tb1.Buttons("PREVIEW").Enabled = True

End Sub

Private Sub GetLotID(ByRef rs As ADODB.Recordset, _
                     strCusCode As String, _
                     strCusPN As String)
 
    Set rs.ActiveConnection = OraConnect
    
    If cbWOType.Text = "�ع�����" Or cbWOType.Text = "ί�⹤��" Then
        rs.Source = "select distinct a.source_batch_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0   and instr(b.substrateid, '+') > 0 and  not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    
    ElseIf cbWOType.Text = "�������" Then
        rs.Source = "select distinct a.source_batch_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and a.flag = 'T'  and instr(b.substrateid,'+') = 0 and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and a.source_batch_id like 'SI%' and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"

    ElseIf cbWOType.Text = "��������" Then
        rs.Source = "select distinct a.source_batch_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and a.flag = 'T' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and a.source_batch_id like 'G%' and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    
    ElseIf cbWOType.Text = "Dummy����" Then
        rs.Source = "select distinct a.source_batch_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and a.flag = 'T' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and (a.source_batch_id like 'D%' or a.source_batch_id like 'SI%') and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    
    ElseIf cbWOType.Text = "FO_CSP����" Then
        rs.Source = "select distinct a.source_batch_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and a.flag = 'T' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and a.source_batch_id like 'SI%' and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"
    
    Else
        rs.Source = "select distinct a.source_batch_id from customeroitbl_test a,mappingdatatest b where a.customershortname = '" & strCusCode & "' and  a.mpn_desc = '" & strCusPN & "' and a.flag = 'Y' and to_char(a.id) = b.filename and a.source_batch_id=b.lotid and a.invflag = 0  and not exists (select 1 from ib_waferlist c where b.substrateid =c.waferid) order by a.source_batch_id"

    End If

    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

End Sub

Private Sub GetAALotID(ByRef rs As ADODB.Recordset, _
                       strCusCode As String, _
                       strCusPN As String)
    
    Dim customerPTTemp As String

    Dim opnTemp        As String
    
    opnTemp = strCusPN
    
    customerPTTemp = GetONOPN_WSG(opnTemp)

    Set rs.ActiveConnection = OraConnect

    rs.Source = " select distinct b.id, b.batchid as source_batch_id from ( select * from (select *  from CUSTOMERFORECASTTBL   order by ID desc) where   out_part_id = '" & customerPTTemp & "'  and rownum = 1 ) a ,CustomerBCtbl b " & "  where a.out_part_id='" & customerPTTemp & "' and a.comments='" & opnTemp & "' and a.flag='Y' and a.start_part_id=b.mtrlnum and b.batchid not in (select lotid from  On_WO_HisTory where flag='Y')   order by b.id "
       
    rs.Open , , adOpenStatic, adLockReadOnly, adCmdText

End Sub

Private Sub ForInit()

    tb1.Buttons("SEARCH").Enabled = True
    tb1.Buttons("PREVIEW").Enabled = False
    ClearData

End Sub

Private Sub ForExit()

    Unload Me
    Unload Frm_ProductionPlanDetailNew

End Sub

Private Sub ForPreview()
    Screen.MousePointer = 11

    Unload Frm_ProductionPlanDetailNew

    If CheckPowerInfo = True Then

        If List1.SelCount > 0 Then
            SetParent Frm_ProductionPlanDetailNew.hWnd, Me.hWnd

            Frm_ProductionPlanDetailNew.Show
        Else
            MsgBox "��ѡ��LOT", vbCritical, "����"
            Screen.MousePointer = 0
            Exit Sub

        End If

    End If

    Screen.MousePointer = 0

End Sub



Private Sub tb1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key

        Case "SEARCH"
            ForSearch
    
        Case "INIT"
            ForInit
        
        Case "EXIT"
            ForExit
        
        Case "PREVIEW"
            ForPreview

    End Select

End Sub

Public Function GetWOID() As String

    Dim FirstChar    As String

    Dim SeqChar      As String

    Dim typenameTemp As String

    Dim yMonthTemp   As String

    Dim seqTemp      As Integer

    Dim headChar     As String

    Dim mdChar       As String

    Dim ID           As Long
    
    Dim strWOID      As String
    
    FirstChar = UCase(Trim(cbWO.Text))
 
    If Len(FirstChar) <> 3 Then
        MsgBox "�����빤��ǰ��λ!"
        cbWO.Text = ""

        Exit Function

    End If

    headChar = FirstChar

    SeqChar = GetWoIDTemp(FirstChar)
    mdChar = Right(Year(DateTime.DATE), 2) & Right("0" & Month(DateTime.DATE), 2)
    FirstChar = FirstChar & "-" & mdChar

    SeqChar = Right("000" & CStr(CInt(SeqChar)), 4)
    
    ID = CLng(SeqChar)
    
    strWOID = FirstChar & SeqChar
    
    cmdStr = "insert into TSV_WO_SEQ_TAB(wotype,ymonth,sequenceid,flag) values ( '" & headChar & "','" & mdChar & "'," & ID & ", 'Y' ) "
    AddSql (cmdStr)
    
    GetWOID = strWOID

End Function

Private Function CheckPowerInfo() As Boolean

    CheckPowerInfo = False

    If cbWO.Text = "" Then
        MsgBox "����ǰ׺������Ϊ��", vbCritical, "����"

        Exit Function

    Else

        If Len(Trim$(cbWO.Text)) <> 3 Then
            MsgBox "����ǰ׺������3λ", vbCritical, "����"

            Exit Function

        End If
    
    End If
    
    If txtWODept.Text = "" Or txtWODept.Text = "_" Then
      
        MsgBox "�������Ų�����Ϊ��", vbCritical, "����"

        Exit Function
        
    End If
    
    If cbCusCode.Text = "" Then

        MsgBox "�ͻ����벻����Ϊ��", vbCritical, "����"

        Exit Function

    End If
        
    If txtCusPN.Text = "" Then
        
        MsgBox "�ͻ����ֲ�����Ϊ��", vbCritical, "����"

        Exit Function

    End If
    
    If cbPN.Text = "" Then

        MsgBox "��Ʒ�ϺŲ�����Ϊ��", vbCritical, "����"

        Exit Function

    Else

        If cbWOType.Text <> "Dummy����" And cbWOType.Text <> "��������" And cbWOType.Text <> "FO_CSP����" And cbWOType.Text <> "�������" And cbWOType.Text <> "�ع�����" Then
            If CheckPN(Trim$(cbPN.Text), Trim(txtWODept.Text)) = False Then

                Exit Function

            End If

        End If

    End If

    If cbHTPN.Text = "" Then
       
        MsgBox "���ڻ��ֲ�����Ϊ��", vbCritical, "����"

        Exit Function

    End If
    
    If cb37Pri.Text = "" Then
       
        MsgBox "37PRI������Ϊ��", vbCritical, "����"

        Exit Function

    End If
    
    If cbLotType.Text = "" Then
        
        MsgBox "�������Ͳ�����Ϊ��", vbCritical, "����"

        Exit Function

    End If
    
    If DTPicker1(0).Value > DTPicker1(1).Value Then
        MsgBox "�������ڱ��������깤����", vbCritical, "����"
        
        Exit Function

    ElseIf DTPicker1(0).Value = DTPicker1(1).Value Then
        MsgBox "�������ڲ����Ե����깤����", vbCritical, "����"
        
        Exit Function

    End If
    
    If cbWOType.Text = "��������" Then
        If CheckBLWO(Trim(cbCusCode.Text), Trim(txtCusPN.Text), Trim(cbHTPN.Text)) = False Then
            MsgBox "��������û��ά���ض�����Ϣ(��ϴ����, CV�߶�, ��ϴ����, �������), ����ϵNPIά����Ӧ���ֵ���Ϣ", vbCritical, "��ʾ"
            Exit Function

        End If
        
    End If
    
    CheckPowerInfo = True

End Function

Private Function CheckBLWO(strCusCode, strCusPN, strHTPN) As Boolean

    Dim strSql As String

    CheckBLWO = False

    strSql = "select * from tbltsvnpiproduct where customershortname = '" & strCusCode & "' and customerptno1 = '" & strCusPN & "' and qtechptno = '" & strHTPN & "' and  customerptno3 is not null and customerptno4 is not null and customerptno5 is not null and customerptno6 is not null"

    If Get_OracleCnt(strSql) = 0 Then
        Exit Function

    End If

    CheckBLWO = True

End Function

Private Function CheckPN(strPN As String, strDept As String) As Boolean
    CheckPN = False

    Dim bomRS2 As New ADODB.Recordset

    Set bomRS2 = GetProductBom(strPN)

    If bomRS2.RecordCount <= 0 Then
        MsgBox "��ϵͳ�����Ϻŵ�Bom�����ڣ�����ϵ��ص��ˣ���ά��Bom ��"

        Exit Function

    End If

    Set bomRS2 = GetProductJDObject(strPN)

    If bomRS2.RecordCount <= 0 Then
        MsgBox "���Ϻ��ڽ��ϵͳ���޳ɱ��������������Աȷ�� ��"
    
        Exit Function

    End If

    '
    '    If InStr(UCase(strDept), "BUMPING") = 0 And InStr(UCase$(strDept), "SSP") = 0 And InStr(UCase(strDept), "WLP") = 0 Then
    '        Set bomRS2 = GetProduct_Check(strPN)
    '
    '        If bomRS2.RecordCount <= 0 Then
    '            MsgBox "�ϺŲ����ڣ�����ϵ��ص��ˣ���ά���Ϻ� ��"
    '
    '            Exit Function
    '
    '        End If
    '
    '    End If

    Set bomRS2 = GetProductBomERpSign(strPN)

    If bomRS2.RecordCount <= 0 Then
        MsgBox "��ϵͳ�����Ϻŵ�Bomû�б����ͨ��������ϵ���̲���"

        Exit Function

    End If

    CheckPN = True

End Function

