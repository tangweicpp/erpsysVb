VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Frm_AlterQbox 
   Caption         =   "����쳣����"
   ClientHeight    =   6765
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20250
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
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   20295
      Begin FPSpreadADO.fpSpread FPS 
         Height          =   5415
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   20055
         _Version        =   524288
         _ExtentX        =   35375
         _ExtentY        =   9551
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
         SpreadDesigner  =   "Frm_AlterQbox.frx":0000
      End
      Begin VB.Frame Frame3 
         Caption         =   " "
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   5760
         Width           =   4935
         Begin VB.CommandButton Command1 
            Caption         =   "��ѯ"
            Height          =   360
            Left            =   3720
            TabIndex        =   8
            Top             =   240
            Width           =   990
         End
         Begin VB.TextBox txtText1 
            Height          =   285
            Left            =   1560
            TabIndex        =   5
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " "
         Height          =   735
         Left            =   5280
         TabIndex        =   1
         Top             =   5760
         Width           =   7335
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   1680
            TabIndex        =   10
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton cmd 
            Caption         =   "����"
            Height          =   360
            Left            =   5520
            TabIndex        =   7
            Top             =   240
            Width           =   990
         End
         Begin VB.CommandButton qboxSelect 
            Caption         =   "��ѯ"
            Height          =   360
            Left            =   3960
            TabIndex        =   2
            Top             =   240
            Width           =   990
         End
         Begin VB.Label lblWAFER 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����WAFER��"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "Frm_AlterQbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cmd_Click()
Dim cmd As New ADODB.Command
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim seqV As Integer
Dim strSql As String

strSql = " SELECT MAX(SEQ) FROM TSV_QBOXNUMBER_DETAILS"

RS.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

seqV = RS.fields(0).Value '��õ�ǰseq

strSql = "INSERT INTO TSV_QBOXNUMBER_DETAILS (SEQ,NDPW,WAFERNUMBER,WAFERSCRIBENUMBER,WORKORDERNAME,QBOXNUMBER,CONTAINERNAME,CUSTOMERNAME,PRODUCTNAME,SPECNAME) VALUES " & _
"('" & seqV + 1 & "','0','LOT��','WAFER��','������','���','������','�ͻ�����','�Ϻ�','վ��')" '����һ�ʵ�ǰseq
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
   
   
strSql = "SELECT * FROM TSV_QBOXNUMBER_DETAILS WHERE SEQ='" & seqV + 1 & "'"
RS1.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText
Set FPS.DataSource = RS1 '��ǰseq��ʾ��FPS�ؼ��ϸ�������Ա����
FPS.MaxRows = RS.RecordCount

End Sub

Private Sub Command1_Click()
Text1.Text = ""
Dim strSql              As String
Dim RS                  As New ADODB.Recordset

strSql = "select * from TSV_QBOXNUMBER_DETAILS where QBOXNUMBER='" & Trim(txtText1.Text) & "'"
If Cnn.State = 0 Then
  ConOracle
End If
    
RS.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

If RS.RecordCount > 0 Then
 Set FPS.DataSource = RS
 FPS.MaxRows = RS.RecordCount

 Else
  MsgBox "��ѯ�����κ���Ϣ"
   FPS.MaxRows = 0
End If

End Sub

Private Sub FPS_EditChange(ByVal Col As Long, ByVal Row As Long)

'��ȡSPC��������
Dim strRow As Integer
Dim strCol As Integer
Dim seq As String
Dim collValue As String

FPS.Row = FPS.ActiveRow
FPS.Col = 32
seq = FPS.Text '��õ�Ԫ��ΨһֵSEQ
FPS.Col = FPS.ActiveCol
collValue = FPS.Text '��̬��ȡ���ж�λ��ֵ


Dim cmd As New ADODB.Command
 '����
 If FPS.Col = 2 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set NDPW='" & collValue & "'  where seq='" & seq & "'"
    ElseIf FPS.Col = 3 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set WAFERNUMBER='" & collValue & "'  where seq='" & seq & "'"
   If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 4 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set WAFERSCRIBENUMBER='" & collValue & "'  where seq='" & seq & "'"
   If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 5 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set WORKORDERNAME='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 6 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set FIRSTNAME='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 7 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set QBOXNUMBER='" & collValue & "'  where seq='" & seq & "'"
   If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 8 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set CONTAINERNAME='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 9 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set WORKORDERATTR1='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 10 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set FABFACILITY='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 11 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set IMAGERREV='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 12 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set DESIGNID='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 13 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set QTY1='" & collValue & "'  where seq='" & seq & "'"
   If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 14 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set QTY2='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 15 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set WORKORDERATTR2='" & collValue & "'  where seq='" & seq & "'"
   If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 16 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set QBOX2='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 17 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set WORKORDERATTR3='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 18 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set DATA_CODE1='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 19 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set DATA_CODE2='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 20 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set IPT='" & collValue & "'  where seq='" & seq & "'"
   If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 21 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set ELOT='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 22 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set MPN='" & collValue & "'  where seq='" & seq & "'"
   If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 26 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set CUSTOMERNAME='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 27 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set PDATA1='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 28 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set PRODUCTNAME='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 29 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set TST_PROGRAM_REV='" & collValue & "'  where seq='" & seq & "'"
    If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
    
    ElseIf FPS.Col = 30 Then
    strSql = "update  TSV_QBOXNUMBER_DETAILS set SPECNAME='" & collValue & "'  where seq='" & seq & "'"
   If Cnn.State = 0 Then
   ConOracle
   End If
   cmd.ActiveConnection = Cnn
   cmd.CommandText = strSql
   cmd.CommandType = adCmdText
   cmd.Execute
     
    
 End If
 
 
 
  

End Sub

Private Sub qboxSelect_Click()
Dim strSql              As String
Dim RS                  As New ADODB.Recordset

txtText1.Text = ""
strSql = "select * from TSV_QBOXNUMBER_DETAILS where WAFERSCRIBENUMBER='" & Trim(Text1.Text) & "'"
If Cnn.State = 0 Then
  ConOracle
End If
    
RS.open strSql, Cnn, adOpenStatic, adLockReadOnly, adCmdText

If RS.RecordCount > 0 Then
 Set FPS.DataSource = RS
 FPS.MaxRows = RS.RecordCount

 Else
  MsgBox "��ѯ�����κ���Ϣ"
   FPS.MaxRows = 0
End If
End Sub
