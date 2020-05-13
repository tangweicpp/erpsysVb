VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form4Temp 
   Caption         =   "WIP报表中QtechRequestDate 设定"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
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
   ScaleHeight     =   6120
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "导出报表"
      Height          =   600
      Left            =   720
      TabIndex        =   12
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "修改"
      Height          =   1935
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   9615
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   1200
         Width           =   5415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改"
         Height          =   360
         Left            =   6960
         TabIndex        =   9
         Top             =   480
         Width           =   990
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   4680
         TabIndex        =   10
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   174260225
         CurrentDate     =   40947
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "W/O issue day Remark："
         Height          =   195
         Left            =   600
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新日期："
         Height          =   195
         Left            =   3960
         TabIndex        =   11
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOTID："
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "新增"
      Height          =   1815
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   9615
      Begin VB.TextBox TxtRemark 
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   1080
         Width           =   5415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "添加"
         Height          =   360
         Left            =   6960
         TabIndex        =   6
         Top             =   360
         Width           =   990
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   174260225
         CurrentDate     =   40947
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "W/O issue day Remark："
         Height          =   195
         Left            =   600
         TabIndex        =   13
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期："
         Height          =   195
         Left            =   4080
         TabIndex        =   5
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOTID："
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   630
      End
   End
End
Attribute VB_Name = "Form4Temp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oiRS        As New ADODB.Recordset

Private Sub Command1_Click()
'增加
Dim lotIDTemp As String
Dim dtTemp As Date
Dim sqlTemp As String
Dim remarkTemp As String



If Trim(Text1.Text) <> "" Then

    lotIDTemp = Trim(Text1.Text)
    dtTemp = DTPicker1.Value
    remarkTemp = Trim(TxtRemark.Text)
    
    '判断输入的Lot号，是否正确
    
    If JudgeLot2(lotIDTemp) Then
    
    
        '判断是否存在 存在则提示信息
        If Not (JudgeLot(lotIDTemp)) Then
        sqlTemp = "insert into WipreportDate(lotid,lotdate,remark) values ( '" & lotIDTemp & "',to_date('" & dtTemp & "','yyyy-mm-dd'),'" & remarkTemp & "' ) "
        AddSql (sqlTemp)
        MsgBox "添加成功!"
        
        Else
        
        MsgBox "LotId:" & lotIDTemp & "已存在！"
        End If
        
    Else
         MsgBox "LotId:" & lotIDTemp & "在Mes系统中不存在，请确认Lot号！"
    
    End If
    

Else
MsgBox "请先输入LotId!"
End If


End Sub

Private Sub Command2_Click()
'修改
Dim lotIDTemp As String
Dim dtTemp As Date
Dim sqlTemp As String
Dim remarkTemp As String


If Trim(Text2.Text) <> "" Then

    lotIDTemp = Trim(Text2.Text)
    dtTemp = DTPicker3.Value
    remarkTemp = Trim(Text3.Text)
    
    '判断是否存在 存在则修改，不存在提示
     If JudgeLot(lotIDTemp) Then
     
        sqlTemp = "update WipreportDate set lotdate=to_date('" & dtTemp & "','yyyy-mm-dd'), remark='" & remarkTemp & "'    where lotid='" & lotIDTemp & "' "
        AddSql (sqlTemp)
        MsgBox "修改成功!"
        
    Else
        
          MsgBox "LotId:" & lotIDTemp & "不存在！"
     End If
    

    

Else
MsgBox "请先输入LotId!"
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

Public Function JudgeLot2(lotIDTemp As String) As Boolean

Dim cmdStr As String
Dim slectResult As Boolean
slectResult = False
cmdStr = "  select * from A_Lotwafers where  wafernumber='" + lotIDTemp + "' "
         
         
slectResult = QueryStr(cmdStr)
JudgeLot2 = slectResult
End Function



Private Sub Command3_Click()
 ExporToExcel ("select lotid,remark,CreateDate from WipreportDateRemark order by CreateDate desc ")
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
