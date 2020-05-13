VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmUploadFile 
   Caption         =   "上传文件"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13350
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
   ScaleHeight     =   8685
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdUpload 
      Caption         =   "上传"
      Height          =   840
      Left            =   6720
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   525
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton cmd 
      Caption         =   ".."
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   720
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   8040
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmUploadFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
On Error Resume Next
Dim FName
    '帅选文件
    CommonDialog2.Filter = "EXCEL文件(*.xlsx)|*.xlsx"
    
    CommonDialog2.ShowOpen
    '得到文件名
    FName = CommonDialog2.filename
    If FName <> "" Then
       Text3.Text = FName
    End If


End Sub

Private Sub cmdUpload_Click()
'上传资料


End Sub
