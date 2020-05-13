VERSION 5.00
Begin VB.Form DlgEventReason 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "删除/修改数据事由"
   ClientHeight    =   3645
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frame 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton btnCommit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "提交"
         Height          =   480
         Left            =   240
         TabIndex        =   2
         Top             =   2880
         Width           =   990
      End
      Begin VB.TextBox txtReason 
         BackColor       =   &H00E0E0E0&
         Height          =   2295
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   360
         Width           =   6015
      End
   End
End
Attribute VB_Name = "DlgEventReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()

End Sub

Private Sub btnCommit_Click()
If txtReason.Text = "" Then
    MsgBox "请填写删除/修改数据的理由", vbInformation, "提示"
    Exit Sub
End If

strGReason = Trim$(txtReason.Text)
Unload Me

End Sub

