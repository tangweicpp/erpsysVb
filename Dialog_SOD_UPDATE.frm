VERSION 5.00
Begin VB.Form Dialog_SOD_UPDATE 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改SOD的原因"
   ClientHeight    =   2115
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRemark 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "不填"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "填入原因"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog_SOD_UPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub OKButton_Click()
frmShippingScheduleSystem.txtAdd.Text = Trim(txtRemark.Text)
Unload Me
End Sub
