VERSION 5.00
Begin VB.Form Form_MLS 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "条码核对综合系统"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15735
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
   ScaleHeight     =   9045
   ScaleWidth      =   15735
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   13455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   22335
      Begin VB.CommandButton cmd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "核对开始"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   4215
      End
      Begin VB.ComboBox cmbCombo1 
         Height          =   315
         ItemData        =   "From_MLS.frx":0000
         Left            =   1560
         List            =   "From_MLS.frx":0016
         TabIndex        =   1
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "核对模板"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   2
         Top             =   637
         Width           =   960
      End
   End
End
Attribute VB_Name = "Form_MLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SetParent _
                Lib "user32.dll" (ByVal hWndChild As Long, _
                                  ByVal hWndNewParent As Long) As Long
                                  
Option Explicit

Private Sub cmd_Click()

If cmbCombo1.Text = "" Then
    MsgBox "请选择核对模板", vbInformation, "提示"
    Exit Sub
End If

Unload Form_MLS_00

Select Case cmbCombo1.ListIndex

    Case 0
        Frm_MatchLabel.Show     ' 内外箱核对
                
    Case 1
        Form_MLS_00.Show        ' 聚成条码
        
    Case 2
        Form_BarCodeCheckSys.Show   '三码核对
    
    Case 3
        Form_MLS_00.Show   ' GD108
        Form_MLS_00.Caption = "GD108"
    Case 4
        Frm_Bumping_LOT_PACKCODE_COMP.Show
        Frm_Bumping_LOT_PACKCODE_COMP.Caption = cmbCombo1.Text
    Case 5
        FrmBarCodeCheck_JX002.Show 1
    
End Select

End Sub

Private Sub Form_Load()
cmbCombo1.ListIndex = 0
End Sub
