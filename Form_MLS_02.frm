VERSION 5.00
Begin VB.Form Form_MLS_02 
   Caption         =   "艾为客户核对"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ClipControls    =   0   'False
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
   ScaleHeight     =   15690
   ScaleWidth      =   28680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   10455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20055
      Begin VB.TextBox txtBoxQty 
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox txtBoxID 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   3960
         Width           =   2535
      End
      Begin VB.TextBox txtCartonQty 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtCartonID 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   2040
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form_MLS_02.frx":0000
         Left            =   1440
         List            =   "Form_MLS_02.frx":000A
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtScan 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内箱数量:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   4440
         Width           =   1065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内箱ID:"
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
         Index           =   3
         Left            =   480
         TabIndex        =   8
         Top             =   3960
         Width           =   825
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱数量:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "外箱ID:"
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
         Index           =   1
         Left            =   600
         TabIndex        =   4
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扫描框:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form_MLS_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
txtScan.SetFocus
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
End Sub

Private Sub txtScan_KeyPress(KeyAscii As Integer)

    If KeyAscii <> vbKeyReturn Then
        Exit Sub

    End If
    
    If txtScan.Text = "" Then
        Exit Sub

    End If
    
    Select Case Combo1.ListIndex

        Case 0  ' 外箱
            Call DoCarton(UCase(Trim$(txtScan.Text)))

        Case 1  ' 内箱
            Call DoBox(UCase(Trim$(txtScan.Text)))

    End Select
    
    txtScan.Text = ""
    
End Sub

Private Sub DoCarton(strCode As String)
Dim sPart

sPart = Split(strCode, ";")



End Sub

Private Sub DoBox(strCode As String)



End Sub
