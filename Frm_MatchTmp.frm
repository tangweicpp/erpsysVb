VERSION 5.00
Begin VB.Form Frm_MatchTmp 
   Caption         =   "临时标签条码比对(一对一)"
   ClientHeight    =   11310
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15285
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
   ScaleHeight     =   11310
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   11415
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   15255
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   9015
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   1560
         Width           =   4935
      End
      Begin VB.CommandButton cmdExportHistory 
         BackColor       =   &H00C0C0C0&
         Caption         =   "导出比对记录"
         Height          =   435
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00808080&
         Caption         =   "退出"
         Height          =   435
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdCheckRes 
         Enabled         =   0   'False
         Height          =   9000
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1560
         Width           =   8895
      End
      Begin VB.TextBox txtItem2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1560
         TabIndex        =   4
         Top             =   930
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox txtItem1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item 2:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item 1:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1170
      End
   End
End
Attribute VB_Name = "Frm_MatchTmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdExportHistory_Click()
Dim strSql As String
strSql = "select * from UNIQUE_TBL_NEW where KEYNAME ='临时一对一' and KEYFROM = '37' order by KEYTIME desc "
ExporToExcel (strSql)
End Sub

Private Sub Form_Activate()
txtItem1.Visible = True
txtItem1.SetFocus
End Sub

Private Sub txtItem1_KeyPress(KeyAscii As Integer)
Dim strItem1 As String
If KeyAscii <> vbKeyReturn Or Len(Trim(txtItem1.Text)) = 0 Then Exit Sub
txtItem2.Visible = True
txtItem2.SetFocus
End Sub

Private Sub txtItem2_KeyPress(KeyAscii As Integer)
Dim strItem1 As String
Dim strItem2 As String
Dim strSql As String

If KeyAscii <> vbKeyReturn Or Len(Trim(txtItem2.Text)) = 0 Then Exit Sub

strItem1 = UCase(Trim(txtItem1.Text))
strItem2 = UCase(Trim(txtItem2.Text))

If strItem1 = strItem2 Then
    cmdCheckRes.BackColor = vbBlue
    txtStatus = strItem1 & "........" & strItem2 & ": 扫描一致!!!" & vbCrLf & txtStatus
    txtItem1.Text = ""
    txtItem1.SetFocus
    txtItem2.Text = ""
    txtItem2.Visible = False
    
    strSql = "insert into UNIQUE_TBL_NEW(KEYNAME,KEYVALUE,KEYFROM,KEYTIME,KEYBY) values('临时一对一','" & strItem1 & "','37',sysdate,'" & gUserName & "')"
    AddSql (strSql)
Else
    cmdCheckRes.BackColor = vbRed
    txtStatus = strItem1 & "........" & strItem2 & ": 扫描错误!!!" & vbCrLf & txtStatus
    MsgBox "标签条码不一致", vbCritical, "警告"
    txtItem1.Text = ""
    txtItem1.SetFocus
    txtItem2.Text = ""
    txtItem2.Visible = False
    
    
    Exit Sub

End If
End Sub
