VERSION 5.00
Begin VB.Form FrmLabelPrinting_37ToHW 
   Caption         =   "��ǩ��ӡϵͳ_37����Ϊ"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13785
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
   ScaleHeight     =   10440
   ScaleWidth      =   13785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Fra 
      Caption         =   "��ɨ����ϸ"
      ForeColor       =   &H00FF0000&
      Height          =   9135
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   13815
   End
   Begin VB.Frame Frame1 
      Caption         =   "�˵�ѡ��"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      Begin VB.TextBox txtScan 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   435
         Width           =   3255
      End
      Begin VB.Label lblScan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ɨ��"
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   360
      End
   End
End
Attribute VB_Name = "FrmLabelPrinting_37ToHW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
