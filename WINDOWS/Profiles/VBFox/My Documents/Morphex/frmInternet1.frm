VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInternet1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Execute URL"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmInternet1.frx":0000
   ScaleHeight     =   4575
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtbSource 
      Height          =   1815
      Left            =   840
      TabIndex        =   8
      Top             =   2520
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmInternet1.frx":08CA
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtMIME 
      Height          =   315
      Left            =   4920
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "com"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   840
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "www.planet-source-code"
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source Code:"
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblDot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   195
      Left            =   4800
      TabIndex        =   4
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label lblClickURL 
      BackStyle       =   0  'Transparent
      Caption         =   "http://"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   840
      MouseIcon       =   "frmInternet1.frx":0993
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label lblURLName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URL(Address):"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label lblTopWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmInternet1.frx":125D
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmInternet1.frx":12EA
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmInternet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdContinue_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblClickURL.Caption = "http://" + txtURL.Text + "." + txtMIME.Text
rtbSource.FileName = App.Path + "\iref_openurl.rtf"
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub

Private Sub lblClickURL_Click()
On Error GoTo vbHandler
OpenURL lblClickURL.Caption
Exit Sub
vbHandler:
MsgBox "There was an error while trying to perform this function. The task was not able to be completed.", vbCritical, "Error"
Exit Sub
End Sub

Private Sub txtMIME_Change()
lblClickURL.Caption = "http://" + txtURL.Text + "." + txtMIME.Text
End Sub

Private Sub txtURL_Change()
lblClickURL.Caption = "http://" + txtURL.Text + "." + txtMIME.Text
End Sub
