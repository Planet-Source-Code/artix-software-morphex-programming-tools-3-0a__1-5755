VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTaskbar3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remove Taskbar Programs"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtbSource 
      Height          =   1935
      Left            =   840
      TabIndex        =   6
      Top             =   2400
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmTaskbar3.frx":0000
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton optRemovePrograms 
      Caption         =   "Remove Taskbar Programs"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.OptionButton optProgramsVisible 
      Caption         =   "Taskbar Programs Are Visible"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source Code:"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblTopWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "Select below if you would like to remove the Taskbar Programs. In order for this operation to work, you must be viewing this form."
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmTaskbar3.frx":00C9
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmTaskbar3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SentCommand As Boolean

Private Sub cmdApply_Click()
If optProgramsVisible.Value = True Then
TaskbarPrograms IsOnTaskbar
SentCommand = False
cmdApply.Enabled = False
ElseIf optRemovePrograms.Value = True Then
On Error GoTo vbHandler
TaskbarPrograms InNotOnTaskbar
SentCommand = True
cmdApply.Enabled = False
End If
Exit Sub
vbHandler:
MsgBox "There was an error while trying to perform this function. The task was not able to be completed.", vbCritical, "Error"
Exit Sub
End Sub

Private Sub cmdContinue_Click()
Unload Me
End Sub

Private Sub Form_Load()
SentCommand = False
rtbSource.FileName = App.Path + "\iref_removetaskbarprograms.rtf"
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub

Private Sub optProgramsVisible_Click()
If SentCommand = True Then
cmdApply.Enabled = True
ElseIf SentCommand = False Then
cmdApply.Enabled = False
End If
End Sub

Private Sub optRemovePrograms_Click()
If SentCommand = True Then
cmdApply.Enabled = False
ElseIf SentCommand = False Then
cmdApply.Enabled = True
End If
End Sub
