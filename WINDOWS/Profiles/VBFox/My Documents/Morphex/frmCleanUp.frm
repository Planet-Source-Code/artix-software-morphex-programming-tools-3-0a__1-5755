VERSION 5.00
Begin VB.Form frmCleanUp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCleaning 
      Interval        =   2566
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait...Cleaning Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   280
      Width           =   4695
   End
End
Attribute VB_Name = "frmCleanUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub

Private Sub tmrCleaning_Timer()
Unload Me
End Sub
