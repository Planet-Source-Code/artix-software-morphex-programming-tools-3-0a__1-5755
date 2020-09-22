VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Morphex Programming Tools"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin TabDlg.SSTab ssFunctions 
      Height          =   3855
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "Desktop "
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeDesktop"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "CD-Rom"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeCDRom"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Forms"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeForms"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Internet"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fmeInternet"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Key Buttons"
      TabPicture(4)   =   "frmMain.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fmeKeyButtons"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Taskbar"
      TabPicture(5)   =   "frmMain.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fmeTaskbar"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Options"
      TabPicture(6)   =   "frmMain.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fmeOptions"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.Frame fmeOptions 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   33
         Top             =   360
         Width           =   6855
         Begin VB.CommandButton cmdExtraDocumentation 
            Caption         =   "Extra &Documents"
            Height          =   375
            Left            =   2880
            TabIndex        =   40
            Top             =   2520
            Width           =   2535
         End
         Begin VB.CommandButton cmdWriter 
            Caption         =   "&Email Developer"
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   2520
            Width           =   2535
         End
         Begin VB.CommandButton cmdCleanUp 
            Caption         =   "&Clean Up Program"
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label lblAboutProgramInfo 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":00C4
            Height          =   495
            Left            =   840
            TabIndex        =   38
            Top             =   2040
            Width           =   5775
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   240
            Picture         =   "frmMain.frx":015C
            Top             =   2040
            Width           =   480
         End
         Begin VB.Label lblAboutProgram 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "About This Program:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   37
            Top             =   1800
            Width           =   1740
         End
         Begin VB.Label lblInfoAndSet 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":28FE
            Height          =   495
            Left            =   840
            TabIndex        =   35
            Top             =   480
            Width           =   5775
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   240
            Picture         =   "frmMain.frx":29A0
            Top             =   480
            Width           =   480
         End
         Begin VB.Label lblCleanAll 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clean Up Program:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   1605
         End
      End
      Begin VB.Frame fmeTaskbar 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   28
         Top             =   360
         Width           =   6855
         Begin VB.CommandButton cmdTaskbar 
            Caption         =   "Activate Item"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4200
            TabIndex        =   32
            Top             =   1080
            Width           =   2415
         End
         Begin VB.ListBox lstTaskbar 
            Height          =   1815
            ItemData        =   "frmMain.frx":5142
            Left            =   840
            List            =   "frmMain.frx":5155
            Sorted          =   -1  'True
            TabIndex        =   31
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label lblTaskbarOptions 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Options:"
            Height          =   195
            Left            =   840
            TabIndex        =   30
            Top             =   840
            Width           =   585
         End
         Begin VB.Label lblTaskbarWelcome 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":51C3
            Height          =   495
            Left            =   840
            TabIndex        =   29
            Top             =   240
            Width           =   5775
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   240
            Picture         =   "frmMain.frx":524E
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fmeKeyButtons 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   23
         Top             =   360
         Width           =   6855
         Begin VB.CommandButton cmdKeyButtons 
            Caption         =   "Activate Item"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4200
            TabIndex        =   27
            Top             =   1080
            Width           =   2415
         End
         Begin VB.ListBox lstKeyButtons 
            Height          =   1815
            ItemData        =   "frmMain.frx":5690
            Left            =   840
            List            =   "frmMain.frx":5697
            TabIndex        =   26
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label lblKeyButtonsOptions 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Options:"
            Height          =   195
            Left            =   840
            TabIndex        =   25
            Top             =   840
            Width           =   585
         End
         Begin VB.Label lblKeyButtonsWelcome 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":56B0
            Height          =   495
            Left            =   840
            TabIndex        =   24
            Top             =   240
            Width           =   5655
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   240
            Picture         =   "frmMain.frx":573F
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fmeInternet 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   18
         Top             =   360
         Width           =   6855
         Begin VB.CommandButton cmdInternet 
            Caption         =   "Activate Item"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4200
            TabIndex        =   22
            Top             =   1080
            Width           =   2415
         End
         Begin VB.ListBox lstInternet 
            Height          =   1815
            ItemData        =   "frmMain.frx":5A49
            Left            =   840
            List            =   "frmMain.frx":5A50
            TabIndex        =   21
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label lblInternetOptions 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Options:"
            Height          =   195
            Left            =   840
            TabIndex        =   20
            Top             =   840
            Width           =   585
         End
         Begin VB.Label lblInternetWelcome 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":5A61
            Height          =   495
            Left            =   840
            TabIndex        =   19
            Top             =   240
            Width           =   5775
         End
         Begin VB.Image imgEarth 
            Height          =   480
            Left            =   240
            Picture         =   "frmMain.frx":5AED
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fmeForms 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   13
         Top             =   360
         Width           =   6855
         Begin VB.CommandButton cmdForms 
            Caption         =   "Activate Item"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4200
            TabIndex        =   17
            Top             =   1080
            Width           =   2415
         End
         Begin VB.ListBox lstForms 
            Height          =   1815
            ItemData        =   "frmMain.frx":828F
            Left            =   840
            List            =   "frmMain.frx":8299
            TabIndex        =   16
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label lblFormsOptions 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Options:"
            Height          =   195
            Left            =   840
            TabIndex        =   15
            Top             =   840
            Width           =   585
         End
         Begin VB.Label lblFormsWelcome 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":82BE
            Height          =   495
            Left            =   840
            TabIndex        =   14
            Top             =   240
            Width           =   5775
         End
         Begin VB.Image imgForms 
            Height          =   480
            Left            =   240
            Picture         =   "frmMain.frx":8347
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fmeCDRom 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   8
         Top             =   360
         Width           =   6855
         Begin VB.CommandButton cmdCDRom 
            Caption         =   "Activate Item"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4200
            TabIndex        =   12
            Top             =   1080
            Width           =   2415
         End
         Begin VB.ListBox lstCDRom 
            Height          =   1815
            ItemData        =   "frmMain.frx":8F09
            Left            =   840
            List            =   "frmMain.frx":8F10
            TabIndex        =   11
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label lblCDRomOptions 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Options:"
            Height          =   195
            Left            =   840
            TabIndex        =   10
            Top             =   840
            Width           =   585
         End
         Begin VB.Label lblCDRomWelcome 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":8F28
            Height          =   495
            Left            =   840
            TabIndex        =   9
            Top             =   240
            Width           =   5775
         End
         Begin VB.Image imgCDRom 
            Height          =   480
            Left            =   240
            Picture         =   "frmMain.frx":8FB1
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fmeDesktop 
         Height          =   3255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   6855
         Begin VB.CommandButton cmdDesktop 
            Caption         =   "Activate Item"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4200
            TabIndex        =   7
            Top             =   1080
            Width           =   2415
         End
         Begin VB.ListBox lstDesktop 
            Height          =   1815
            ItemData        =   "frmMain.frx":B753
            Left            =   840
            List            =   "frmMain.frx":B75A
            TabIndex        =   6
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label lblDesktopOptions 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Options:"
            Height          =   195
            Left            =   840
            TabIndex        =   5
            Top             =   840
            Width           =   585
         End
         Begin VB.Label lblDesktopWelcome 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":B774
            Height          =   495
            Left            =   840
            TabIndex        =   4
            Top             =   240
            Width           =   5775
         End
         Begin VB.Image imgDesktop 
            Height          =   480
            Left            =   240
            Picture         =   "frmMain.frx":B7FF
            Top             =   240
            Width           =   480
         End
      End
   End
   Begin VB.Timer tmrLoading 
      Interval        =   1500
      Left            =   240
      Top             =   5760
   End
   Begin VB.Label lblTopWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":DFA1
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
   Begin VB.Image imgSearch 
      Height          =   480
      Left            =   240
      Picture         =   "frmMain.frx":E0AC
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SentCommand As Boolean

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal lpSize As Integer) As Long



Private Sub cmdCDRom_Click()
Select Case lstCDRom.Text
 Case "Control CDRom Door"
 frmCDRom1.Show
 frmMain.Enabled = False
End Select
End Sub

Private Sub cmdCleanUp_Click()
Call CleanUp
End Sub

Private Sub cmdDesktop_Click()
Select Case lstDesktop.Text
 Case "Remove Desktop Icons"
 frmDesktop1.Show
 frmMain.Enabled = False
End Select
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdExtraDocumentation_Click()
Dim strA As String
Dim intA As Integer

strA = Space$(256)
intA = GetWindowsDirectory(strA, 256)
strA = Left$(strA, intA)
Shell strA + "\Notepad.exe " + App.Path + "\releasenotes.txt", vbMaximizedFocus

End Sub

Private Sub cmdForms_Click()
Select Case lstForms.Text
 Case "Center Forms"
 frmForms1.Show
 frmMain.Enabled = False
 Case "Place Forms OnTop"
 frmForms2.Show
 frmMain.Enabled = False
End Select
End Sub

Private Sub cmdInternet_Click()
Select Case lstInternet.Text
 Case "Execute URL"
 frmInternet1.Show
 frmMain.Enabled = False
End Select
End Sub

Private Sub cmdKeyButtons_Click()
Select Case lstKeyButtons.Text
 Case "Disable Key Buttons"
 frmKeyButtons1.Show
 frmMain.Enabled = False
End Select
End Sub

Private Sub cmdTaskbar_Click()
Select Case lstTaskbar.Text
 Case "Remove Start Button"
 frmTaskbar1.Show
 frmMain.Enabled = False
 Case "Remove Taskbar Icons"
 frmTaskbar2.Show
 frmMain.Enabled = False
 Case "Remove Taskbar Programs"
 frmTaskbar3.Show
 frmMain.Enabled = False
 Case "Remove Taskbar Clock"
 frmTaskbar4.Show
 frmMain.Enabled = False
 Case "Remove Taskbar"
 frmTaskbar5.Show
 frmMain.Enabled = False
End Select
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdWriter_Click()
OpenURL "mailto:ocxm@netzero.net?subject=Software Reply"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim SendValue As Long
SendValue = MsgBox("Are you sure that you want to exit this program now?", vbQuestion + vbYesNo, "Exit")
If SendValue = vbYes Then
Cancel = 0
Else
Cancel = 1
End If
End Sub

Private Sub lstCDRom_Click()
cmdCDRom.Enabled = True
End Sub

Private Sub lstDesktop_Click()
cmdDesktop.Enabled = True
End Sub

Private Sub lstForms_Click()
cmdForms.Enabled = True
End Sub

Private Sub lstInternet_Click()
cmdInternet.Enabled = True
End Sub

Private Sub lstKeyButtons_Click()
cmdKeyButtons.Enabled = True
End Sub

Private Sub lstTaskbar_Click()
cmdTaskbar.Enabled = True
End Sub

Private Sub tmrLoading_Timer()
Unload frmSplash
tmrLoading.Enabled = False
End Sub
