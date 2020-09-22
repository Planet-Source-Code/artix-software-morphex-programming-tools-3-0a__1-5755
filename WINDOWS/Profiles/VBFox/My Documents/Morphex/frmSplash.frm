VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1320
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5145
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0582
   ScaleHeight     =   1320
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoading 
      Interval        =   1500
      Left            =   4560
      Top             =   120
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum OSVersion_Constants
 Compatible = 1
 NotCompatible = 0
End Enum

Dim OSType As OSVersion_Constants

Private Sub Form_Load()
Dim OSVersion As String
OSVersion = Left$(GetWindowsVersion, 19)
If OSVersion = "Windows 98 version:" Or OSVersion = "Windows 95 version:" Then
OSType = Compatible
Else
OSType = NotCompatible
End If
If OSType = NotCompatible Then
MsgBox "Your operating system is not compatible with this program. Please use this program on the Microsoft Windows 95/98 operating system.", vbCritical, "Operating System"
End
End If
End Sub

Private Sub tmrLoading_Timer()
frmMain.Show
frmSplash.ZOrder 0
frmMain.ZOrder 1
tmrLoading.Enabled = False
End Sub
