VERSION 5.00
Begin VB.Form frmSplashScreen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   240
      Top             =   2880
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    Screen.MousePointer = 13
End Sub

Private Sub Timer1_Timer()
    Load frmMain
    Unload frmSplashScreen
    frmMain.Show
End Sub
