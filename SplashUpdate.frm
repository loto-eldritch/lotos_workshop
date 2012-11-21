VERSION 5.00
Begin VB.Form SplashUpdate 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Loto's Character Workshop - Update"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3750
      Top             =   1020
   End
   Begin VB.Label lbl_Version 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   7470
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label ProgressBar 
      BackColor       =   &H00FF8080&
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Tag             =   "8610"
      Top             =   2025
      Width           =   8610
   End
   Begin VB.Image Image_Splash 
      Height          =   1980
      Left            =   0
      Picture         =   "SplashUpdate.frx":0000
      Top             =   0
      Width           =   8670
   End
   Begin VB.Label lbl_ProgressBG 
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   1980
      Width           =   8715
   End
   Begin VB.Label Status 
      BackColor       =   &H00000000&
      Caption         =   "Initializing..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   225
      Left            =   45
      TabIndex        =   2
      Top             =   2280
      Width           =   8670
   End
End
Attribute VB_Name = "SplashUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    'Call Image_Splash_Click
End Sub

Private Sub Form_Load()
    TimerUpdate.Enabled = True
    CHECK_UPDATE = False
    
End Sub

Private Sub Image_Splash_Click()
    'Unload Me
End Sub

Private Sub TimerUpdate_Timer()
    If Not CHECK_UPDATE Then
        CHECK_UPDATE = True
        TimerUpdate.Enabled = False
        Call AutoUpdate(ProgressBar, Status, Me.hWnd)
        CHECK_UPDATE = False
        Unload Me
    End If
End Sub
