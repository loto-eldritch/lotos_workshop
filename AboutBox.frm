VERSION 5.00
Begin VB.Form AboutBox 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl_STitle2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Special Thanks:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   7290
      TabIndex        =   6
      Top             =   1830
      Width           =   1395
   End
   Begin VB.Label lbl_STitle1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Special Thanks:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   1830
      Width           =   1395
   End
   Begin VB.Label lbl_Mythic 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   3090
      Width           =   8685
   End
   Begin VB.Label lbl_Special2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   6495
      TabIndex        =   3
      Top             =   2010
      Width           =   2175
   End
   Begin VB.Label lbl_Special1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   2010
      Width           =   2175
   End
   Begin VB.Label lbl_WebSiteUrl 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "http://www.experimental-playground.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   2820
      Width           =   4350
   End
   Begin VB.Label lbl_About 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   2175
      TabIndex        =   0
      Top             =   1980
      Width           =   4320
   End
   Begin VB.Image Image_About 
      Height          =   1980
      Left            =   0
      Picture         =   "AboutBox.frx":0000
      Top             =   0
      Width           =   8670
   End
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SpecialList1 As String = "Maala - Igraine" & vbCrLf & _
                                        "Naoie - Igraine" & vbCrLf & _
                                        "Unseenn - Igraine" & vbCrLf & _
                                        "Serenitee - Ywain"
                                       
Private Const SpecialList2 As String = "Broniol - Igraine" & vbCrLf & _
                                        "Clanmaster - Igraine" & vbCrLf & _
                                        "Due - Igraine"
                                       
Private Const MYTHIC_COPYRIGHT_NOTICE As String = "Icons, screenshots, logos, and interface samples are all part of Dark Age of Camelot and therefore the sole property of EA Mythic. Used only by permission."

Private Sub Form_Load()

    Call StayOnTop(Me)
    
    lbl_About.Caption = ABOUT_BOX_MESSAGE
    
    lbl_WebSiteUrl.Caption = WEBSITE
    lbl_WebSiteUrl.Tag = WEBSITE_FORUM
    
    lbl_Special1.Caption = SpecialList1
    lbl_Special2.Caption = SpecialList2
    lbl_Mythic.Caption = MYTHIC_COPYRIGHT_NOTICE
    
End Sub

Private Sub lbl_Mythic_Click()

    Unload Me
    
End Sub

Private Sub lbl_Special1_Click()

    Unload Me
    
End Sub

Private Sub lbl_Special2_Click()

    Unload Me
    
End Sub

Private Sub lbl_WebSiteUrl_Click()

    Call JumpToWeb(lbl_WebSiteUrl.Tag)
    Unload Me
    
End Sub

Private Sub Form_Click()

    Unload Me
    
End Sub

Private Sub Image_About_Click()

    Unload Me
    
End Sub

Private Sub lbl_About_Click()

    Unload Me
    
End Sub



