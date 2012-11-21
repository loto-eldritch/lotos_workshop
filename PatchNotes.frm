VERSION 5.00
Begin VB.Form PatchNotes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Patch Notes!"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6465
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
   ScaleHeight     =   3570
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_PatchNotes 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   3570
      Left            =   7
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6450
   End
End
Attribute VB_Name = "PatchNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GetPatchNotes()

    Dim lResult As Long
    Dim sBuffer As String
    Dim sMsg As String
    
    If Connection_Online Then
        
        sMsg = "Loading Patch Notes..."
        txt_PatchNotes.Text = sMsg
        
        sBuffer = Download_File(PATCH_NOTES_LOC)
                
        txt_PatchNotes.Text = Replace(sBuffer, Chr$(10), vbCrLf)
        
    Else
        Status.Caption = "No Internet Connection!"
    End If

End Sub

Private Sub Form_Load()
    Call GetPatchNotes
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
