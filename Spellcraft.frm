VERSION 5.00
Begin VB.Form SC 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loto's Character Workshop - Spellcraft Calculator"
   ClientHeight    =   6420
   ClientLeft      =   6885
   ClientTop       =   8625
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00008000&
   Icon            =   "Spellcraft.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Spellcraft.frx":0CCA
   ScaleHeight     =   6420
   ScaleWidth      =   10335
   Visible         =   0   'False
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuBreak_File_01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "&Recent"
         Begin VB.Menu mnuRecentFile 
            Caption         =   "-"
            Index           =   0
         End
      End
      Begin VB.Menu mnuBreak_File_02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuBreak_File_03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuBuilder 
         Caption         =   "Character &Builder"
      End
      Begin VB.Menu mnuBreak_Tools_01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemDatabase 
         Caption         =   "Item &Database"
      End
      Begin VB.Menu mnuBreak_Tools_02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSwapMain 
         Caption         =   "&Swap Gems With..."
         Begin VB.Menu mnuSwp 
            Caption         =   "&Chest"
            Index           =   1
         End
         Begin VB.Menu mnuSwp 
            Caption         =   "&Arms"
            Index           =   2
         End
         Begin VB.Menu mnuSwp 
            Caption         =   "&Head"
            Index           =   3
         End
         Begin VB.Menu mnuSwp 
            Caption         =   "&Legs"
            Index           =   4
         End
         Begin VB.Menu mnuSwp 
            Caption         =   "&Hands"
            Index           =   5
         End
         Begin VB.Menu mnuSwp 
            Caption         =   "&Feet"
            Index           =   6
         End
         Begin VB.Menu mnuSwp_Break1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSwpWep 
            Caption         =   "&Main Hand"
            Index           =   1
         End
         Begin VB.Menu mnuSwpWep 
            Caption         =   "&Off Hand"
            Index           =   2
         End
         Begin VB.Menu mnuSwpWep 
            Caption         =   "&2 Handed"
            Index           =   3
         End
         Begin VB.Menu mnuSwpWep 
            Caption         =   "&Ranged"
            Index           =   4
         End
         Begin VB.Menu mnuSwp_Break 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuSwpSpare 
            Caption         =   "Main Spare"
            Index           =   1
         End
         Begin VB.Menu mnuSwpSpare 
            Caption         =   "Offhand Spare"
            Index           =   2
         End
         Begin VB.Menu mnuSwpSpare 
            Caption         =   "2Handed Spare"
            Index           =   3
         End
         Begin VB.Menu mnuSwpSpare 
            Caption         =   "Ranged Spare"
            Index           =   4
         End
      End
      Begin VB.Menu mnuSetCraftBars 
         Caption         =   "Setup &Craft Bars"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuConfiguration 
         Caption         =   "&Configuration"
      End
      Begin VB.Menu mnuMaterials 
         Caption         =   "&Materials"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "&Topics"
      End
      Begin VB.Menu mnuHelp_Break 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPatchNotes 
         Caption         =   "Patch Notes"
      End
   End
End
Attribute VB_Name = "SC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
