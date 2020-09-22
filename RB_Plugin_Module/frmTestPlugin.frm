VERSION 5.00
Begin VB.Form frmTestPlugin 
   Caption         =   "Plug In Host Form"
   ClientHeight    =   4545
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4410
   Icon            =   "frmTestPlugin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox FileListHidden 
      Height          =   1650
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Host to Plugins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   3840
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "click on the Plug-Ins menu to run the detected plugins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plugins in the app directory"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "To register a plugin, just double click on the exe file. The ActiveX exe file will the make itself available in the rot."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuallPlugins 
      Caption         =   "Plug-ins"
      Begin VB.Menu mnuPlugin 
         Caption         =   "About Plugins"
         Index           =   0
      End
      Begin VB.Menu mnuPlugin 
         Caption         =   "-"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmTestPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

'Call the generic code to add the plugins to the menu
AddPlugins Me, FileListHidden

' "Me" is this form that we pass, as an object of course
' "filelisthidden" is a filelist control that we use to detect all the
' exe files in the directory.

End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuPlugin_Click(Index As Integer)

If Index > 1 Then ' we have plugins in the menu
  Call RunPlugin(mnuPlugin(Index).Tag, Me) ' Execute the plug-in
End If

End Sub
