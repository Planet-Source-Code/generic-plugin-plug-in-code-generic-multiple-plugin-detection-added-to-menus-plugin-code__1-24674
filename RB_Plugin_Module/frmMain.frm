VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Plugin"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Increase Hosts Size"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Decrease Hosts Size"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Host"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide Host"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Host's Caption"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Plugin"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public objHost As Object

Private Sub Command1_Click()
objHost.Visible = False
End Sub

Private Sub Command2_Click()
objHost.Visible = True
End Sub

Private Sub Command3_Click()
objHost.Caption = Text1
End Sub

Private Sub Command4_Click()
objHost.Width = objHost.Width - 500
End Sub

Private Sub Command5_Click()
objHost.Width = objHost.Width + 500
End Sub
