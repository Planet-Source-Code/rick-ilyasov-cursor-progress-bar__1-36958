VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Cursor Bar"
      Height          =   1515
      Left            =   120
      TabIndex        =   5
      Top             =   2460
      Width           =   5835
      Begin VB.CommandButton Command2 
         Caption         =   "Show Progress"
         Height          =   315
         Left            =   2220
         TabIndex        =   6
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label3 
         BackColor       =   &H00DDEDEE&
         Caption         =   "This is how Cursor Progress Bar Works, showing the progress of the same For Loop."
         Height          =   1155
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5595
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conventional"
      Height          =   1515
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5835
      Begin VB.CommandButton Command1 
         Caption         =   "Show Progress"
         Height          =   315
         Left            =   2220
         TabIndex        =   4
         Top             =   600
         Width           =   1395
      End
      Begin MSComctlLib.ProgressBar pbConventional 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DDEDEE&
         Caption         =   "This is how conventional Progress Bar looks and Works, showing the progress of a For Loop."
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5595
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cursor Progress Bar Sample"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    'Show Progress using Standard Bar
    
    Dim I As Long
    
    pbConventional.Min = 0
    pbConventional.Max = 250000
    
    For I = 0 To 250000
        pbConventional.Value = I
    Next
    
    pbConventional.Value = 0
    
    
End Sub

Private Sub Command2_Click()
    
    Dim I As Long
    
    'I guess there are two extra lines of code, oh well...
    Dim pbCursor As CursorObject.ProgressBar
    Set pbCursor = New CursorObject.ProgressBar
    
    pbCursor.Min = 0
    pbCursor.Max = 250000
    
    For I = 0 To 250000
        pbCursor.Value = I
    Next
    
    'MAKE SURE IT GOES ALL THE WAY
    pbCursor.Value = pbCursor.Max
        
End Sub
