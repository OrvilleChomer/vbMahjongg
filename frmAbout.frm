VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VB Master Mahjongg (BitBlt Edition)"
   ClientHeight    =   3600
   ClientLeft      =   1650
   ClientTop       =   2325
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&CLOSE"
      Default         =   -1  'True
      Height          =   375
      Left            =   2460
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblWebsite 
      BackStyle       =   0  'Transparent
      Caption         =   "http://vbmaster.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2220
      MouseIcon       =   "frmAbout.frx":E1CF
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "click to visit VB Master!"
      Top             =   1020
      Width           =   1755
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(c) Copyright 2000, Orville Chomer, Chomer.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   660
      TabIndex        =   2
      Top             =   660
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VB Master Mahjongg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   600
      TabIndex        =   1
      Top             =   60
      Width           =   4875
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cPiMahjongg     As clsMahjongg

Private Sub cmdOK_Click()
    Unload Me
    
End Sub


Private Sub Form_Activate()
    cPiMahjongg.PauseGame
    
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
End Sub



Public Sub PassObject(ByVal ciObj As clsMahjongg)
    Set cPiMahjongg = ciObj
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cPiMahjongg.ResumeGame
    
End Sub


Private Sub lblWebsite_Click()
   Shell "c:\program files\internet explorer\iexplore.exe http://vbmaster.net "
End Sub


