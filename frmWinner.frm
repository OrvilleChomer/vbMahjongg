VERSION 5.00
Begin VB.Form frmWinner 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "You are a Winner!!!"
   ClientHeight    =   2160
   ClientLeft      =   1890
   ClientTop       =   3135
   ClientWidth     =   5625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTopScore 
      Height          =   735
      Left            =   180
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   5235
      Begin VB.TextBox txtPlayer 
         Height          =   285
         Left            =   3780
         MaxLength       =   8
         TabIndex        =   7
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "You have a top score! Please enter your name:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdCLOSEPROG 
      Caption         =   "CLOSE PROGRAM"
      Height          =   315
      Left            =   3300
      TabIndex        =   2
      Top             =   1740
      Width           =   1695
   End
   Begin VB.CommandButton cmdPLAYANOTHER 
      Caption         =   "PLAY ANOTHER"
      Default         =   -1  'True
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   1740
      Width           =   1635
   End
   Begin VB.Label lblFinalTime 
      Caption         =   "???"
      Height          =   195
      Left            =   4620
      TabIndex        =   4
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Your final time was:"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "You have won the game!"
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   2475
   End
End
Attribute VB_Name = "frmWinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bPiCloseProgram          As Boolean
Private nPiMinutes               As Integer
Private nPiSeconds               As Integer

Private Sub cmdCLOSEPROG_Click()


    bPiCloseProgram = True
    
    Me.Hide
    
End Sub

Private Sub cmdPLAYANOTHER_Click()
    bPiCloseProgram = False

    Me.Hide
End Sub


Public Property Get LastWinner() As String

End Property

Public Property Let LastWinner(ByVal vNewValue As String)

End Property

Public Property Get WinnerName() As String
    WinnerName = txtPlayer.Text
End Property



Public Property Get EndProgram() As Boolean
    EndProgram = bPiCloseProgram
    
End Property


Public Sub TopScore()
    'CALLED IF USER GOT A TOP SCORE!
    
    cmdPLAYANOTHER.Enabled = False
    cmdCLOSEPROG.Enabled = False
    fraTopScore.Visible = True
    
End Sub

Public Property Get Minutes() As Integer
    Minutes = nPiMinutes
End Property

Public Property Let Minutes(ByVal vNewValue As Integer)
    nPiMinutes = vNewValue
    
End Property

Public Property Get Seconds() As Integer
    Seconds = nPiSeconds
End Property

Public Property Let Seconds(ByVal vNewValue As Integer)
    nPiSeconds = vNewValue
    
End Property

Private Sub Form_Activate()
    If fraTopScore.Visible Then
        txtPlayer.SetFocus
    End If
    
End Sub

Private Sub Form_Load()
    'CENTER FORM
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    
End Sub



Public Sub Display()
    lblFinalTime = Format$(nPiMinutes, "00") & ":" & Format$(nPiSeconds, "00")
    
End Sub

Private Sub txtPlayer_Change()
    If Trim$(txtPlayer) = "" Then
        cmdPLAYANOTHER.Enabled = False
        cmdCLOSEPROG.Enabled = False
    Else
        cmdPLAYANOTHER.Enabled = True
        cmdCLOSEPROG.Enabled = True
    End If
    
End Sub


