VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   Caption         =   "VB Master Mahjongg - Version 1.0 (BitBLt Edition)"
   ClientHeight    =   4935
   ClientLeft      =   2340
   ClientTop       =   3705
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MouseIcon       =   "frmMain.frx":0000
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList"
      _Version        =   327682
      BorderStyle     =   1
   End
   Begin VB.PictureBox picStatus 
      BackColor       =   &H0080FF80&
      Height          =   375
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   445
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Label lblBulletin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO MORE MOVES ARE LEFT!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4980
         TabIndex        =   19
         Top             =   60
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lblTilesLeft1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tiles Left:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   60
         Width           =   690
      End
      Begin VB.Label lblTilesLeft2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   900
         TabIndex        =   17
         Top             =   60
         Width           =   270
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   16
         Top             =   60
         Width           =   390
      End
      Begin VB.Label lblTime2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2580
         TabIndex        =   15
         Top             =   60
         Width           =   405
      End
      Begin VB.Label lblMovesLeft 
         BackStyle       =   0  'Transparent
         Caption         =   "Moves Left:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3420
         TabIndex        =   14
         Top             =   60
         Width           =   915
      End
      Begin VB.Label lblMovesLeft2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4440
         TabIndex        =   13
         Top             =   60
         Width           =   90
      End
   End
   Begin VB.Timer timBulletin 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   2700
      Top             =   1200
   End
   Begin VB.PictureBox picSideBar 
      BackColor       =   &H00C0FFC0&
      Height          =   2100
      Left            =   5175
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   2040
      ScaleWidth      =   1335
      TabIndex        =   7
      Top             =   2175
      Visible         =   0   'False
      Width           =   1395
      Begin VB.PictureBox picTopScores 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   1155
         TabIndex        =   9
         Top             =   660
         Width           =   1155
         Begin VB.Label lblScore 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   11
            Top             =   300
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   10
            Top             =   300
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Label lblTopScores 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TOP SCORES:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   8
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.PictureBox picBrightTiles 
      AutoRedraw      =   -1  'True
      Height          =   1455
      Left            =   3300
      Picture         =   "frmMain.frx":E4D9
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer timClock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2700
      Top             =   660
   End
   Begin VB.PictureBox picSel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF8080&
      Height          =   915
      Left            =   1140
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   5
      Top             =   1260
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox picTileEdgeMask 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   2160
      Picture         =   "frmMain.frx":4ADDD
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   2580
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picTileEdge 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   1260
      Picture         =   "frmMain.frx":4C973
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   3
      Top             =   2580
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox picGame 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF8080&
      Height          =   555
      Left            =   0
      Picture         =   "frmMain.frx":4E509
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox picDarkTiles 
      AutoRedraw      =   -1  'True
      Height          =   1395
      Left            =   4920
      Picture         =   "frmMain.frx":68646
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.PictureBox picTiles 
      AutoRedraw      =   -1  'True
      Height          =   1395
      Left            =   3300
      Picture         =   "frmMain.frx":A4F4A
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   1515
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   2100
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":E184E
            Key             =   "smiley"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgDim 
      Height          =   435
      Left            =   1320
      Tag             =   "Used to calc image dimensions"
      Top             =   420
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBackup 
         Caption         =   "&Back up a move"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileHint 
         Caption         =   "Show &Hint"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuFileDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuPopupNew 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuPopupBackup 
         Caption         =   "&Back up a move"
      End
      Begin VB.Menu mnuPopupHint 
         Caption         =   "Show &Hint"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#####################################################
'   VB MASTER MAHJONGG (BITBLT EDITION)
'
'   WRITTEN BY: ORVILLE CHOMER
'   DATE:       JULY 31, 2000
'
'   (C) COPYRIGHT 2000, ORVILLE CHOMER, CHOMER.COM
'
'#####################################################

'#####################################################
' Usage: Developer may not resell this application
'        without the written permission of the author.
'        Developer may copy code in this app for their own
'        applications. Developer may redistribute this source
'        and/or executable as long as source code including
'        these unmodified remarks. The developer is
'        solely responsable for making sure any app using
'        the code herein is tested and debugged.
'######################################################

'######################################################
'     DOWNLOADED FROM:   VBMASTER.NET
'          ENJOY!
'######################################################

Private WithEvents cPiMahjongg    As clsMahjongg
Attribute cPiMahjongg.VB_VarHelpID = -1

Private Sub cPiMahjongg_GameWasPaused()
    timClock.Enabled = False
    
End Sub


Private Sub cPiMahjongg_GameWasResumed()
    timClock.Enabled = True
End Sub


Private Sub cPiMahjongg_NewGame()
    lblBulletin.Visible = False
    lblBulletin.Refresh
    lblBulletin.Left = lblMovesLeft2.Left + lblMovesLeft2.Width
    lblBulletin.Caption = "New Game Started..."
    lblBulletin.Visible = True
    timBulletin.Enabled = True

End Sub

Private Sub cPiMahjongg_NoMoves()
    lblBulletin.Visible = False
    lblBulletin.Refresh
    lblBulletin.Left = lblMovesLeft2.Left + lblMovesLeft2.Width
    lblBulletin.Caption = "No Moves Are Left!"
    lblBulletin.Visible = True
    timBulletin.Enabled = True
    
End Sub

Private Sub cPiMahjongg_PiecesTaken()
    lblMovesLeft2.Caption = cPiMahjongg.MovesLeft
    
End Sub

Private Sub cPiMahjongg_Winner()
    Dim bTopScore          As Boolean
    
    bTopScore = cPiMahjongg.IsTopScore()
    
    If bTopScore Then
        frmWinner.TopScore
    End If
    
    frmWinner.Seconds = cPiMahjongg.Seconds
    frmWinner.Minutes = cPiMahjongg.Minutes
    frmWinner.Display
    
    frmWinner.Show vbModal
    
    If bTopScore Then
        cPiMahjongg.SaveTopScores (frmWinner.WinnerName)
    End If
    
    'USER DECIDES TO QUIT!
    If frmWinner.EndProgram Then
        Unload frmWinner
        Unload Me
        Exit Sub
    End If
    
    Unload frmWinner
       
    NewGame
    
End Sub

Private Sub Form_DblClick()
    Form_MouseUp vbLeftButton, 0, 0, 0
    
    
End Sub


Private Sub Form_Initialize()
    Set cPiMahjongg = New clsMahjongg
    
    cPiMahjongg.SetRefs _
       Me, _
       picTiles, _
       picBrightTiles, _
       picDarkTiles, _
       picSel, _
       picGame, _
       lblTime2, _
       lblTilesLeft2

             
    
    

End Sub

Private Sub Form_Load()
    
    Dim n               As Integer
    
    mnuPopup.Visible = False
    
    
    'Dynamically Add buttons to the toolbar.
    Toolbar1.Buttons.Add , "new", "New Game", , "smiley"
    Toolbar1.Buttons.Add , "back", "Back Up", , "smiley"
    Toolbar1.Buttons.Add , "hint", "Hint", , "smiley"
    
    Me.Left = 0
    Me.Top = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height - 300
    
    
    picTiles.Width = 800
    picTiles.Height = 600
    picTiles.Top = 0
    picTiles.Left = 0
    
    
    picDarkTiles.Width = picTiles.Width
    picDarkTiles.Height = picTiles.Height
    
    picBrightTiles.Width = picTiles.Width
    picBrightTiles.Height = picTiles.Height
    
    picSel.Width = ScaleWidth
    picSel.Height = ScaleHeight
    
    For n = 0 To 9
        If n > 0 Then
            Load lblName(n)
            Load lblScore(n)
        End If
        
        lblName(n).Top = n * lblName(n).Height + 2
        lblScore(n).Top = lblName(n).Top
        lblScore(n).Left = 1300
        
    Next n
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim n              As Integer
    
    n = cPiMahjongg.CurrentTile(x, y - 28)
    
    If n <> cPiMahjongg.ActiveTile Then
        cPiMahjongg.ActiveTile = n
        'DisplayTiles n
        cPiMahjongg.HighlightTile
        
    End If
    
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup
        Exit Sub
    End If
    
    
    cPiMahjongg.ClickGame
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    cPiMahjongg.CleanUp
    
    timClock.Enabled = False
    
    Set cPiMahjongg = Nothing
    
    
    'Make window cleanly disappear:
    Me.Hide
    Me.Refresh
    DoEvents
    DoEvents
End Sub


Private Sub Form_Resize()
    If WindowState = vbMinimized And Not cPiMahjongg.GamePaused Then
        'IF WINDOW IS MINIMIZED, PAUSE THE GAME
        '(WE CAN'T SEE IT AFTER ALL!)
        cPiMahjongg.PauseGame
        Exit Sub
    End If
    
    If Not Me.Visible Then Exit Sub
    
    
    If cPiMahjongg.GamePaused Then
        cPiMahjongg.ResumeGame
    End If
    
    picGame.Width = Me.ScaleWidth
    picGame.Height = Me.ScaleHeight '- picGame.Top - 100
    picStatus.Width = ScaleWidth
    
    If Not cPiMahjongg.GameInProgress Then
        cPiMahjongg.NewGame
        lblMovesLeft2.Caption = cPiMahjongg.MovesLeft
        
    End If
    
    picStatus.Top = Me.ScaleHeight - lblTilesLeft1.Height - 8 - 5
    
    
    picSideBar.Width = 155
    
    picSideBar.Left = ScaleWidth - picSideBar.Width
    picSideBar.Top = 28
    picSideBar.Height = ScaleHeight - (picSideBar.Top * 2) + 3
    
    lblTopScores.Width = picSideBar.ScaleWidth
    
    picTopScores.Width = lblTopScores.Width - 160
    picTopScores.Left = (picSideBar.ScaleWidth - picTopScores.Width) / 2
    picTopScores.Height = picTopScores.Width * 1.25
    
    If Not picSideBar.Visible Then
        picSideBar.Visible = True
    End If
    
    If Not picStatus.Visible Then
        picStatus.Visible = True
    End If
End Sub












Private Sub mnuFileBackup_Click()
    Backup
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
    
End Sub

Private Sub mnuFileHint_Click()
    ShowHint
End Sub

Private Sub mnuFileNew_Click()
    NewGame
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.PassObject cPiMahjongg
    
    frmAbout.Show
    
End Sub

Private Sub mnuPopupBackup_Click()
    Backup
End Sub

Private Sub mnuPopupHint_Click()
    ShowHint
End Sub

Private Sub mnuPopupNew_Click()
    NewGame
End Sub



Private Sub picSideBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cPiMahjongg.ClearCursor
    
End Sub


Private Sub timBulletin_Timer()

    '*** SCROLL THE BULLETIN TO THE RIGHT!
    timBulletin.Enabled = False
    
    lblBulletin.Left = lblBulletin.Left + 8
    lblBulletin.Refresh
    
    If lblBulletin.Left + lblBulletin.Width + 2 < picStatus.ScaleWidth Then
        timBulletin.Enabled = True
    End If
End Sub

Private Sub timClock_Timer()
    timClock.Enabled = False
    
    'SHUT THE CLOCK COMPLETELY OFF
    'IF THE GAME IS DONE!
    If cPiMahjongg.GameComplete Then Exit Sub
    
    ' TEMPERARILY STOP CLOCK IF THE
    ' GAME IS PAUSED.
    If cPiMahjongg.GamePaused Then Exit Sub
    
    cPiMahjongg.DisplayTime
    timClock.Enabled = True
    
End Sub





Private Sub NewGame()
    cPiMahjongg.NewGame
    lblMovesLeft2.Caption = cPiMahjongg.MovesLeft
End Sub

Private Sub Backup()
    cPiMahjongg.Backup
    lblMovesLeft2.Caption = cPiMahjongg.MovesLeft
    
End Sub

Private Sub ShowHint()
    cPiMahjongg.ShowHint
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "new"
            NewGame
        Case "back"
            Backup
        Case "hint"
            ShowHint
            
    End Select
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cPiMahjongg.ClearCursor
End Sub


