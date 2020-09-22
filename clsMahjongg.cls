VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMahjongg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


    '** API Used to play wav files:
    Private Declare Function sndPlaySound& Lib "winmm.dll" Alias _
       "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long)
  
  ' WAV Sound values
  Private Const SND_SYNC = &H0
  Private Const SND_ASYNC = &H1
  Private Const SND_NODEFAULT = &H2
  Private Const SND_LOOP = &H8
  Private Const SND_NOSTOP = &H10
  
  Private Const SRCCOPY = &HCC0020
  Private Const SRCPAINT = &HEE0086
  Private Const SRCAND = &H8800C6

Private Declare Function BitBlt Lib "gdi32" _
  (ByVal hDestDC As Long, ByVal x As Long, _
  ByVal y As Long, ByVal nWidth As Long, _
  ByVal nHeight As Long, ByVal hSrcDC As Long, _
  ByVal xSrc As Long, ByVal ySrc As Long, _
  ByVal dwRop As Long) As Long




'Defines a tile and its state
Private Type udtTile
   TileID                 As Integer
   BlockingRightTile      As Integer
   BlockingRightTile2     As Integer
   BlockingLeftTile       As Integer
   BlockingLeftTile2      As Integer
   CoveringTile           As Integer
   CompletelyCovered      As Boolean  ' if true, covering tile completely covers!
   TileBelow              As Integer  ' # of tile in row below if any
   Showing                As Boolean
   x                      As Single   'pos of tile in window
   y                      As Single
   Row                    As Integer  'pos in image of tiles
   Col                    As Integer
   Caption                As String
   Suite                  As String
   Taken                  As Boolean
End Type


Private Type udtTileRef
   TileID          As Integer
   Row             As Integer
   Col             As Integer
   Caption         As String
   Suite           As String
   Allocated       As Boolean
End Type

Private Tiles(143) As udtTile
Private pTile(143) As udtTileRef

Private wrkTiles()           As Integer

Private nPiTileReadWidth     As Integer
Private nPiTileReadHeight    As Integer

Private nPiTileWriteWidth    As Integer
Private nPiTileWriteHeight   As Integer

'Timing game
Private nPiSeconds                 As Integer
Private nPiMinutes                 As Integer
Private nPiTotalSeconds            As Integer


Private nPiTilesLeft               As Integer

Private Type udtTileCombo
    Tile1                As Integer
    Tile2                As Integer
End Type

Private Type udtTopScores
    TotalSeconds         As Integer
    Minutes              As Integer
    Seconds              As Integer
    Player               As String
End Type

Private uTopScores(1 To 10) As udtTopScores

Private uPiTileHistory()   As udtTileCombo
Private uPiMovesLeft()     As udtTileCombo

Private nPiActiveTile           As Integer
Private nPiFirstActiveTile      As Integer
Private nPiSecondActiveTile     As Integer

'Object References
'++++++++++++++++++++
Private picTiles                As PictureBox
Private picBrightTiles          As PictureBox
Private picDarkTiles            As PictureBox
Private picGame                 As PictureBox
Private picSel                  As PictureBox
Private picLabel                As PictureBox

Private frm                     As Form
Private lblTime                 As Label
Private lblTiles                As Label
Private lblTileID               As Label

Private nPiHintIndex            As Integer

Private bPiGameInProgress       As Boolean

Private bPiGameComplete         As Boolean
Private bPiGamePaused           As Boolean

Private nPiTopScoreCount        As Integer

'CUSTOM CURSOR?
Private bPiCustom               As Boolean

Event GameWasPaused()
Event GameWasResumed()
Event PiecesTaken()
Event NoMoves()
Event Winner()
Event NewGame()

'
'   CHECKS IF BOTH TILES SELECTED CAN BE TAKEN TOGETHER.
'   IF SO, IT REMOVES THEM AND ADDS THEM TO THE GAME'S TILE
'   HISTORY ARRAY, SO THAT IF THE USER WISHES TO, THEY CAN
'   UNDO ANY MOVES.
'
'
Public Function ToRemove(ByVal niFromTileID As Integer, ByVal niToTileID) As Boolean

    Dim n            As Integer
    
    If Tiles(niFromTileID).Suite <> Tiles(niToTileID).Suite Then
        
        Exit Function
    End If
   
    Select Case Tiles(niFromTileID).Suite
        Case "season", "harvest"
        Case Else
            If Tiles(niFromTileID).Caption <> Tiles(niToTileID).Caption Then
                Exit Function
            End If
    End Select
    
    
   Tiles(niFromTileID).Taken = True
   Tiles(niToTileID).Taken = True
   DisplayTiles
   
    nPiTilesLeft = nPiTilesLeft - 2
    lblTiles = nPiTilesLeft
    
    If UBound(uPiTileHistory) = 0 And uPiTileHistory(0).Tile1 = -1 Then
        uPiTileHistory(0).Tile1 = niFromTileID
        uPiTileHistory(0).Tile2 = niToTileID
    Else
        ReDim Preserve uPiTileHistory(UBound(uPiTileHistory) + 1)
        n = UBound(uPiTileHistory)
        uPiTileHistory(n).Tile1 = niFromTileID
        uPiTileHistory(n).Tile2 = niToTileID
    End If
    
    nPiHintIndex = -1
    
    BuildMovesLeftList
    
    ToRemove = True
    PlayWav App.Path & "\swish.wav"
    
    RaiseEvent PiecesTaken
    
End Function


'
'
' *** IF THE TILE FOR THE PASSED TILE ID IS BLOCKED BY OTHER
' *** TILES, IT WILL RETURN TRUE!
'
'
Public Function TileBlocked(ByVal niTileID As Integer) As Boolean
    Dim bLeftBlocked     As Boolean
    Dim bRightBlocked     As Boolean
    Dim bCovered          As Boolean
    
    bLeftBlocked = False
    bRightBlocked = False
    bCovered = False
    
    If Tiles(niTileID).BlockingLeftTile > -1 Then
        If Not Tiles(Tiles(niTileID).BlockingLeftTile).Taken Then
            bLeftBlocked = True
        End If
    End If
    
    
    If Tiles(niTileID).BlockingLeftTile2 > -1 Then
        If Not Tiles(Tiles(niTileID).BlockingLeftTile2).Taken Then
            bLeftBlocked = True
        End If
    End If
    
    If Tiles(niTileID).BlockingRightTile > -1 Then
        If Not Tiles(Tiles(niTileID).BlockingRightTile).Taken Then
            bRightBlocked = True
        End If
    End If
    
    
    If Tiles(niTileID).BlockingRightTile2 > -1 Then
        If Not Tiles(Tiles(niTileID).BlockingRightTile2).Taken Then
            bRightBlocked = True
        End If
    End If
    
    'Is that tile on the tippy-top of the stack still there?
    'It covers, but not completely, so, although the pieces
    'it covers are not blocked on the sides, and are somewhat
    'visible, they cannot be taken until that top tile is taken!
    If Tiles(niTileID).CoveringTile <> -1 Then
        If Not Tiles(Tiles(niTileID).CoveringTile).Taken Then
            bCovered = True
        End If
    End If
    
    TileBlocked = False
    
    If (bLeftBlocked And bRightBlocked) Or bCovered Then
        TileBlocked = True
    End If
    
End Function

'
' *** Check if both tiles match same "suite."
' *** If so, return TRUE
'
Public Function SuiteOK(ByVal niFromTileID As Integer, ByVal niToTileID) As Boolean
    If Tiles(niFromTileID).Suite <> Tiles(niToTileID).Suite Then Exit Function
    SuiteOK = True
End Function

'
'  'Takes our collection of 144 tiles, and shuffles them
'   really good!
'
'
Public Sub ShuffleTiles()
    Dim n             As Integer
    Dim i2            As Integer
    Dim i             As Integer
    Dim t             As Integer
    
    nPiFirstActiveTile = -1
    
    'DO A LITTLE MAGIC TO GET SOME MORE TRULLY RANDOM NUMBERS
    Randomize Second(Now)
    For n = 0 To Second(Now) + Minute(Now) + Hour(Now)
       i = Int(Rnd(1) * Second(Now)) + 1
    Next n
    
    
    
    For n = 0 To 143
        pTile(n).Allocated = False
    Next n
    
    i = -1
    
    Do
        i = i + 1
        
        If Not Tiles(i).Taken Then
        
            ReDim wrkTiles(0)
            
            'rebuild sort xref table:
            i2 = -1
            For n = 0 To 143
                If Not pTile(n).Allocated And Not Tiles(n).Taken Then
                    i2 = i2 + 1
                    ReDim Preserve wrkTiles(i2)
                    wrkTiles(i2) = n
                End If
            Next n
                        
            t = Int(Rnd(1) * (UBound(wrkTiles) + 1))
            Tiles(i).TileID = wrkTiles(t)
            Tiles(i).Row = pTile(wrkTiles(t)).Row
            Tiles(i).Col = pTile(wrkTiles(t)).Col
            Tiles(i).Suite = pTile(wrkTiles(t)).Suite
            Tiles(i).Caption = pTile(wrkTiles(t)).Caption
            pTile(wrkTiles(t)).Allocated = True
                        
        End If
        
    Loop While i < 143

End Sub

'
' *** BUILD OUR SUITES OF TILES
'
Private Sub SetTileRef()
    Dim n      As Integer
    Dim p      As Integer
    Dim x      As Integer
    Dim nX     As Integer
    Dim nY     As Integer
    
    p = -1
    
    For n = 0 To 26
        For x = 1 To 4
            p = p + 1
                            
            'find row & col on bitmap for this tile
            nY = Int(n / 9)
            nX = n - nY * 9
        
            'Exception (SHOULD NEVER HAPPEN)
            If nX > 8 Then Debug.Assert False
            If nY > 4 Then Debug.Assert False
            
            pTile(p).Row = nY
            pTile(p).Col = nX
            
            Select Case n
                Case 0, 9, 18
                    pTile(p).Caption = "1"
                Case 1, 10, 19
                    pTile(p).Caption = "2"
                Case 2, 11, 20
                    pTile(p).Caption = "3"
                Case 3, 12, 21
                    pTile(p).Caption = "4"
                Case 4, 13, 22
                    pTile(p).Caption = "5"
                Case 5, 14, 23
                    pTile(p).Caption = "6"
                Case 6, 15, 24
                    pTile(p).Caption = "7"
                Case 7, 16, 25
                    pTile(p).Caption = "8"
                Case 8, 17, 26
                    pTile(p).Caption = "9"
            End Select
            Select Case n
                Case 0 To 8
                    pTile(p).Suite = "flower"
                Case 9 To 17
                    pTile(p).Suite = "bamboo"
                Case 18 To 26
                    pTile(p).Suite = "character"
            End Select
        Next x
    Next n
    
    '4 Seasons
    For n = 0 To 3
        p = p + 1
        pTile(p).Row = 3
        pTile(p).Col = n
        pTile(p).Suite = "season"
        Select Case n
            Case 0
                pTile(p).Caption = "WIN"  ' Winter
            Case 1
                pTile(p).Caption = "SPR"  ' Spring
            Case 2
                pTile(p).Caption = "SUM"  ' Summer
            Case 3
                pTile(p).Caption = "AUT"  ' Autumn
        End Select
    Next n

   '4 harvests
    For n = 0 To 3
        p = p + 1
        pTile(p).Row = 4
        pTile(p).Col = n
        pTile(p).Suite = "harvest"
        Select Case n
            Case 0
                pTile(p).Caption = "BAM"  'Bamboo
            Case 1
                pTile(p).Caption = "PLUM"
            Case 2
                pTile(p).Caption = "MUM"
            Case 3
                pTile(p).Caption = "ORC"  'Orchid
        End Select
    Next n

    'NORTH
    For n = 0 To 3
        p = p + 1
        pTile(p).Row = 3
        pTile(p).Col = 4
        pTile(p).Caption = "N"
        pTile(p).Suite = "N"
    Next n
    
    'SOUTH
    For n = 0 To 3
        p = p + 1
        pTile(p).Row = 3
        pTile(p).Col = 5
        pTile(p).Caption = "S"
        pTile(p).Suite = "S"
    Next n

     'EAST
    For n = 0 To 3
        p = p + 1
        pTile(p).Row = 3
        pTile(p).Col = 6
        pTile(p).Caption = "E"
        pTile(p).Suite = "E"
    Next n
   
    'WEST
    For n = 0 To 3
        p = p + 1
        pTile(p).Row = 3
        pTile(p).Col = 7
        pTile(p).Caption = "W"
        pTile(p).Suite = "W"
    Next n
   
    'D1
    For n = 0 To 3
        p = p + 1
        pTile(p).Row = 4
        pTile(p).Col = 4
        pTile(p).Caption = "D1"
        pTile(p).Suite = "D1"
    Next n

    'D2
    For n = 0 To 3
        p = p + 1
        pTile(p).Row = 4
        pTile(p).Col = 5
        pTile(p).Caption = "D2"
        pTile(p).Suite = "D2"
    Next n
    
    'D3
    For n = 0 To 3
        p = p + 1
        pTile(p).Row = 4
        pTile(p).Col = 6
        pTile(p).Caption = "D3"
        pTile(p).Suite = "D3"
    Next n

End Sub


Public Property Let SecondActiveTile(ByVal vNewValue As Integer)
    nPiSecondActiveTile = vNewValue
    
End Property
Public Property Get SecondActiveTile() As Integer
    SecondActiveTile = nPiSecondActiveTile
    
End Property
Public Sub SaveTopScores(ByVal sWinnerName As String)
    If nPiTopScoreCount = 0 Then Exit Sub
    
    'save total # of top scores there are in the registry...
    'max is 10
    SaveSetting "vbMahjongg", "topscore", "total", Trim$(Str$(nPiTopScoreCount))
    
    Dim x                 As Integer
    
    'Find the slot for the winner name and plug it in!
    For x = 1 To 10
        If uTopScores(x).Player = "???" Then
            uTopScores(x).Player = sWinnerName
            Exit For
        End If
    Next x
    
    For x = 1 To nPiTopScoreCount
        SaveSetting "vbMahjongg", "topscore", _
           "totalsecs" & x, Trim$(Str$(uTopScores(x).TotalSeconds))
           
        SaveSetting "vbMahjongg", "topscore", _
           "min" & x, Trim$(Str$(uTopScores(x).Minutes))
        
        SaveSetting "vbMahjongg", "topscore", _
           "seconds" & x, Trim$(Str$(uTopScores(x).Seconds))
        
        SaveSetting "vbMahjongg", "topscore", _
           "player" & x, Trim$(uTopScores(x).Player)
        
    Next x
    
    DisplayTopScores
End Sub
'
' Used to play game sound effects
'
Public Sub PlayWav(ByVal siWavFile As String)
    Dim wFlags            As Long
    Dim rc                As Long
    
    wFlags = SND_ASYNC Or SND_NODEFAULT
    
    rc = sndPlaySound(siWavFile, wFlags)
    
End Sub
' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'
'
'  Use BitBlt API instead of PaintPicture method
'  because its faster!
'
Public Sub PaintPic(ByVal picSource, ByVal picDest, _
  ByVal fLeft As Single, ByVal fTop As Single, _
  ByVal lWidth As Long, ByVal lHeight As Long, _
  ByVal fLeftDest As Single, ByVal fTopDest As Single, _
  ByVal lDestWidth As Long, ByVal lDestHeight As Long, ByVal Index As Integer)
    
    Dim lResult     As Long
    Dim bFlag       As Boolean
    Dim bSkip       As Boolean
    
    bFlag = False   'flag to decide if we will draw edge of tile
    bSkip = False   'flag to decide if we will skip drawing the tile itself!
    
    If Tiles(Index).CompletelyCovered Then
        If Not Tiles(Tiles(Index).CoveringTile).Taken Then
            bSkip = True
        End If
    End If
    
    If Tiles(Index).TileBelow = -1 Then
        bFlag = True
    ElseIf Tiles(Tiles(Index).TileBelow).Taken Then
        bFlag = True
    End If
    
    If Tiles(Index).BlockingLeftTile = -1 Then
        bFlag = True
    ElseIf Tiles(Tiles(Index).BlockingLeftTile).Taken Then
        bFlag = True
    End If
    
    '*** ONLY DRAW TILE EDGE IF IT WOULD BE VISIBLE
    If bFlag Then
        'draw tile edge mask
        lResult = BitBlt(picDest.hDC, fLeftDest - 6, fTopDest, 43, 53, _
            frmMain.picTileEdgeMask.hDC, 0, 0, SRCAND)
                        
        'draw tile edge
        lResult = BitBlt(picDest.hDC, fLeftDest - 6, fTopDest, 43, 53, _
            frmMain.picTileEdge.hDC, 0, 0, SRCPAINT)
    End If
                    'w37 h47
                    
    '**** ONLY DRAW TILE FACE IF IT WOULD BE VISIBLE
    If Not bSkip Then
        'Draw the tile's face:
        lResult = BitBlt(picDest.hDC, fLeftDest, fTopDest, 37, 47, _
            picSource.hDC, fLeft, fTop, SRCCOPY)
    End If
    
End Sub


'
'
' *** INITIALIZE ALL VARIABLES NEEDED FOR A NEW GAME
'
' *** AND REFRESH DISPLAY WITH RESET VALUES!
'
'
Public Sub NewGame()
    Screen.MousePointer = vbHourglass
    
    ' bitmaps read in are in a grid of 9 x 5
    
    Dim n            As Integer
    Dim t            As Integer
    Dim bFound       As Boolean
    Dim i            As Integer
    Dim i2           As Integer
    Dim fX           As Single
    Dim fY           As Single
    
    
    ReDim uPiTileHistory(0)
    uPiTileHistory(0).Tile1 = -1
    uPiTileHistory(0).Tile2 = -1
    
    nPiFirstActiveTile = -1
    
    PlayWav App.Path & "\chimes.wav"
    
    'Build XRef table for which tiles point to which image
    SetTileRef
    
    'Reset tiles
    For n = 0 To 143
        Tiles(n).TileID = -1
        Tiles(n).Taken = False
        Tiles(n).BlockingLeftTile = -1
        Tiles(n).BlockingLeftTile2 = -1
        Tiles(n).BlockingRightTile = -1
        Tiles(n).BlockingRightTile2 = -1
        Tiles(n).CoveringTile = -1
        Tiles(n).CompletelyCovered = False
        Tiles(n).TileBelow = -1
        
    Next n
    
    'Shuffle tiles
    '==============
    ShuffleTiles
    i = -1
     
    
    With frmMain
        .picGame.Cls
        
        
        '**** INDIVIDUAL TILE DIMENSIONS
        nPiTileReadWidth = 37 '59  '.picTiles.ScaleWidth / 9
        nPiTileReadHeight = 50 '79 ' .picTiles.ScaleHeight / 5
        
        nPiTileWriteWidth = 37  'nPiTileReadWidth * 0.79
        nPiTileWriteHeight = 48  'nPiTileReadHeight * 0.79
                
        'LAYER 1 of TILES
        Tiles(0).x = 580
        Tiles(0).y = 220
       

        Tiles(1).x = Tiles(0).x - nPiTileWriteWidth
        Tiles(1).y = 220
        
        Tiles(2).x = Tiles(1).x - nPiTileWriteWidth
        Tiles(2).y = 220 - nPiTileWriteHeight / 2
       
        Tiles(3).x = Tiles(2).x
        Tiles(3).y = 220 + nPiTileWriteHeight / 2
        
        Tiles(4).x = Tiles(3).x
        Tiles(4).y = Tiles(3).y + nPiTileWriteHeight * 3

        ')==
        Tiles(5).x = Tiles(2).x
        Tiles(5).y = Tiles(2).y - nPiTileWriteHeight * 3

        Tiles(6).x = Tiles(5).x - nPiTileWriteWidth
        Tiles(6).y = Tiles(5).y
        
        Tiles(7).x = Tiles(6).x
        Tiles(7).y = Tiles(6).y + nPiTileWriteHeight * 2

        Tiles(8).x = Tiles(7).x
        Tiles(8).y = Tiles(7).y + nPiTileWriteHeight

        Tiles(9).x = Tiles(8).x
        Tiles(9).y = Tiles(8).y + nPiTileWriteHeight

        Tiles(10).x = Tiles(9).x
        Tiles(10).y = Tiles(9).y + nPiTileWriteHeight

        Tiles(11).x = Tiles(10).x
        Tiles(11).y = Tiles(10).y + nPiTileWriteHeight * 2

        '))
        Tiles(12).x = Tiles(6).x - nPiTileWriteWidth
        Tiles(12).y = Tiles(6).y

        For n = 1 To 7  '7
            Tiles(12 + n).x = Tiles(12).x
            Tiles(12 + n).y = Tiles(12 + n - 1).y + nPiTileWriteHeight
        Next n

        i = 19

        For n = 1 To 7
            For i2 = 1 To 8
                i = i + 1
                
                'If i > 68 Then Exit For
                Tiles(i).x = Tiles(i - 8).x - nPiTileWriteWidth
                Tiles(i).y = Tiles(i - 8).y
            Next i2
        Next n

        i = i + 1
        Tiles(i).x = Tiles(i - 8).x - nPiTileWriteWidth
        Tiles(i).y = Tiles(i - 8).y

        For n = 1 To 4   '4
            i = i + 1
            Tiles(i).x = Tiles(i - 7).x - nPiTileWriteWidth
            Tiles(i).y = Tiles(i - 7).y
        Next n

        i = i + 1
        Tiles(i).x = Tiles(i - 6).x - nPiTileWriteWidth
        Tiles(i).y = Tiles(i - 6).y

        i = i + 1
        Tiles(i).x = Tiles(i - 6).x - nPiTileWriteWidth
        Tiles(i).y = Tiles(i - 6).y

        i = i + 1
        Tiles(i).x = Tiles(i - 5).x - nPiTileWriteWidth
        Tiles(i).y = Tiles(i - 5).y

        i = i + 1
        Tiles(i).x = Tiles(i - 5).x - nPiTileWriteWidth
        Tiles(i).y = Tiles(i - 5).y

        i = i + 1
        Tiles(i).x = Tiles(i - 4).x - nPiTileWriteWidth
        Tiles(i).y = Tiles(i - 4).y

        Tiles(86).x = Tiles(84).x - nPiTileWriteWidth
        Tiles(86).y = Tiles(84).y - nPiTileWriteHeight / 2

        'NEXT LAYER OF TILES  (LAYER 2)
        Tiles(87).x = Tiles(21).x + 6
        Tiles(87).y = Tiles(21).y - 6

        For n = 1 To 5   '5
            Tiles(87 + n).x = Tiles(87).x
            Tiles(87 + n).y = Tiles(87 + n - 1).y + nPiTileWriteHeight
        Next n

        i = 87 + 5

        For n = 1 To 5
            For i2 = 1 To 6
                i = i + 1
                
                'If i > 93 Then Exit For
                Tiles(i).x = Tiles(i - 6).x - nPiTileWriteWidth
                Tiles(i).y = Tiles(i - 6).y
            Next i2
        Next n

        'NEXT LAYER OF TILES (LAYER 3)
        Tiles(123).x = Tiles(87 + 5 + 2).x + 6
        Tiles(123).y = Tiles(87 + 5 + 2).y - 6

        For n = 1 To 3
            Tiles(123 + n).x = Tiles(123).x
            Tiles(123 + n).y = Tiles(123 + n - 1).y + nPiTileWriteHeight
        Next n

        i = 123 + 3
        For n = 1 To 3
            For i2 = 1 To 4
                i = i + 1
                Tiles(i).x = Tiles(i - 4).x - nPiTileWriteWidth
                Tiles(i).y = Tiles(i - 4).y
            Next i2
        Next n

        'AND YET ANOTHER LAYER (LAYER 4)
        Tiles(139).x = Tiles(128).x + 6
        Tiles(139).y = Tiles(128).y - 6
        
        Tiles(140).x = Tiles(139).x
        Tiles(140).y = Tiles(139).y + nPiTileWriteHeight

        Tiles(141).x = Tiles(139).x - nPiTileWriteWidth
        Tiles(141).y = Tiles(139).y

        Tiles(142).x = Tiles(141).x
        Tiles(142).y = Tiles(140).y

        'AND THE ONE TILE TOP LAYER (LAYER 5)
        Tiles(143).x = Tiles(142).x + nPiTileWriteWidth / 2 + 6
        Tiles(143).y = Tiles(141).y + nPiTileWriteHeight / 2 - 6

        ' COULD THIS ROUTINE HAVE BEEN DONE IN LESS LINES OF CODE?
        '
        ' OF COURSE!  BUT I WANTED TO MAKE IT EASY TO READ AND
        ' UNDERSTAND.   ~ochomer
        
        '***** BLOCKING STUFF!
        Tiles(0).BlockingRightTile = 1
        Tiles(1).BlockingRightTile = 0
        Tiles(1).BlockingLeftTile = 2
        Tiles(1).BlockingLeftTile2 = 3
        Tiles(2).BlockingRightTile = 1
        Tiles(2).BlockingLeftTile = 8
        Tiles(2).TileBelow = 3
        Tiles(3).BlockingRightTile = 1
        Tiles(3).BlockingLeftTile = 9
        Tiles(4).BlockingLeftTile = 11
        Tiles(5).BlockingLeftTile = 6
        Tiles(6).BlockingRightTile = 5
        Tiles(6).BlockingLeftTile = 12
        Tiles(7).BlockingLeftTile = 14
        Tiles(7).TileBelow = 8
        Tiles(8).BlockingLeftTile = 15
        Tiles(8).BlockingRightTile = 2
        Tiles(8).TileBelow = 9
        Tiles(9).BlockingLeftTile = 16
        Tiles(9).BlockingRightTile = 3
        Tiles(9).TileBelow = 10
        Tiles(10).BlockingLeftTile = 17
        Tiles(11).BlockingLeftTile = 19
        Tiles(11).BlockingRightTile = 4
        Tiles(12).BlockingLeftTile = 20
        Tiles(12).BlockingRightTile = 6
        Tiles(12).TileBelow = 13
        Tiles(13).BlockingLeftTile = 21
        Tiles(13).TileBelow = 14
        Tiles(14).BlockingLeftTile = 22
        Tiles(14).BlockingRightTile = 7
        Tiles(14).TileBelow = 15
        Tiles(15).BlockingLeftTile = 23
        Tiles(15).BlockingRightTile = 8
        Tiles(15).TileBelow = 16
        Tiles(16).BlockingLeftTile = 24
        Tiles(16).BlockingRightTile = 9
        Tiles(16).TileBelow = 17
        Tiles(17).BlockingLeftTile = 25
        Tiles(17).BlockingRightTile = 10
        Tiles(17).TileBelow = 18
        Tiles(18).BlockingLeftTile = 26
        Tiles(18).TileBelow = 19
        Tiles(19).BlockingLeftTile = 27
        Tiles(19).BlockingRightTile = 11
        Tiles(20).BlockingLeftTile = 28
        Tiles(20).BlockingRightTile = 12
        Tiles(20).TileBelow = 21
        Tiles(21).BlockingLeftTile = 29
        Tiles(21).BlockingRightTile = 13
        Tiles(21).CompletelyCovered = True
        Tiles(21).CoveringTile = 87
        Tiles(21).TileBelow = 22
        Tiles(22).BlockingLeftTile = 30
        Tiles(22).BlockingRightTile = 14
        Tiles(22).CompletelyCovered = True
        Tiles(22).CoveringTile = 88
        Tiles(22).TileBelow = 23
        Tiles(23).BlockingLeftTile = 31
        Tiles(23).BlockingRightTile = 15
        Tiles(23).CompletelyCovered = True
        Tiles(23).CoveringTile = 89
        Tiles(23).TileBelow = 24
        Tiles(24).BlockingLeftTile = 32
        Tiles(24).BlockingRightTile = 16
        Tiles(24).CompletelyCovered = True
        Tiles(24).CoveringTile = 90
        Tiles(24).TileBelow = 25
        Tiles(25).BlockingLeftTile = 33
        Tiles(25).BlockingRightTile = 17
        Tiles(25).CompletelyCovered = True
        Tiles(25).CoveringTile = 91
        Tiles(25).TileBelow = 26
        Tiles(26).BlockingLeftTile = 34
        Tiles(26).BlockingRightTile = 18
        Tiles(26).CompletelyCovered = True
        Tiles(26).CoveringTile = 92
        Tiles(26).TileBelow = 27
        Tiles(27).BlockingLeftTile = 35
        Tiles(27).BlockingRightTile = 19
        Tiles(28).BlockingLeftTile = 36
        Tiles(28).BlockingRightTile = 20
        Tiles(28).TileBelow = 29
        Tiles(29).BlockingLeftTile = 37
        Tiles(29).BlockingRightTile = 21
        Tiles(29).CompletelyCovered = True
        Tiles(29).CoveringTile = 93
        Tiles(29).TileBelow = 30
        Tiles(30).BlockingLeftTile = 38
        Tiles(30).BlockingRightTile = 22
        Tiles(30).CompletelyCovered = True
        Tiles(30).CoveringTile = 94
        Tiles(30).TileBelow = 31
        Tiles(31).BlockingLeftTile = 39
        Tiles(31).BlockingRightTile = 23
        Tiles(31).CompletelyCovered = True
        Tiles(31).CoveringTile = 95
        Tiles(31).TileBelow = 32
        Tiles(32).BlockingLeftTile = 40
        Tiles(32).BlockingRightTile = 24
        Tiles(32).CompletelyCovered = True
        Tiles(32).CoveringTile = 96
        Tiles(32).TileBelow = 33
        Tiles(33).BlockingLeftTile = 41
        Tiles(33).BlockingRightTile = 25
        Tiles(33).CompletelyCovered = True
        Tiles(33).CoveringTile = 97
        Tiles(33).TileBelow = 34
        Tiles(34).BlockingLeftTile = 42
        Tiles(34).BlockingRightTile = 26
        Tiles(34).CompletelyCovered = True
        Tiles(34).CoveringTile = 98
        Tiles(34).TileBelow = 35
        Tiles(35).BlockingLeftTile = 43
        Tiles(35).BlockingRightTile = 27
        Tiles(36).BlockingLeftTile = 44
        Tiles(36).BlockingRightTile = 28
        Tiles(36).TileBelow = 37
        Tiles(37).BlockingLeftTile = 45
        Tiles(37).BlockingRightTile = 29
        Tiles(37).CompletelyCovered = True
        Tiles(37).CoveringTile = 99
        Tiles(37).TileBelow = 38
        Tiles(38).BlockingLeftTile = 46
        Tiles(38).BlockingRightTile = 30
        Tiles(38).CompletelyCovered = True
        Tiles(38).CoveringTile = 100
        Tiles(38).TileBelow = 39
        Tiles(39).BlockingLeftTile = 47
        Tiles(39).BlockingRightTile = 31
        Tiles(39).CompletelyCovered = True
        Tiles(39).CoveringTile = 101
        Tiles(39).TileBelow = 40
        Tiles(40).BlockingLeftTile = 48
        Tiles(40).BlockingRightTile = 32
        Tiles(40).CompletelyCovered = True
        Tiles(40).CoveringTile = 102
        Tiles(40).TileBelow = 41
        Tiles(41).BlockingLeftTile = 49
        Tiles(41).BlockingRightTile = 33
        Tiles(41).CompletelyCovered = True
        Tiles(41).CoveringTile = 103
        Tiles(41).TileBelow = 42
        Tiles(42).BlockingLeftTile = 50
        Tiles(42).BlockingRightTile = 34
        Tiles(42).CompletelyCovered = True
        Tiles(42).CoveringTile = 104
        Tiles(42).TileBelow = 43
        Tiles(43).BlockingLeftTile = 51
        Tiles(43).BlockingRightTile = 35
        Tiles(44).BlockingLeftTile = 52
        Tiles(44).BlockingRightTile = 36
        Tiles(44).TileBelow = 45
        Tiles(45).BlockingLeftTile = 53
        Tiles(45).BlockingRightTile = 37
        Tiles(45).CompletelyCovered = True
        Tiles(45).CoveringTile = 105
        Tiles(45).TileBelow = 46
        Tiles(46).BlockingLeftTile = 54
        Tiles(46).BlockingRightTile = 38
        Tiles(46).CompletelyCovered = True
        Tiles(46).CoveringTile = 106
        Tiles(46).TileBelow = 47
        Tiles(47).BlockingLeftTile = 55
        Tiles(47).BlockingRightTile = 39
        Tiles(47).CompletelyCovered = True
        Tiles(47).CoveringTile = 107
        Tiles(47).TileBelow = 48
        Tiles(48).BlockingLeftTile = 56
        Tiles(48).BlockingRightTile = 40
        Tiles(48).CompletelyCovered = True
        Tiles(48).CoveringTile = 108
        Tiles(48).TileBelow = 49
        Tiles(49).BlockingLeftTile = 57
        Tiles(49).BlockingRightTile = 41
        Tiles(49).CompletelyCovered = True
        Tiles(49).CoveringTile = 109
        Tiles(49).TileBelow = 50
        Tiles(50).BlockingLeftTile = 58
        Tiles(50).BlockingRightTile = 42
        Tiles(50).CompletelyCovered = True
        Tiles(50).CoveringTile = 110
        Tiles(50).TileBelow = 51
        Tiles(51).BlockingLeftTile = 59
        Tiles(51).BlockingRightTile = 43
        Tiles(52).BlockingLeftTile = 60
        Tiles(52).BlockingRightTile = 44
        Tiles(52).TileBelow = 53
        Tiles(53).BlockingLeftTile = 61
        Tiles(53).BlockingRightTile = 45
        Tiles(53).CompletelyCovered = True
        Tiles(53).CoveringTile = 111
        Tiles(53).TileBelow = 54
        Tiles(54).BlockingLeftTile = 62
        Tiles(54).BlockingRightTile = 46
        Tiles(54).CompletelyCovered = True
        Tiles(54).CoveringTile = 112
        Tiles(54).TileBelow = 55
        Tiles(55).BlockingLeftTile = 63
        Tiles(55).BlockingRightTile = 47
        Tiles(55).CompletelyCovered = True
        Tiles(55).CoveringTile = 113
        Tiles(55).TileBelow = 56
        Tiles(56).BlockingLeftTile = 64
        Tiles(56).BlockingRightTile = 48
        Tiles(56).CompletelyCovered = True
        Tiles(56).CoveringTile = 114
        Tiles(56).TileBelow = 57
        Tiles(57).BlockingLeftTile = 65
        Tiles(57).BlockingRightTile = 49
        Tiles(57).CompletelyCovered = True
        Tiles(57).CoveringTile = 115
        Tiles(57).TileBelow = 58
        Tiles(58).BlockingLeftTile = 66
        Tiles(58).BlockingRightTile = 50
        Tiles(58).CompletelyCovered = True
        Tiles(58).CoveringTile = 116
        Tiles(58).TileBelow = 59
        Tiles(59).BlockingLeftTile = 67
        Tiles(59).BlockingRightTile = 51
        Tiles(60).BlockingLeftTile = 68
        Tiles(60).BlockingRightTile = 52
        Tiles(60).TileBelow = 61
        Tiles(61).BlockingLeftTile = 69
        Tiles(61).BlockingRightTile = 53
        Tiles(61).CompletelyCovered = True
        Tiles(61).CoveringTile = 117
        Tiles(61).TileBelow = 62
        Tiles(62).BlockingLeftTile = 70
        Tiles(62).BlockingRightTile = 54
        Tiles(62).CompletelyCovered = True
        Tiles(62).CoveringTile = 118
        Tiles(62).TileBelow = 63
        Tiles(63).BlockingLeftTile = 71
        Tiles(63).BlockingRightTile = 55
        Tiles(63).CompletelyCovered = True
        Tiles(63).CoveringTile = 119
        Tiles(63).TileBelow = 64
        Tiles(64).BlockingLeftTile = 72
        Tiles(64).BlockingRightTile = 56
        Tiles(64).CompletelyCovered = True
        Tiles(64).CoveringTile = 120
        Tiles(64).TileBelow = 65
        Tiles(65).BlockingLeftTile = 73
        Tiles(65).BlockingRightTile = 57
        Tiles(65).CompletelyCovered = True
        Tiles(65).CoveringTile = 121
        Tiles(65).TileBelow = 66
        Tiles(66).BlockingLeftTile = 74
        Tiles(66).BlockingRightTile = 58
        Tiles(66).CompletelyCovered = True
        Tiles(66).CoveringTile = 122
        Tiles(66).TileBelow = 67
        Tiles(67).BlockingLeftTile = 75
        Tiles(67).BlockingRightTile = 59
        Tiles(68).BlockingLeftTile = 76
        Tiles(68).BlockingRightTile = 60
        Tiles(68).TileBelow = 69
        Tiles(69).BlockingRightTile = 61
        Tiles(69).TileBelow = 70
        Tiles(70).BlockingLeftTile = 77
        Tiles(70).BlockingRightTile = 62
        Tiles(70).TileBelow = 71
        Tiles(71).BlockingLeftTile = 78
        Tiles(71).BlockingRightTile = 63
        Tiles(71).TileBelow = 72
        Tiles(72).BlockingLeftTile = 79
        Tiles(72).BlockingRightTile = 64
        Tiles(72).TileBelow = 73
        Tiles(73).BlockingLeftTile = 80
        Tiles(73).BlockingRightTile = 65
        Tiles(73).TileBelow = 74
        Tiles(74).BlockingRightTile = 66
        Tiles(74).TileBelow = 75
        Tiles(75).BlockingLeftTile = 81
        Tiles(75).BlockingRightTile = 67
        Tiles(76).BlockingLeftTile = 82
        Tiles(76).BlockingRightTile = 68
        Tiles(77).BlockingRightTile = 70
        Tiles(77).TileBelow = 78
        Tiles(78).BlockingLeftTile = 83
        Tiles(78).BlockingRightTile = 71
        Tiles(78).TileBelow = 79
        Tiles(79).BlockingLeftTile = 84
        Tiles(79).BlockingRightTile = 72
        Tiles(79).TileBelow = 80
        Tiles(80).BlockingRightTile = 73
        Tiles(81).BlockingLeftTile = 85
        Tiles(81).BlockingRightTile = 75
        Tiles(82).BlockingRightTile = 76
        Tiles(83).BlockingLeftTile = 86
        Tiles(83).BlockingRightTile = 78
        Tiles(83).TileBelow = 84
        Tiles(84).BlockingLeftTile = 86
        Tiles(84).BlockingRightTile = 79
        Tiles(85).BlockingRightTile = 81
        Tiles(86).BlockingRightTile = 83
        Tiles(86).BlockingRightTile2 = 84
        Tiles(87).BlockingLeftTile = 93
        Tiles(87).TileBelow = 88
        Tiles(88).BlockingLeftTile = 94
        Tiles(88).TileBelow = 89
        Tiles(89).BlockingLeftTile = 95
        Tiles(89).TileBelow = 90
        Tiles(90).BlockingLeftTile = 96
        Tiles(90).TileBelow = 91
        Tiles(91).BlockingLeftTile = 97
        Tiles(91).TileBelow = 92
        Tiles(92).BlockingLeftTile = 98
        Tiles(93).BlockingLeftTile = 99
        Tiles(93).BlockingRightTile = 87
        Tiles(93).TileBelow = 94
        Tiles(94).BlockingLeftTile = 100
        Tiles(94).BlockingRightTile = 88
        Tiles(94).CompletelyCovered = True
        Tiles(94).CoveringTile = 123
        Tiles(94).TileBelow = 95
        Tiles(95).BlockingLeftTile = 101
        Tiles(95).BlockingRightTile = 89
        Tiles(95).CompletelyCovered = True
        Tiles(95).CoveringTile = 124
        Tiles(95).TileBelow = 96
        Tiles(96).BlockingLeftTile = 102
        Tiles(96).BlockingRightTile = 90
        Tiles(96).CompletelyCovered = True
        Tiles(96).CoveringTile = 125
        Tiles(96).TileBelow = 97
        Tiles(97).BlockingLeftTile = 103
        Tiles(97).BlockingRightTile = 91
        Tiles(97).CompletelyCovered = True
        Tiles(97).CoveringTile = 126
        Tiles(97).TileBelow = 98
        Tiles(98).BlockingLeftTile = 104
        Tiles(98).BlockingRightTile = 92
        Tiles(99).BlockingLeftTile = 105
        Tiles(99).BlockingRightTile = 93
        Tiles(99).TileBelow = 100
        Tiles(100).BlockingLeftTile = 106
        Tiles(100).BlockingRightTile = 94
        Tiles(100).CompletelyCovered = True
        Tiles(100).CoveringTile = 127
        Tiles(100).TileBelow = 101
        Tiles(101).BlockingLeftTile = 107
        Tiles(101).BlockingRightTile = 95
        Tiles(101).CompletelyCovered = True
        Tiles(101).CoveringTile = 128
        Tiles(101).TileBelow = 102
        Tiles(102).BlockingLeftTile = 108
        Tiles(102).BlockingRightTile = 96
        Tiles(102).CompletelyCovered = True
        Tiles(102).CoveringTile = 129
        Tiles(102).TileBelow = 103
        Tiles(103).BlockingLeftTile = 109
        Tiles(103).BlockingRightTile = 97
        Tiles(103).CompletelyCovered = True
        Tiles(103).CoveringTile = 130
        Tiles(103).TileBelow = 104
        Tiles(104).BlockingLeftTile = 110
        Tiles(104).BlockingRightTile = 98
        Tiles(105).BlockingLeftTile = 111
        Tiles(105).BlockingRightTile = 99
        Tiles(105).TileBelow = 106
        Tiles(106).BlockingLeftTile = 112
        Tiles(106).BlockingRightTile = 100
        Tiles(106).CompletelyCovered = True
        Tiles(106).CoveringTile = 131
        Tiles(106).TileBelow = 107
        Tiles(107).BlockingLeftTile = 113
        Tiles(107).BlockingRightTile = 101
        Tiles(107).CompletelyCovered = True
        Tiles(107).CoveringTile = 132
        Tiles(107).TileBelow = 108
        Tiles(108).BlockingLeftTile = 114
        Tiles(108).BlockingRightTile = 102
        Tiles(108).CompletelyCovered = True
        Tiles(108).CoveringTile = 133
        Tiles(108).TileBelow = 109
        Tiles(109).BlockingLeftTile = 115
        Tiles(109).BlockingRightTile = 103
        Tiles(109).CompletelyCovered = True
        Tiles(109).CoveringTile = 134
        Tiles(109).TileBelow = 110
        Tiles(110).BlockingLeftTile = 116
        Tiles(110).BlockingRightTile = 104
        Tiles(111).BlockingLeftTile = 117
        Tiles(111).BlockingRightTile = 105
        Tiles(111).TileBelow = 112
        Tiles(112).BlockingLeftTile = 118
        Tiles(112).BlockingRightTile = 106
        Tiles(112).CompletelyCovered = True
        Tiles(112).CoveringTile = 135
        Tiles(112).TileBelow = 113
        Tiles(113).BlockingLeftTile = 119
        Tiles(113).BlockingRightTile = 107
        Tiles(113).CompletelyCovered = True
        Tiles(113).CoveringTile = 136
        Tiles(113).TileBelow = 114
        Tiles(114).BlockingLeftTile = 120
        Tiles(114).BlockingRightTile = 108
        Tiles(114).CompletelyCovered = True
        Tiles(114).CoveringTile = 137
        Tiles(114).TileBelow = 115
        Tiles(115).BlockingLeftTile = 121
        Tiles(115).BlockingRightTile = 109
        Tiles(115).CompletelyCovered = True
        Tiles(115).CoveringTile = 138
        Tiles(115).TileBelow = 116
        Tiles(116).BlockingLeftTile = 122
        Tiles(116).BlockingRightTile = 110
        Tiles(117).BlockingRightTile = 111
        Tiles(117).TileBelow = 118
        Tiles(118).BlockingRightTile = 112
        Tiles(118).TileBelow = 119
        Tiles(119).BlockingRightTile = 113
        Tiles(119).TileBelow = 120
        Tiles(120).BlockingRightTile = 114
        Tiles(120).TileBelow = 121
        Tiles(121).BlockingRightTile = 115
        Tiles(121).TileBelow = 122
        Tiles(122).BlockingRightTile = 116
        Tiles(123).BlockingLeftTile = 127
        Tiles(123).TileBelow = 124
        Tiles(124).BlockingLeftTile = 128
        Tiles(124).TileBelow = 125
        Tiles(125).BlockingLeftTile = 129
        Tiles(125).TileBelow = 126
        Tiles(126).BlockingLeftTile = 130
        Tiles(127).BlockingLeftTile = 131
        Tiles(127).BlockingRightTile = 123
        Tiles(127).TileBelow = 128
        Tiles(128).BlockingLeftTile = 132
        Tiles(128).BlockingRightTile = 124
        Tiles(128).CompletelyCovered = True
        Tiles(128).CoveringTile = 139
        Tiles(128).TileBelow = 129
        Tiles(129).BlockingLeftTile = 133
        Tiles(129).BlockingRightTile = 125
        Tiles(129).CompletelyCovered = True
        Tiles(129).CoveringTile = 140
        Tiles(129).TileBelow = 130
        Tiles(130).BlockingLeftTile = 134
        Tiles(130).BlockingRightTile = 126
        Tiles(131).BlockingLeftTile = 135
        Tiles(131).BlockingRightTile = 127
        Tiles(131).TileBelow = 132
        Tiles(132).BlockingLeftTile = 136
        Tiles(132).BlockingRightTile = 128
        Tiles(132).CompletelyCovered = True
        Tiles(132).CoveringTile = 141
        Tiles(132).TileBelow = 133
        Tiles(133).BlockingLeftTile = 137
        Tiles(133).BlockingRightTile = 129
        Tiles(133).CompletelyCovered = True
        Tiles(133).CoveringTile = 142
        Tiles(133).TileBelow = 134
        Tiles(134).BlockingLeftTile = 138
        Tiles(134).BlockingRightTile = 130
        Tiles(135).BlockingRightTile = 131
        Tiles(135).TileBelow = 136
        Tiles(136).BlockingRightTile = 132
        Tiles(136).TileBelow = 137
        Tiles(137).BlockingRightTile = 133
        Tiles(137).TileBelow = 138
        Tiles(138).BlockingRightTile = 134
        
        '*** tiles 139-141 are covered but not
        '*** COMPLETELY covered!
        Tiles(139).BlockingLeftTile = 141
        Tiles(139).CoveringTile = 143
        Tiles(139).TileBelow = 140
        Tiles(140).BlockingLeftTile = 142
        Tiles(140).CoveringTile = 143
        Tiles(141).BlockingRightTile = 139
        Tiles(141).CoveringTile = 143
        Tiles(141).TileBelow = 142
        Tiles(142).BlockingRightTile = 140
        Tiles(142).CoveringTile = 143
        
        '*** THERE IS NOTHING FOR TILE 143
        '*** SINCE IT IS NOT BLOCKED OR COVERED
        '*** OR HAS ANY TILE PIECE BELOW IT.
        
        nPiHintIndex = -1
        
        'Reset total # of tiles untaken
        nPiTilesLeft = 144
        
        'Actually render the display of the tiles
        DisplayTiles
        
         'Reset game time
        nPiSeconds = 0
        nPiMinutes = 0
        nPiTotalSeconds = 0
        
        .timClock.Enabled = True
        
        
        
        lblTiles = nPiTilesLeft
        
    End With
    
    
    
    BuildMovesLeftList
    
    bPiGameInProgress = True
    bPiGameComplete = False
    
    If bPiGamePaused Then
        ResumeGame
    End If
    
    RaiseEvent NewGame
    
    Screen.MousePointer = vbDefault
End Sub


'
' *** THIS ROUTINE CAN BE USED FOR DEBUGGING!
'
'
Public Sub ListXRef(ByVal lst As ListBox)
    Dim n As Integer
    Dim s As String
    
    For n = 0 To 143
        s = "Tile# " & n & ":  Suite: "
        s = s & pTile(n).Suite
        s = s & "   Caption: " & pTile(n).Caption
        s = s & "   Row: " & pTile(n).Row
        s = s & "   Col: " & pTile(n).Col
        
        lst.AddItem s
        
    Next n
End Sub
Public Property Get FirstActiveTile() As Integer
    FirstActiveTile = nPiFirstActiveTile
    
End Property
Public Property Let FirstActiveTile(ByVal vNewValue As Integer)
    nPiFirstActiveTile = vNewValue
    
End Property
'
'  *** DRAW GREEN RECTANGLE OVER TILE THAT
'  *** THE MOUSE IS HOVERING OVER
'
Public Sub HighlightTile()

    Dim lResult       As Long
    Dim bFlag         As Boolean
    Dim n             As Integer
    
    n = nPiActiveTile
    
    With frm
    
        lResult = BitBlt(picSel.hDC, 0, 0, .ScaleWidth, _
            .ScaleHeight, _
            picGame.hDC, 0, 0, SRCCOPY)
        
        bFlag = False
        
        If n > -1 Then
            If Tiles(n).CompletelyCovered Then
                If Tiles(Tiles(n).CoveringTile).Taken Then
                    bFlag = True
                End If
            Else
                bFlag = True
            End If
            
            'Only highlight the edge of the tile if it is not
            'completely covered
            If bFlag And Not Tiles(n).Taken Then
                picSel.DrawWidth = 3
                picSel.Line (Tiles(n).x, Tiles(n).y)-(Tiles(n).x + _
                   nPiTileWriteWidth, Tiles(n).y + nPiTileWriteHeight) _
                   , vbGreen, B
                picSel.DrawWidth = 1
                
                If Not bPiCustom Then
                    frm.MousePointer = vbCustom
                    bPiCustom = True
                End If
            End If
            
        Else
            If bPiCustom Then
                frm.MousePointer = vbDefault
                bPiCustom = False
            End If
        End If
        
    
    
        'do double buffering
        lResult = BitBlt(frm.hDC, 0, 28, .ScaleWidth, _
            .ScaleHeight, _
            picSel.hDC, 0, 0, SRCCOPY)
        frm.Refresh

    End With
    
End Sub

'
'  *** DISPLAY HOW LONG THE CURRENT GAME HAS BEEN GOING!
'
'
Public Sub DisplayTime()
    Dim lResult          As Long
    
    nPiSeconds = nPiSeconds + 1
    nPiTotalSeconds = nPiTotalSeconds + 1
    
    If nPiSeconds > 59 Then
        nPiSeconds = 0
        nPiMinutes = nPiMinutes + 1
    End If
    
    lblTime.Caption = Format$(nPiMinutes, "00") & ":" _
        & Format$(nPiSeconds, "00")
    
   
    lblTime.Refresh
   
    
End Sub
Private Sub DisplayTiles(Optional ByVal vActiveTile)
    'this routine happens when:
    '1) a new game is started
    '2) after tiles are removed
    '3) when tiles are highlighted
    
    Dim fY             As Single
    Dim fX             As Single
    Dim n              As Integer
    Dim lResult        As Long
    
    Dim nActiveTile    As Integer
    
    Dim bFlag          As Boolean
    Dim bHint          As Boolean
    
    nActiveTile = -1
    
    If Not IsMissing(vActiveTile) Then
        nActiveTile = vActiveTile
    End If
    
    With frm
        .AutoRedraw = True   'SET JUST TO BE THOROUGH!
        .picGame.Font.Name = "arial"
        .picGame.Font.Size = 6
        .picGame.Font.Bold = True
        .picGame.Cls
        
        For n = 0 To 143
            If Not Tiles(n).Taken Then
                bFlag = False
                bHint = False
                
                fX = Tiles(n).Col * nPiTileReadWidth
                fY = Tiles(n).Row * nPiTileReadHeight
                
                '*** HIGHLIGHT "HINT" TILES
                If nPiHintIndex > -1 Then
                    If n = uPiMovesLeft(nPiHintIndex).Tile1 Then
                        PaintPic picBrightTiles, picGame, fX, fY, nPiTileReadWidth, _
                          nPiTileReadHeight, Tiles(n).x, Tiles(n).y, nPiTileWriteWidth, _
                          nPiTileWriteHeight, n
                          bFlag = True
                          bHint = True
                    End If
                    If n = uPiMovesLeft(nPiHintIndex).Tile2 Then
                        PaintPic picBrightTiles, picGame, fX, fY, nPiTileReadWidth, _
                          nPiTileReadHeight, Tiles(n).x, Tiles(n).y, nPiTileWriteWidth, _
                          nPiTileWriteHeight, n
                          bFlag = True
                          bHint = True
                    End If
                End If
                
                '*** OTHER TILE RENDERING OPTIONS
                
                '*** ANY SELECTED TILES
                If (n = nPiFirstActiveTile Or n = nPiSecondActiveTile) And Not bHint Then
                    PaintPic picDarkTiles, picGame, fX, fY, nPiTileReadWidth, _
                      nPiTileReadHeight, Tiles(n).x, Tiles(n).y, nPiTileWriteWidth, _
                      nPiTileWriteHeight, n
                    bFlag = True
                Else  'ANY TILE NOT OTHERWISE DESIGNATED
                    If Not bHint Then
                          PaintPic picTiles, picGame, fX, fY, nPiTileReadWidth, _
                        nPiTileReadHeight, Tiles(n).x, Tiles(n).y, nPiTileWriteWidth, _
                        nPiTileWriteHeight, n
                        bFlag = True
                    End If
                    
                End If
                
                          
                
                '*** HIGHLIGHT TILE MOUSE IS OVER WITH A GREEN RECTANGLE!
                If nActiveTile = n Then
                    .picGame.DrawWidth = 3
                    .picGame.Line (Tiles(n).x, Tiles(n).y)-(Tiles(n).x + _
                      nPiTileWriteWidth, Tiles(n).y + nPiTileWriteHeight) _
                      , vbGreen, B
                    .picGame.DrawWidth = 1
                End If
                
                'DRAW THE TILE LABELING (ONLY WHEN TILE IS VISIBLE):
                If bFlag Then
                    .picGame.CurrentX = Tiles(n).x + 2
                    .picGame.CurrentY = Tiles(n).y + 1
                    .picGame.Print Tiles(n).Caption
                End If
                
            End If
        Next n
        
        DisplayTime
        
        'do double buffering
        lResult = BitBlt(frm.hDC, 0, 28, .ScaleWidth, _
            .ScaleHeight, _
            picGame.hDC, 0, 0, SRCCOPY)
            
        DisplayTime
        
        frm.Refresh
        
    End With
    
End Sub
Public Function CurrentTile(ByVal x As Single, ByVal y As Single) As Integer
    '###############################################################
    '
    '  based on mouse pos, returns tile# that its hovering over,
    '  if none being hovered over, it returns a -1.
    '
    '###############################################################
    
    Dim n              As Integer
    
    '*** NOTICE HERE WE START WITH THE TILES ON TOP, AND WORK OUR WAY DOWN!
    For n = 143 To 0 Step -1
        If Not Tiles(n).Taken Then
            If x >= Tiles(n).x And x < Tiles(n).x + nPiTileWriteWidth Then
                If y >= Tiles(n).y And y < Tiles(n).y + nPiTileWriteHeight Then
                    CurrentTile = n
                    
                    Exit Function
                End If
            End If
            
        End If
    Next n
    
    CurrentTile = -1
    Exit Function
    
End Function

'
' ***** BACK UP ONE MOVE!
'
Public Sub Backup()
    Dim n               As Integer
    
    n = UBound(uPiTileHistory)
    
    If uPiTileHistory(n).Tile1 = -1 Then Exit Sub
    
    Tiles(uPiTileHistory(n).Tile1).Taken = False
    Tiles(uPiTileHistory(n).Tile2).Taken = False

    nPiTilesLeft = nPiTilesLeft + 2
    
    '*** DE-SELECT ANY TILES WHICH WHERE PREVIOUSLY SELECTED!
    nPiFirstActiveTile = -1
    nPiSecondActiveTile = -1
    
    '*** DISPLAY THE RESULTS OF OUR WORK!
    DisplayTiles
    
    'Show new count of tiles left to take:
    lblTiles = nPiTilesLeft
    
    If n > 0 Then
        ReDim Preserve uPiTileHistory(UBound(uPiTileHistory) - 1)
    Else
        uPiTileHistory(0).Tile1 = -1
        uPiTileHistory(0).Tile2 = -1
    End If
    
    BuildMovesLeftList
    
End Sub


Private Sub Class_Initialize()
    nPiFirstActiveTile = -1
    bPiGameInProgress = False
    
End Sub



Public Property Get ActiveTile() As Long
    ActiveTile = nPiActiveTile
    
End Property

Public Property Let ActiveTile(ByVal vNewValue As Long)
    nPiActiveTile = vNewValue
    
End Property

Public Sub ClickGame()
    'Was a tile selected by the player?
    
    ' (-1 means player clicked background not tile)
    
    If nPiActiveTile <> -1 Then
    
        '*** CANNOT MOVE TILE?
        If TileBlocked(nPiActiveTile) Then
            PlayWav App.Path & "\badmove.wav"
            Exit Sub
        End If

        
        
        'sound of clicking tiles
        PlayWav App.Path & "\click.wav"
        
        'Stop
        
        'HAS FIRST TILE BEEN SELECTED FOR A SEQUENCE???
        If nPiFirstActiveTile <> -1 Then
        
           'NO?
           
           'IS THE USER CLICKIN ON THE SAME TILE TO DE-SELECT IT?
           If nPiFirstActiveTile = nPiActiveTile Then
                nPiFirstActiveTile = -1
                nPiSecondActiveTile = -1
                DisplayTiles
                Exit Sub
           End If
        
           nPiSecondActiveTile = nPiActiveTile
           
           If Not ToRemove(nPiFirstActiveTile, nPiSecondActiveTile) Then
                nPiFirstActiveTile = nPiActiveTile
                nPiSecondActiveTile = -1
           Else
                nPiFirstActiveTile = -1
                nPiSecondActiveTile = -1
                
                '*** USER HAS WON GAME!
                If nPiTilesLeft = 0 Then
                    bPiGameComplete = True
                    RaiseEvent Winner
                    Exit Sub
                End If
                
                If MovesLeft = 0 Then
                    PlayWav App.Path & "\uhoh.wav"
                    RaiseEvent NoMoves
                End If
           End If
           
        Else   'NO TILES HAVE YET BEEN
           nPiFirstActiveTile = nPiActiveTile
           nPiSecondActiveTile = -1
        End If
        
        '** CALL THIS SO ANY TILE HIGHLIGHTING WILL SHOW UP
        '** FOR USER FEEDBACK:
        DisplayTiles
        HighlightTile
    End If

End Sub

'
'
'  *** SET OBJECT REFERENCES
'
'
'
Public Sub SetRefs(ByVal f As Form, _
    ByVal ipicTiles As PictureBox, _
    ByVal ipicBrightTiles As PictureBox, _
    ByVal ipicDarkTiles As PictureBox, _
    ByVal ipicSel As PictureBox, _
    ByVal ipicGame As PictureBox, _
    ByVal ilblTime As Label, _
    ByVal ilblTiles As Label)

    Set frm = f
    Set picTiles = ipicTiles
    Set picBrightTiles = ipicBrightTiles
    Set picDarkTiles = ipicDarkTiles
    Set picSel = ipicSel
    Set picGame = ipicGame
    
    
    Set lblTime = ilblTime
    Set lblTiles = ilblTiles
    
    LoadTopScores
    
End Sub

'
'
' * If game in progress Returns TRUE!
'
'
Public Property Get GameInProgress() As Boolean
    GameInProgress = bPiGameInProgress
    
End Property



Public Property Get GameComplete() As Boolean
    GameComplete = bPiGameComplete
    
End Property



Public Sub CleanUp()
    bPiGameComplete = True
    
End Sub

Public Sub PauseGame()
    bPiGamePaused = True
    RaiseEvent GameWasPaused
    
End Sub

Public Sub ResumeGame()
    bPiGamePaused = False
    RaiseEvent GameWasResumed
    
End Sub

Public Property Get GamePaused() As Boolean
    GamePaused = bPiGamePaused
    
End Property



'
' RETURNS HOW MANY POSSIBLE
'
'
Public Property Get MovesLeft() As Integer
    On Error GoTo TrapError
    
    If UBound(uPiMovesLeft) = 0 And uPiMovesLeft(0).Tile1 = -1 Then
        MovesLeft = 0
        Exit Property
    End If
    
    MovesLeft = UBound(uPiMovesLeft) + 1
    Exit Property
    
TrapError:
    MovesLeft = 0
    
End Property



'
'
' * BUILD AN ARRAY LIST OF VALID MOVES LEFT!
'
'
Private Sub BuildMovesLeftList()
    Dim n          As Integer
    Dim x          As Integer
    Dim n2         As Integer

    ReDim uPiMovesLeft(0)
    uPiMovesLeft(0).Tile1 = -1
    
    
    n = -1
    
    For x = 0 To UBound(Tiles)
        If Not Tiles(x).Taken Then
            If Not TileBlocked(x) Then
                For n2 = x + 1 To UBound(Tiles) - 1
                    If n2 <> x Then
                        If Not Tiles(n2).Taken Then
                            If Not TileBlocked(n2) Then
                                If Tiles(x).Suite = Tiles(n2).Suite Then
                                    Select Case Tiles(x).Suite
                                        Case "season", "harvest"
                                            n = n + 1
                                            ReDim Preserve uPiMovesLeft(n)
                                            uPiMovesLeft(n).Tile1 = x
                                            uPiMovesLeft(n).Tile2 = n2
                                        Case Else
                                            If Tiles(x).Caption = Tiles(n2).Caption Then
                                                n = n + 1
                                                ReDim Preserve uPiMovesLeft(n)
                                                uPiMovesLeft(n).Tile1 = x
                                                uPiMovesLeft(n).Tile2 = n2
                                            End If
                                    End Select
                                End If
                            End If
                        End If
                    End If
                Next n2
            End If
        End If
    Next x
    
End Sub

Public Sub ShowHint()
    
    'No moves left?
    If uPiMovesLeft(0).Tile1 = -1 Then
        MsgBox "No Moves Left.", vbInformation
        Exit Sub
    End If
    
    nPiHintIndex = nPiHintIndex + 1
    
    If nPiHintIndex > UBound(uPiMovesLeft) Then
        nPiHintIndex = 0
    End If
    
    'show tiles (highlight hint)
    DisplayTiles
    
End Sub

Private Sub LoadTopScores()

'this routine reads any top scores saved into the registry
'into an array.

    Dim n            As Integer
    Dim x            As Integer
    Dim sTest        As String
    
    
    'Initialize blank top score slots
    For n = 1 To 10
        uTopScores(n).TotalSeconds = 32767
        uTopScores(n).Minutes = -1
    Next n
    
    nPiTopScoreCount = 0
    
    sTest = GetSetting("vbMahjongg", "topscore", "total")
    
    If sTest = "" Then Exit Sub
    
    n = Val(sTest) 'total top scores saved in registry
    
    nPiTopScoreCount = n
    
    For x = 1 To n
        uTopScores(x).TotalSeconds = Val(GetSetting("vbMahjongg", _
           "topscore", "totalsecs" & x))
           
        uTopScores(x).Minutes = Val(GetSetting("vbMahjongg", _
           "topscore", "min" & x))
           
        uTopScores(x).Seconds = Val(GetSetting("vbMahjongg", _
           "topscore", "seconds" & x))
        
        uTopScores(x).Player = GetSetting("vbMahjongg", _
           "topscore", "player" & x)
        
    Next x
    
    DisplayTopScores
End Sub

Public Property Get TopScoreCount() As Integer
    TopScoreCount = nPiTopScoreCount
    
End Property



Public Function IsTopScore() As Boolean

'*** Find out if score is a top score,
'*** If it is, stick it where it belongs

    Dim n               As Integer
    Dim n2              As Integer
    
    IsTopScore = False
    
    For n = 1 To 10
        
        If uTopScores(n).Minutes = -1 Then
                uTopScores(n).TotalSeconds = nPiTotalSeconds
                uTopScores(n).Minutes = nPiMinutes
                uTopScores(n).Seconds = nPiSeconds
                uTopScores(n).Player = "???"
                nPiTopScoreCount = n
                IsTopScore = True
                
                Exit Function
        End If
        
        If nPiTotalSeconds < uTopScores(n).TotalSeconds Then
            
            If n = 10 Then
                uTopScores(n).TotalSeconds = nPiTotalSeconds
                uTopScores(n).Minutes = nPiMinutes
                uTopScores(n).Seconds = nPiSeconds
                uTopScores(n).Player = "???"
                nPiTopScoreCount = 10
                IsTopScore = True
                Exit Function
            End If
            
            'roll everything down
            For n2 = 9 To n Step -1
                uTopScores(n2 + 1).TotalSeconds = uTopScores(n2).TotalSeconds
                uTopScores(n2 + 1).Minutes = uTopScores(n2).Minutes
                uTopScores(n2 + 1).Seconds = uTopScores(n2).Seconds
                uTopScores(n2 + 1).Player = uTopScores(n2).Player
                
                If uTopScores(n2 + 1).Minutes > -1 And nPiTopScoreCount < n2 + 1 Then
                    nPiTopScoreCount = n2 + 1
                End If
            Next n2
                
            'Insert the new top score into the proper spot
            uTopScores(n).TotalSeconds = nPiTotalSeconds
            uTopScores(n).Minutes = nPiMinutes
            uTopScores(n).Seconds = nPiSeconds
            uTopScores(n).Player = "???"
            IsTopScore = True
            Exit Function
            
        End If
    Next n
    
End Function

Public Property Get Minutes() As Integer
    Minutes = nPiMinutes
    
End Property



Public Property Get Seconds() As Integer
    Seconds = nPiSeconds
    
End Property



Public Sub ClearCursor()
    'Get rid of green retangle on tile
    
    If nPiActiveTile = -1 Then Exit Sub
    
    nPiActiveTile = -1
    DisplayTiles
    
    If frm.MousePointer <> vbDefault Then
        frm.MousePointer = vbDefault
    End If
End Sub

Private Sub DisplayTopScores()
    Dim n               As Integer
    Dim n2              As Integer
    
    For n = 1 To nPiTopScoreCount
        n2 = n - 1
        frm.lblName(n2) = uTopScores(n).Player
        frm.lblName(n2).Visible = True
        frm.lblScore(n2) = Format$(uTopScores(n).Minutes, "00") & ":" & _
            Format$(uTopScores(n).Seconds, "00")
        frm.lblScore(n2).Visible = True
    Next n
    
    frm.Refresh
    
End Sub
