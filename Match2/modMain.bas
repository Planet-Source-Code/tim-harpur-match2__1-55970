Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public gameState As Long, gameSubState As Long, timeControl As Long, verifyQuit As Boolean
Public sneakPeak As Boolean, sneakCard As Long, lastSneakCard As Long
Public mouseButtonLState As Boolean, mouseButtonState As Boolean
Public lastKey As Long, lastKeyX As Long, musicPlaying As Long, cCardSound As Long

Public pointerList As cls2D_Object, pointerObject As cls2D_Object, transEnable As Boolean
Public gameDeck As cls3D_Object, introDecks As cls3D_Object, introTag As cls3D_Object, introCards As cls3D_Object
Public currentDeck As Long, rotateDirection As Long, introSpinMode(0 To 5) As Long, introSpinCount(0 To 5) As Long

Public startTime As Long, stopTime As Long, movesMade As Long, cardsLeft As Long, cardState(1 To 44) As Long
Public flippingCard(0 To 1) As Long, flippingMode(0 To 1) As Long, flippingState(0 To 1) As Long, maxCard As Long
Public matchWait As Long, matchMode As Long, cardList(1 To 2, 1 To 22) As Long, cardFlash As Long

Public Sub Main()
  'program start
  
  timeControl = DXDraw.DelayTillTime(0, 0, True) + 32
  
  Do While gameState >= 0
    timeControl = DXDraw.DelayTillTime(timeControl, 32) + 32
    
    CheckMainInput
    
    Select Case gameState
      Case 0 'load primary form for display
        frmMain.Show
      Case 1 'initialize the game display/sound/input
        InitializeGame
      Case 2 'show logicon splash
        LogiconSplash
      Case 3 'initialize game splash
        InitializeGameSplash
      Case 4 'display game splash
        DisplayGameSplash
      Case 5 'initialize new game
        InitGameScreen
      Case 6 'main game routine
        DisplayGameScreen
      Case 7 'game won screen
        DisplayGameWon
    End Select
    
    If gameState >= 4 Then
      DXDraw.RefreshDisplay True
      
      lastKey = lastKeyX
    End If
  Loop
  
  If Not (pointerList Is Nothing) Then pointerList.Destroy_Chain 'relaease all chains
  If Not (introDecks Is Nothing) Then introDecks.Destroy_Chain
  If Not (introCards Is Nothing) Then introCards.Destroy_Chain
  If Not (introTag Is Nothing) Then introTag.Destroy_Chain
  If Not (gameDeck Is Nothing) Then gameDeck.Destroy_Chain
  
  DXSound.CleanUp_DXSound
  DXInput.CleanUp_DXInput
  DXDraw.CleanUp_DXDraw
  
  gameState = -2
  
  Unload frmMain
End Sub

Public Sub CheckMainInput()
  If gameState >= 4 Then
    lastKeyX = -1
    
    DXInput.Poll_Keyboard
    
    If gameState = 7 Then
      With DXInput.dx_KeyboardState
        If .Key(DIK_SPACE) <> 0 Then
          If lastKey <> DIK_SPACE Then gameState = 3
          
          lastKeyX = DIK_SPACE
        ElseIf .Key(DIK_RETURN) <> 0 Then
          If lastKey <> DIK_RETURN Then gameState = 3
            
          lastKeyX = DIK_RETURN
        End If
      End With
    Else
      With DXInput.dx_KeyboardState
        If .Key(DIK_ESCAPE) <> 0 Then
          If lastKey <> DIK_ESCAPE Then
            If gameState = 4 Then
              gameState = -1
            ElseIf verifyQuit Then
              verifyQuit = False
            Else
              verifyQuit = True
            End If
          End If
          
          lastKeyX = DIK_ESCAPE
        ElseIf .Key(DIK_RETURN) <> 0 Then
          If lastKey <> DIK_RETURN Then
            If verifyQuit Then
              verifyQuit = False
              
              gameState = 3
            End If
          End If
          
          lastKeyX = DIK_RETURN
        ElseIf .Key(DIK_N) <> 0 Then
          If lastKey <> DIK_N Then
            If verifyQuit Then verifyQuit = False
          End If
          
          lastKeyX = DIK_N
        ElseIf .Key(DIK_Y) <> 0 Then
          If lastKey <> DIK_Y Then
            If verifyQuit Then
              verifyQuit = False
              
              gameState = 3
            End If
          End If
          
          lastKeyX = DIK_Y
        End If
      End With
    End If
  
    DXInput.Poll_Mouse
    
    If DXInput.dx_MouseState.buttons(0) <> 0 Then 'ensure that mouse only responds to single click
      If mouseButtonLState = False Then
        mouseButtonState = True
        mouseButtonLState = True
        
        If gameState = 7 Then gameState = 3
      Else
        mouseButtonState = False
      End If
    Else
      mouseButtonState = False
      mouseButtonLState = False
    End If
  End If
End Sub

Public Sub InitializeGame() 'gamestate 1
  If Command$() = "-s" Then
    DXDraw.Init_DXDrawScreen frmMain, , , , 29, True, 0
  Else
    'DXDraw.Init_DXDrawWindow frmMain, frmMain, 29, , 0
    DXDraw.Init_DXDrawScreen frmMain, , , , 29, , 0
    
    transEnable = True
  End If
  
  DXDraw.Set_D3DFilterMode transEnable
  DXDraw.Set_D3DCamera 0, 0, -250, 0, 0
  
  DXDraw.Set_SSurfacePath "\Graphics"
  
  DXInput.Init_DXInput frmMain
  DXInput.Acquire_Keyboard
  DXInput.Acquire_Mouse IM_ForegroundExclusive
  
  DXSound.Init_DXSound frmMain, 10, 4
  
  DXSound.Set_MusicPath "\Music"
  DXSound.Set_SoundPath "\Sound"
  
  DXDraw.Init_StaticSurface 1, "LogiconPlaque.jpg", 600, 198
  DXSound.Load_MusicFromMidi 1, "Splash.mid"
    
  gameState = 2
End Sub

Public Sub LogiconSplash() 'gamestate 2
  Dim tColour As Long
  
  If gameSubState = 30 Then DXSound.Play_Music 1
  
  gameSubState = gameSubState + 2
  
  If gameSubState > 255 Then
    If gameSubState >= 315 Then
      If gameSubState >= (315 + 255) Then
        tColour = 255
        
        DXDraw.Init_StaticSurface 1, "Notices.gif", 560, 342
        
        gameState = 3 'ready to advance to the game splash screen initialization
      Else
        tColour = gameSubState - 315
      End If
    ElseIf gameSubState = 256 Then
      'load all remaining music/sound/decks
      
      timeControl = DXDraw.DelayTillTime(0, 0, True) + 2000
      
      DXSound.Load_MusicFromMidi 2, "Intro.mid"
      
      DXSound.Load_SoundBuffer 1, "RotateDeck.wav"
      DXSound.Load_SoundBuffer 2, "RotateCard.wav"
      DXSound.Duplicate_SoundBuffer 3, 2
      DXSound.Duplicate_SoundBuffer 4, 2
      DXSound.Duplicate_SoundBuffer 5, 2
      DXSound.Duplicate_SoundBuffer 6, 2
      DXSound.Duplicate_SoundBuffer 7, 2
      DXSound.Load_SoundBuffer 8, "Match.wav"
      DXSound.Load_SoundBuffer 9, "Win.wav"
      DXSound.Load_SoundBuffer 10, "Button.wav"
      
      gameSubState = 314
      
      timeControl = DXDraw.DelayTillTime(timeControl, 0) + 32
    End If
  Else
    tColour = 255 - gameSubState
  End If
  
  DXDraw.Clear_Display
  If tColour <> 255 Then DXDraw.BlitSolid 1, 100, 199, 0, 0, 600, 198
  If transEnable Then DXDraw.Wash 0, 0, 0, tColour / 255
  DXDraw.RefreshDisplay True
End Sub

Public Sub InitializeGameSplash() 'gamestate 3
  Dim loop1 As Long
  
  DXSound.StopAll_Music
  DXSound.Play_Music 2
  musicPlaying = 1
  
  DXSound.StopAll_Sound
  
  DXDraw.Clear_Display
  
  For loop1 = 2 To 29 'release vid memory for card deck
    DXDraw.Release_StaticSurface loop1
  Next loop1
  
  DXDraw.Set_TexturePath "\Graphics"
  
  DXDraw.Init_TextureSurface 2, "FlagBack.gif", 128, 256
  DXDraw.Init_TextureSurface 3, "AnimalBack.gif", 128, 256
  DXDraw.Init_TextureSurface 4, "NumberBack.gif", 128, 256
  DXDraw.Init_TextureSurface 5, "TimeBack.gif", 128, 256
  DXDraw.Init_TextureSurface 6, "SymbolBack.gif", 128, 256
  DXDraw.Init_TextureSurface 7, "ShapeBack.gif", 128, 256
  DXDraw.Init_TextureSurface 8, "FishBack.gif", 128, 256
  DXDraw.Init_TextureSurface 9, "FlowerBack.gif", 128, 256
  DXDraw.Init_TextureSurface 10, "BirdBack.gif", 128, 256
  DXDraw.Init_TextureSurface 11, "DinosaurBack.gif", 128, 256
  DXDraw.Init_TextureSurface 12, "CardBase.gif", 64, 128
  DXDraw.Init_TextureSurface 13, "FlagTag.gif", 256, 64, -1
  DXDraw.Init_TextureSurface 14, "AnimalTag.gif", 256, 64, -1
  DXDraw.Init_TextureSurface 15, "NumberTag.gif", 256, 64, -1
  DXDraw.Init_TextureSurface 16, "TimeTag.gif", 256, 64, -1
  DXDraw.Init_TextureSurface 17, "SymbolTag.gif", 256, 64, -1
  DXDraw.Init_TextureSurface 18, "ShapeTag.gif", 256, 64, -1
  DXDraw.Init_TextureSurface 19, "FishTag.gif", 256, 64, -1
  DXDraw.Init_TextureSurface 20, "FlowerTag.gif", 256, 64, -1
  DXDraw.Init_TextureSurface 21, "BirdTag.gif", 256, 64, -1
  DXDraw.Init_TextureSurface 22, "DinosaurTag.gif", 256, 64, -1
  DXDraw.Init_TextureSurface 23, "CardBack.gif", 64, 128
  DXDraw.Init_TextureSurface 24, "M.gif", 64, 128
  DXDraw.Init_TextureSurface 25, "A.gif", 64, 128
  DXDraw.Init_TextureSurface 26, "T.gif", 64, 128
  DXDraw.Init_TextureSurface 27, "C.gif", 64, 128
  DXDraw.Init_TextureSurface 28, "H.gif", 64, 128
  DXDraw.Init_TextureSurface 29, "2.gif", 64, 128
  
  Init_PointerList
  Init_IntroDecks
  
  rotateDirection = 0
  
  gameState = 4
End Sub

Public Sub DisplayGameSplash() 'gamestate 4
  Select Case musicPlaying 'check to make sure the intro music is playing, and if not then restart it
    Case 0
      DXSound.Play_Music 2
      
      musicPlaying = 1
    Case 1
      If DXSound.IsMusicPlaying(2) = True Then musicPlaying = 2
    Case Else
      If DXSound.IsMusicPlaying(2) = False Then musicPlaying = 0
  End Select
  
  DXDraw.Clear_Display
  
  Process_PointerList
  Process_IntroDecks
  
  DXDraw.Begin_D3DScene
  introCards.ReDraw_Chain transEnable, transEnable
  introDecks.Render_Object transEnable, transEnable
  introTag.Render_Object , transEnable
  DXDraw.End_D3DScene
  
  pointerList.ReDraw_Chain 0, 0
End Sub

Public Sub InitGameScreen() 'gamestate 5
  Dim loop1 As Long, subTexture As String
  
  DXDraw.BlitTransparent 1, 285, 475, 254, 248, 230, 32
  
  DXDraw.RefreshDisplay True
  
  DXSound.Stop_Music 2
  
  musicPlaying = 0
  
  pointerList.Get_ChainObjectID(2).RemoveFrom_Chain pointerList 'remove left/right arrows from pointer list
  pointerList.Get_ChainObjectID(3).RemoveFrom_Chain pointerList
  pointerList.Get_ChainObjectID(4).RemoveFrom_Chain pointerList
  
  'load deck graphics
  For loop1 = 2 To 29 'release vid memory for intro graphics
    DXDraw.Release_StaticSurface loop1
  Next loop1
  
  Select Case currentDeck
    Case 0
      DXDraw.Set_TexturePath "\Graphics\FlagDeck"
      maxCard = 88
    Case 1
      DXDraw.Set_TexturePath "\Graphics\AnimalDeck"
      maxCard = 60
    Case 2
      DXDraw.Set_TexturePath "\Graphics\NumberDeck"
      maxCard = 48
    Case 3
      DXDraw.Set_TexturePath "\Graphics\TimeDeck"
      maxCard = 48
    Case 4
      DXDraw.Set_TexturePath "\Graphics\SymbolDeck"
      maxCard = 48
    Case 5
      DXDraw.Set_TexturePath "\Graphics\ShapeDeck"
      maxCard = 48
    Case 6
      DXDraw.Set_TexturePath "\Graphics\FishDeck"
      maxCard = 48
    Case 7
      DXDraw.Set_TexturePath "\Graphics\FlowerDeck"
      maxCard = 54
    Case 8
      DXDraw.Set_TexturePath "\Graphics\BirdDeck"
      maxCard = 48
    Case Else
      DXDraw.Set_TexturePath "\Graphics\DinosaurDeck"
      maxCard = 26
  End Select
  
  'initialize card deck
  Init_gameDeck
  
  movesMade = 0
  cardsLeft = 44
  matchMode = 0
  startTime = timeGetTime() \ 1000
  
  verifyQuit = False
  sneakPeak = False
  
  gameState = 6
  gameSubState = -59
End Sub

Public Sub DisplayGameScreen() 'gamestate 6
  Dim temp1 As Long, temp2 As Long, temp3 As Long, temp4 As Long, shiftPos As Long
  
  DXDraw.Clear_Display
  
  Process_PointerList
  Process_GameCards
  
  DXDraw.Begin_D3DScene
  gameDeck.ReDraw_Chain transEnable, transEnable
  DXDraw.End_D3DScene
  
  'redraw title and score
  DXDraw.BlitTransparent 1, 245, 0, 0, 143, 310, 54
  DXDraw.BlitTransparent 1, 132, 570, 0, 1, 490, 19
  
  temp1 = timeGetTime() \ 1000 - startTime
  
  temp2 = temp1 Mod 60 'seconds
  temp1 = (temp1 \ 60) Mod 1000 'minutes
  
  temp3 = temp1 \ 100
  temp1 = temp1 Mod 100
  temp4 = temp1 \ 10
  temp1 = temp1 Mod 10
  
  If temp3 <> 0 Then DXDraw.BlitTransparent 1, 273, 570, 13 * temp3, 21, 13, 19
  If temp3 <> 0 Or temp4 <> 0 Then DXDraw.BlitTransparent 1, 286, 570, 13 * temp4, 21, 13, 19
  DXDraw.BlitTransparent 1, 299, 570, 13 * temp1, 21, 13, 19
  
  temp3 = temp2 \ 10
  temp2 = temp2 Mod 10
  
  If temp3 <> 0 Then DXDraw.BlitTransparent 1, 376, 570, 13 * temp3, 21, 13, 19
  DXDraw.BlitTransparent 1, 389, 570, 13 * temp2, 21, 13, 19
  
  temp1 = movesMade Mod 1000 'moves made
  
  temp3 = temp1 \ 100
  temp1 = temp1 Mod 100
  temp4 = temp1 \ 10
  temp1 = temp1 Mod 10
  
  shiftPos = 0
  
  If temp3 <> 0 Then
    DXDraw.BlitTransparent 1, 631, 570, 13 * temp3, 21, 13, 19
  Else
    shiftPos = -13
  End If
  
  If temp3 <> 0 Or temp4 <> 0 Then
    DXDraw.BlitTransparent 1, 644 + shiftPos, 570, 13 * temp4, 21, 13, 19
  Else
    shiftPos = shiftPos - 13
  End If
  
  DXDraw.BlitTransparent 1, 657 + shiftPos, 570, 13 * temp1, 21, 13, 19
  
  If verifyQuit = False Then
    If matchMode = 1 Then
      DXDraw.BlitTransparent 1, 318, 480, 374, 53, 163, 74
    ElseIf matchMode = 2 Then
      DXDraw.BlitTransparent 1, 318, 480, 374, 135, 163, 67
    End If
  End If
  
  pointerList.ReDraw_Chain 0, 0
End Sub

Public Sub DisplayGameWon() 'gamestate 7
  Dim temp1 As Long, temp2 As Long, temp3 As Long, temp4 As Long, shiftPos As Long
  
  DXDraw.Clear_Display
  
  Process_PointerList
  Process_GameCardsWin
  
  DXDraw.Begin_D3DScene
  gameDeck.ReDraw_Chain transEnable, transEnable
  DXDraw.End_D3DScene
  
  'redraw title and score
  DXDraw.BlitTransparent 1, 245, 0, 0, 143, 310, 54
  DXDraw.BlitTransparent 1, 132, 570, 0, 1, 490, 19
  
  temp2 = stopTime Mod 60 'seconds
  temp1 = (stopTime \ 60) Mod 1000 'minutes
  
  temp3 = temp1 \ 100
  temp1 = temp1 Mod 100
  temp4 = temp1 \ 10
  temp1 = temp1 Mod 10
  
  If temp3 <> 0 Then DXDraw.BlitTransparent 1, 273, 570, 13 * temp3, 21, 13, 19
  If temp3 <> 0 Or temp4 <> 0 Then DXDraw.BlitTransparent 1, 286, 570, 13 * temp4, 21, 13, 19
  DXDraw.BlitTransparent 1, 299, 570, 13 * temp1, 21, 13, 19
  
  temp3 = temp2 \ 10
  temp2 = temp2 Mod 10
  
  If temp3 <> 0 Then DXDraw.BlitTransparent 1, 376, 570, 13 * temp3, 21, 13, 19
  DXDraw.BlitTransparent 1, 389, 570, 13 * temp2, 21, 13, 19
  
  temp1 = movesMade Mod 1000 'moves made
  
  temp3 = temp1 \ 100
  temp1 = temp1 Mod 100
  temp4 = temp1 \ 10
  temp1 = temp1 Mod 10
  
  shiftPos = 0
  
  If temp3 <> 0 Then
    DXDraw.BlitTransparent 1, 631, 570, 13 * temp3, 21, 13, 19
  Else
    shiftPos = -13
  End If
  
  If temp3 <> 0 Or temp4 <> 0 Then
    DXDraw.BlitTransparent 1, 644 + shiftPos, 570, 13 * temp4, 21, 13, 19
  Else
    shiftPos = shiftPos - 13
  End If
  
  DXDraw.BlitTransparent 1, 657 + shiftPos, 570, 13 * temp1, 21, 13, 19
  
  DXDraw.BlitTransparent 1, 220, 257, 0, 40, 359, 96
  
  pointerList.ReDraw_Chain 0, 0
End Sub

Public Sub flipCardSound()
  DXSound.Play_Sound cCardSound + 2
  
  cCardSound = cCardSound + 1
  If cCardSound > 5 Then cCardSound = 0
End Sub
