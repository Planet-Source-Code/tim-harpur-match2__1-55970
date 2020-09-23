Attribute VB_Name = "modObjects"
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Sub Init_IntroDecks()
  Dim tDeck As cls3D_Object, loop1 As Long
  
  If introDecks Is Nothing Then 'create main deck rack
    Set introDecks = New cls3D_Object
    
    Set tDeck = New cls3D_Object
    
    With tDeck
      .Create_Rod 10, 30, 30, 200, , , , , , , 0.8, 0.9, , , 12
      
      .PrimaryAxisRotation = Pi / 2
      .BasePosY = -80
      
      introDecks.Add_SurfacesFromTemplate tDeck, , True
    End With
    
    Set tDeck = New cls3D_Object
    
    With tDeck
      .Create_Rod 4, 10, 10, 240, , , , , , , 0.8, 0.9, , , 12
      .Double_ObjectComplexity
      
      .primaryAxis = 1
      
      introDecks.Add_SurfacesFromTemplate tDeck
      
      .PrimaryAxisRotation = Pi / 5
      introDecks.Add_SurfacesFromTemplate tDeck
      
      .PrimaryAxisRotation = Pi * 2 / 5
      introDecks.Add_SurfacesFromTemplate tDeck
      
      .PrimaryAxisRotation = Pi * 3 / 5
      introDecks.Add_SurfacesFromTemplate tDeck
      
      .PrimaryAxisRotation = Pi * 4 / 5
      introDecks.Add_SurfacesFromTemplate tDeck, , True
    End With
    
    introDecks.Join_AllSurfaces
    introDecks.Calculate_Normals True
    
    Set tDeck = New cls3D_Object
    
    With tDeck
      .Create_Box 50, 80, 15, , , , , 0.8, 0.9, , , 12
      .primaryAxis = 1
      
      With .Get_SurfaceID(6)
        .TextureSurfaceIndex = 2 'flags
        
        .SurfaceEmissiveRed = 1
        .SurfaceEmissiveGreen = 1
        .SurfaceEmissiveBlue = 1
      End With
      
      .BasePosX = Cos(Pi * 1.5) * 120
      .BasePosZ = Sin(Pi * 1.5) * 120
      .PrimaryAxisRotation = Pi * (1.5 - 1.5)
      
      introDecks.Add_SurfacesFromTemplate tDeck, 100
      
      .Get_SurfaceID(6).TextureSurfaceIndex = 3 'animals
      
      .BasePosX = Cos(Pi * 1.3) * 120
      .BasePosZ = Sin(Pi * 1.3) * 120
      .PrimaryAxisRotation = Pi * (1.3 - 1.5)
      
      introDecks.Add_SurfacesFromTemplate tDeck, 200
      
      .Get_SurfaceID(6).TextureSurfaceIndex = 4 'number
      
      .BasePosX = Cos(Pi * 1.1) * 120
      .BasePosZ = Sin(Pi * 1.1) * 120
      .PrimaryAxisRotation = Pi * (1.1 - 1.5)
      
      introDecks.Add_SurfacesFromTemplate tDeck, 300
      
      .Get_SurfaceID(6).TextureSurfaceIndex = 5 ' time
      
      .BasePosX = Cos(Pi * 0.9) * 120
      .BasePosZ = Sin(Pi * 0.9) * 120
      .PrimaryAxisRotation = Pi * (0.9 - 1.5)
      
      introDecks.Add_SurfacesFromTemplate tDeck, 400
      
      .Get_SurfaceID(6).TextureSurfaceIndex = 6 'symbols
      
      .BasePosX = Cos(Pi * 0.7) * 120
      .BasePosZ = Sin(Pi * 0.7) * 120
      .PrimaryAxisRotation = Pi * (0.7 - 1.5)
      
      introDecks.Add_SurfacesFromTemplate tDeck, 500
      
      .Get_SurfaceID(6).TextureSurfaceIndex = 7 ' shapes
      
      .BasePosX = Cos(Pi * 0.5) * 120
      .BasePosZ = Sin(Pi * 0.5) * 120
      .PrimaryAxisRotation = Pi * (0.5 - 1.5)
      
      introDecks.Add_SurfacesFromTemplate tDeck, 600
      
      .Get_SurfaceID(6).TextureSurfaceIndex = 8 ' fish
      
      .BasePosX = Cos(Pi * 0.3) * 120
      .BasePosZ = Sin(Pi * 0.3) * 120
      .PrimaryAxisRotation = Pi * (0.3 - 1.5)
      
      introDecks.Add_SurfacesFromTemplate tDeck, 700
      
      .Get_SurfaceID(6).TextureSurfaceIndex = 9 ' flower
      
      .BasePosX = Cos(Pi * 0.1) * 120
      .BasePosZ = Sin(Pi * 0.1) * 120
      .PrimaryAxisRotation = Pi * (0.1 - 1.5)
      
      introDecks.Add_SurfacesFromTemplate tDeck, 800
      
      .Get_SurfaceID(6).TextureSurfaceIndex = 10 ' bird
      
      .BasePosX = Cos(Pi * 1.9) * 120
      .BasePosZ = Sin(Pi * 1.9) * 120
      .PrimaryAxisRotation = Pi * (1.9 - 1.5)
      
      introDecks.Add_SurfacesFromTemplate tDeck, 900
      
      .Get_SurfaceID(6).TextureSurfaceIndex = 11 ' dinosaur
      
      .BasePosX = Cos(Pi * 1.7) * 120
      .BasePosZ = Sin(Pi * 1.7) * 120
      .PrimaryAxisRotation = Pi * (1.7 - 1.5)
      
      introDecks.Add_SurfacesFromTemplate tDeck, 1000, True
    End With
    
    With introDecks
      .primaryAxis = 1
      
      .Apply_VertexAmbient 0, 0, 0, 180, 1, 1, 1, 1
      
      .BasePosY = 20
      
      .Calculate_and_Update_Chain
    End With
    
    Set introTag = New cls3D_Object 'create deck tag
    
    With introTag
      .Create_Panel 64, 16, , , , 13, , True, 1#
      
      .Get_SurfaceID(2).Scale_Texture -1, -1
      
      Set tDeck = New cls3D_Object
      
      With tDeck
        .Create_Rod 6, 2, 2, 32, , , , , , , 0.8, 0.9, , , 12
        
        .primaryAxis = 1
        .PrimaryAxisRotation = Pi / 2
        
        .Calculate_Normals True
        
        .Apply_VertexAmbient 0, 0, -16, 32, 1, 1, 1, 1
        
        .BasePosX = -46
      End With
      
      .Add_SurfacesFromTemplate tDeck, 10
      
      With tDeck
        .PrimaryAxisRotation = -Pi / 2
        
        .BasePosX = 46
      End With
      
      .Add_SurfacesFromTemplate tDeck, 11, True
      
      .BasePosY = -45
      .BasePosZ = -140
      
      .Calculate_and_Update_Chain
    End With
    
    For loop1 = 0 To 5 'create title cards
      Set tDeck = New cls3D_Object
      
      With tDeck
        .Create_Panel 58, 110, , , , 23, 24 + loop1, True, 1#
        
        With .Get_SurfaceID(2)
          .Double_SurfaceComplexity
          
          .Scale_Texture -1, -1
          
          .SurfaceEmissiveRed = 0
          .SurfaceEmissiveGreen = 0
          .SurfaceEmissiveBlue = 0
          
          .Apply_VertexAmbient 0, 0, 0, 94, 1, 1, 1, 1
        End With
        
        .TypeID = loop1
        
        .BasePosX = -200 + 80 * loop1
        .BasePosY = 260
        .BasePosZ = 100
        
        .AddTo_Chain introCards
      End With
    Next loop1
  End If
  
  For loop1 = 0 To 5
    introSpinMode(loop1) = 0
    introSpinCount(loop1) = 10 + loop1 * 5
    
    introCards.Get_ChainObjectID(loop1).PrimaryAxisRotation = 0
  Next loop1
  
  introCards.Calculate_and_Update_Chain
End Sub

Public Sub Process_IntroDecks()
  Static rotateSubCount As Long, nextDeck As Long
  
  Dim loop1 As Long
  
  If rotateDirection > 0 Then
    rotateSubCount = rotateSubCount - 2
    
    With introDecks
      .PrimaryAxisRotation = (Pi / 5) * currentDeck + (Pi / 150) * (30 - rotateSubCount)
      
      .Update_ObjectPosition
    End With
    
    With introTag
      If currentDeck Mod 2 = 0 Then
        .PrimaryAxisRotation = (Pi / 30) * (30 - rotateSubCount)
      Else
        .PrimaryAxisRotation = (Pi / 30) * (30 - rotateSubCount) + Pi
      End If
      
      .Update_ObjectPosition
    End With
    
    If rotateSubCount <= 0 Then
      rotateDirection = 0
      
      currentDeck = nextDeck
    End If
  ElseIf rotateDirection < 0 Then
    rotateSubCount = rotateSubCount - 2
    
    With introDecks
      .PrimaryAxisRotation = (Pi / 5) * currentDeck - (Pi / 150) * (30 - rotateSubCount)
      
      .Update_ObjectPosition
    End With
    
    With introTag
      If currentDeck Mod 2 = 0 Then
        .PrimaryAxisRotation = -(Pi / 30) * (30 - rotateSubCount)
      Else
        .PrimaryAxisRotation = -(Pi / 30) * (30 - rotateSubCount) + Pi
      End If
      
      .Update_ObjectPosition
    End With
    
    If rotateSubCount <= 0 Then
      rotateDirection = 0
      
      currentDeck = nextDeck
    End If
  Else
    With DXInput.dx_KeyboardState
      If .Key(DIK_RIGHT) <> 0 Then
        If lastKey <> DIK_RIGHT Then
          rotateSubCount = 30
          rotateDirection = 1
          
          nextDeck = currentDeck + 1
          If nextDeck > 9 Then nextDeck = 0
        Else
          lastKeyX = DIK_RIGHT
        End If
      ElseIf .Key(DIK_LEFT) <> 0 Then
        If lastKey <> DIK_LEFT Then
          rotateSubCount = 30
          rotateDirection = -1
          
          nextDeck = currentDeck - 1
          If nextDeck < 0 Then nextDeck = 9
        Else
          lastKeyX = DIK_LEFT
        End If
      End If
    End With
    
    If rotateDirection <> 0 Then
      If currentDeck Mod 2 = 0 Then
        introTag.Get_SurfaceID(2).TextureSurfaceIndex = 13 + nextDeck
      Else
        introTag.Get_SurfaceID(1).TextureSurfaceIndex = 13 + nextDeck
      End If
      
      DXSound.Play_Sound 1
    End If
  End If
  
  For loop1 = 0 To 5
    With introCards.Get_ChainObjectID(loop1)
      Select Case introSpinMode(loop1)
        Case 0 'parked face down
          introSpinCount(loop1) = introSpinCount(loop1) - 1
          
          If introSpinCount(loop1) <= 0 Then
            introSpinMode(loop1) = 3
            
            flipCardSound
            
            introSpinCount(loop1) = 30
          End If
        Case 1 'parked face up
          introSpinCount(loop1) = introSpinCount(loop1) - 1
          
          If introSpinCount(loop1) <= 0 Then
            introSpinMode(loop1) = 2
            
            flipCardSound
            
            introSpinCount(loop1) = 30
          End If
        Case 2 'spin to face down
          introSpinCount(loop1) = introSpinCount(loop1) - 3
          
          If introSpinCount(loop1) <= 0 Then
            introSpinMode(loop1) = 0
            
            introSpinCount(loop1) = 120
            
            .PrimaryAxisRotation = 0
            
            .Update_ObjectPosition
          Else
            .PrimaryAxisRotation = (Pi / 30) * (30 - introSpinCount(loop1)) + Pi
            
            .Update_ObjectPosition
          End If
        Case 3 'spin to face up
          introSpinCount(loop1) = introSpinCount(loop1) - 3
          
          If introSpinCount(loop1) <= 0 Then
            introSpinMode(loop1) = 1
            
            introSpinCount(loop1) = 120
            
            .PrimaryAxisRotation = Pi
            
            .Update_ObjectPosition
          Else
            .PrimaryAxisRotation = (Pi / 30) * (30 - introSpinCount(loop1))
            
            .Update_ObjectPosition
          End If
      End Select
    End With
  Next loop1
End Sub

Public Sub Init_PointerList()
  Dim tObject As cls2D_Object, sObject As cls2D_ObjectSurface, loop1 As Long
  
  If pointerObject Is Nothing Then
    Set pointerObject = New cls2D_Object
    
    With pointerObject
      .objectName = "pointer"
      .TypeID = 0
      
      .Collision_WidthX_1000ths = 1000
      .Collision_WidthY_1000ths = 1000
      .Collision_Mask = -1
      
      .PosX_1000ths = 400000
      .PosY_1000ths = 300000
      
      Set sObject = New cls2D_ObjectSurface
      
      With sObject
        .SurfaceIndex = 1
        
        .SurfaceOffsetX = 482
        .SurfaceOffsetY = 208
        
        .ImageWidth = 53
        .ImageHeight = 35
        
        .ImagesPerRow = 1
      End With
      
      .Add_Surface sObject
      .AddTo_Chain pointerList
    End With
    
    Set tObject = New cls2D_Object
    
    With tObject
      .objectName = "quit"
      .TypeID = 1
      
      .Collision_WidthX_1000ths = 11000
      .Collision_WidthY_1000ths = 11000
      .Collision_Mask = -1
      
      .PosX_1000ths = 780000
      .PosY_1000ths = 20000
      
      .Calculate_CollsionBox
      
      Set sObject = New cls2D_ObjectSurface
      
      With sObject
        .SurfaceIndex = 1
        
        .SurfaceOffsetX = 372
        .SurfaceOffsetY = 26
        
        .DisplayOffsetX = -11
        .DisplayOffsetY = -11
        
        .ImageWidth = 23
        .ImageHeight = 22
        
        .ImagesPerRow = 1
      End With
      
      .Add_Surface sObject
      .AddTo_Chain pointerList
    End With
  End If
  
  Set tObject = New cls2D_Object
  
  With tObject
    .objectName = "left"
    .TypeID = 2
    
    .Collision_WidthX_1000ths = 19000
    .Collision_WidthY_1000ths = 19000
    .Collision_Mask = -1
    
    .PosX_1000ths = 300000
    .PosY_1000ths = 550000
    
    .Calculate_CollsionBox
    
    Set sObject = New cls2D_ObjectSurface
    
    With sObject
      .SurfaceIndex = 1
      
      .SurfaceOffsetX = 373
      .SurfaceOffsetY = 203
      
      .DisplayOffsetX = -19
      .DisplayOffsetY = -19
      
      .ImageWidth = 39
      .ImageHeight = 39
      
      .ImagesPerRow = 1
    End With
    
    .Add_Surface sObject
    .AddTo_Chain pointerList
  End With
  
  Set tObject = New cls2D_Object
  
  With tObject
    .objectName = "right"
    .TypeID = 3
    
    .Collision_WidthX_1000ths = 19000
    .Collision_WidthY_1000ths = 19000
    .Collision_Mask = -1
    
    .PosX_1000ths = 500000
    .PosY_1000ths = 550000
    
    .Calculate_CollsionBox
    
    Set sObject = New cls2D_ObjectSurface
    
    With sObject
      .SurfaceIndex = 1
      
      .SurfaceOffsetX = 438
      .SurfaceOffsetY = 203
      
      .DisplayOffsetX = -19
      .DisplayOffsetY = -19
      
      .ImageWidth = 39
      .ImageHeight = 39
      
      .ImagesPerRow = 1
    End With
    
    .Add_Surface sObject
    .AddTo_Chain pointerList
  End With
  
  Set tObject = New cls2D_Object
  
  With tObject
    .objectName = "start"
    .TypeID = 4
    
    .Collision_WidthX_1000ths = 62000
    .Collision_WidthY_1000ths = 22000
    .Collision_Mask = -1
    
    .PosX_1000ths = 400000
    .PosY_1000ths = 550000
    
    .Calculate_CollsionBox
    
    Set sObject = New cls2D_ObjectSurface
    
    With sObject
      .SurfaceIndex = 1
      
      .SurfaceOffsetX = 127
      .SurfaceOffsetY = 236
      
      .DisplayOffsetX = -62
      .DisplayOffsetY = -22
      
      .ImageWidth = 124
      .ImageHeight = 44
      
      .ImagesPerRow = 1
    End With
    
    .Add_Surface sObject
    .AddTo_Chain pointerList
  End With
  
  Set tObject = New cls2D_Object
  
  With tObject
    .objectName = "yes"
    .TypeID = 5
    
    .Collision_WidthX_1000ths = 32000
    .Collision_WidthY_1000ths = 15000
    .Collision_Mask = -1
    
    .PosX_1000ths = 550000
    .PosY_1000ths = 540000
    
    .Calculate_CollsionBox
    
    Set sObject = New cls2D_ObjectSurface
    
    With sObject
      .SurfaceIndex = 1
      
      .SurfaceOffsetX = 0
      .SurfaceOffsetY = 236
      
      .DisplayOffsetX = -32
      .DisplayOffsetY = -15
      
      .ImageWidth = 64
      .ImageHeight = 31
      
      .ImagesPerRow = 1
    End With
    
    .Add_Surface sObject
    .AddTo_Chain pointerList
  End With
  
  Set tObject = New cls2D_Object
  
  With tObject
    .objectName = "no"
    .TypeID = 6
    
    .Collision_WidthX_1000ths = 26000
    .Collision_WidthY_1000ths = 15000
    .Collision_Mask = -1
    
    .PosX_1000ths = 650000
    .PosY_1000ths = 540000
    
    .Calculate_CollsionBox
    
    Set sObject = New cls2D_ObjectSurface
    
    With sObject
      .SurfaceIndex = 1
      
      .SurfaceOffsetX = 74
      .SurfaceOffsetY = 236
      
      .DisplayOffsetX = -26
      .DisplayOffsetY = -15
      
      .ImageWidth = 52
      .ImageHeight = 31
      
      .ImagesPerRow = 1
    End With
    
    .Add_Surface sObject
    .AddTo_Chain pointerList
  End With
  
  Set tObject = New cls2D_Object
  
  With tObject
    .objectName = "verifyQuit"
    .TypeID = 7
    
    .PosX_1000ths = 100000
    .PosY_1000ths = 540000
    
    Set sObject = New cls2D_ObjectSurface
    
    With sObject
      .SurfaceIndex = 1
      
      .SurfaceOffsetX = 0
      .SurfaceOffsetY = 197
      
      .DisplayOffsetX = 0
      .DisplayOffsetY = -15
      
      .ImageWidth = 373
      .ImageHeight = 39
      
      .ImagesPerRow = 1
    End With
    
    .Add_Surface sObject
    .AddTo_Chain pointerList
  End With
  
  Set tObject = New cls2D_Object
  
  With tObject
    .objectName = "gameWon"
    .TypeID = 8
    
    .PosX_1000ths = 400000
    .PosY_1000ths = 300000
    
    Set sObject = New cls2D_ObjectSurface
    
    With sObject
      .SurfaceIndex = 1
      
      .SurfaceOffsetX = 0
      .SurfaceOffsetY = 40
      
      .DisplayOffsetX = -180
      .DisplayOffsetY = -48
      
      .ImageWidth = 359
      .ImageHeight = 96
      
      .ImagesPerRow = 1
    End With
    
    .Add_Surface sObject
    .AddTo_Chain pointerList
  End With
  
  Set tObject = New cls2D_Object
  
  With tObject
    .objectName = "sneakPeak"
    .TypeID = 9
    
    .Collision_WidthX_1000ths = 79000
    .Collision_WidthY_1000ths = 30000
    .Collision_Mask = -1
    
    .PosX_1000ths = 79000
    .PosY_1000ths = 30000
    
    .Calculate_CollsionBox
    
    Set sObject = New cls2D_ObjectSurface
    
    With sObject
      .SurfaceIndex = 1
      
      .SurfaceOffsetX = 0
      .SurfaceOffsetY = 281
      
      .DisplayOffsetX = -79
      .DisplayOffsetY = -30
      
      .ImageWidth = 158
      .ImageHeight = 61
      
      .ImagesPerRow = 2
      
      .Set_AnimationSequence Array(0, 1)
    End With
    
    .Add_Surface sObject
    .AddTo_Chain pointerList
  End With
  
  For loop1 = 1 To 44
    Set tObject = New cls2D_Object
    
    With tObject
      .objectName = "cardPos"
      .TypeID = 100 + loop1
      
      .PosX_1000ths = ((loop1 - 1) Mod 11) * 70000 + 47000
      .PosY_1000ths = ((loop1 - 1) \ 11) * 100000 + 130000
      
      .Collision_Mask = -1
      .CollisionBoxColour = &HFFFFFF
      .CollisionBoxVisible = True
      .Collision_WidthX_1000ths = 32000
      .Collision_WidthY_1000ths = 46000
      
      .Calculate_CollsionBox
      
      .AddTo_Chain pointerList
    End With
  Next loop1
End Sub

Public Sub Process_PointerList()
  Dim pObject As cls2D_Object, temp1 As Long
  
  With pointerObject ' do mouse pointer first
    .PosX_1000ths = .PosX_1000ths + DXInput.dx_MouseState.X * 1000
    .PosY_1000ths = .PosY_1000ths + DXInput.dx_MouseState.Y * 1000
    
    If .PosX_1000ths < 0 Then .PosX_1000ths = 0
    If .PosX_1000ths > 800000 Then .PosX_1000ths = 800000
    
    If .PosY_1000ths < 0 Then .PosY_1000ths = 0
    If .PosY_1000ths > 600000 Then .PosY_1000ths = 600000
    
    .Calculate_CollsionBox
  End With
  
  If gameState > 6 Then Exit Sub
  
  Set pObject = pointerList
  
  Do While Not (pObject Is Nothing)
    With pObject
      Select Case .TypeID
        Case 1 ' exit button
          If verifyQuit = False Then
            If mouseButtonState Then
              If .Test_Collision(pointerObject) Then
                If gameState = 4 Then
                  gameState = -1
                Else
                  verifyQuit = True
                End If
              End If
            End If
          End If
        Case 2 ' left arrow
          If DXInput.dx_MouseState.buttons(0) <> 0 Then
            If .Test_Collision(pointerObject) Then DXInput.dx_KeyboardState.Key(DIK_LEFT) = 1
          End If
        Case 3 ' right arrow
          If DXInput.dx_MouseState.buttons(0) <> 0 Then
            If .Test_Collision(pointerObject) Then DXInput.dx_KeyboardState.Key(DIK_RIGHT) = 1
          End If
        Case 4 ' start button
          If mouseButtonState Then
            If .Test_Collision(pointerObject) Then gameState = 5
          End If
        Case 5 'yes button
          If verifyQuit Then
            .Visible = True
            
            If mouseButtonState Then
              If .Test_Collision(pointerObject) Then
                verifyQuit = False
                
                gameState = 3
              End If
            End If
          Else
            .Visible = False
          End If
        Case 6 'no button
          If verifyQuit Then
            .Visible = True
            
            If mouseButtonState Then
              If .Test_Collision(pointerObject) Then verifyQuit = False
            End If
          Else
            .Visible = False
          End If
        Case 7 'verify quit panel
          If verifyQuit Then
            .Visible = True
          Else
            .Visible = False
          End If
        Case 8 'game won panel
          If gameState = 7 Then
            .Visible = True
          Else
            .Visible = False
          End If
        Case 9 'sneak peak
          If gameState = 6 Then
            .Visible = True
            
            If verifyQuit = False And gameSubState = 0 And sneakPeak = False Then
              .SurfaceList.ActionFrame = 0
              
              If mouseButtonState Then
                If .Test_Collision(pointerObject) Then
                  sneakPeak = True
                  
                  sneakCard = 100
                  
                  movesMade = movesMade + 10
                End If
              End If
            Else
              .SurfaceList.ActionFrame = 1
            End If
          Else
            .Visible = False
          End If
        Case Is > 100 'card
          If gameState = 6 Then
            If verifyQuit = False And sneakPeak = False Then
              If .Test_Collision(pointerObject) Then
                temp1 = .TypeID - 100
                
                If gameSubState >= 0 And gameSubState < 2 And cardState(temp1) = 1 Then
                  .Visible = True
                  
                  If mouseButtonState Then
                    cardState(temp1) = 2 'card has been selected
                    
                    flipCardSound
                    
                    flippingCard(gameSubState) = temp1
                    flippingMode(gameSubState) = 1
                    flippingState(gameSubState) = 0
                    
                    gameSubState = gameSubState + 1
                  End If
                Else
                  .Visible = False
                End If
              Else
                .Visible = False
              End If
            Else
              .Visible = False
            End If
          Else
            .Visible = False
          End If
      End Select
      
      Set pObject = .chainNext
    End With
  Loop
End Sub

Public Sub Init_gameDeck()
  Dim tDeck As cls3D_Object, loop1 As Long, loop2 As Long, temp1 As Long, temp2 As Long
  Dim cardValues(1 To 44) As Long, allCardValues() As Long
  
  ReDim allCardValues(1 To maxCard)
  
  DXDraw.Init_TextureSurface 29, "CardBack.gif", 64, 128
  
  Randomize
  
  flippingMode(0) = 0
  flippingMode(1) = 0
  
  For loop1 = 1 To 44
    cardState(loop1) = 1 ' initialize all cards as selectable
  Next loop1
  
  For loop1 = 1 To maxCard 'initialize all possible cards for deck
    allCardValues(loop1) = loop1
  Next loop1
  
  temp2 = maxCard
  
  For loop1 = 1 To 22 'load 22 cards for this game
    temp1 = Rnd() * temp2 + 1
    If temp1 > temp2 Then temp1 = temp2
    
    DXDraw.Init_TextureSurface 1 + loop1, "Card" & allCardValues(temp1) & ".gif", 64, 128
    
    temp2 = temp2 - 1
    
    For loop2 = temp1 To temp2 'shift remaining available cards down
      allCardValues(loop2) = allCardValues(loop2 + 1)
    Next loop2
    
    cardValues(loop1) = loop1 ' intialize cards in pairs from 1 to 22
    cardValues(loop1 + 22) = loop1
    
    cardList(1, loop1) = 0
    cardList(2, loop1) = 0
  Next loop1
  
  If Not (gameDeck Is Nothing) Then
    gameDeck.Destroy_Chain
    
    Set gameDeck = Nothing
  End If
  
  temp2 = 44
  
  'create game deck
  For loop1 = 1 To 44
    Set tDeck = New cls3D_Object
    
    With tDeck
      .TypeID = 100 + loop1
      
      temp1 = Rnd() * temp2 + 1 'select a card from the 22 loaded pairs
      If temp1 > temp2 Then temp1 = temp2
      
      If cardList(1, cardValues(temp1)) = 0 Then
        cardList(1, cardValues(temp1)) = .TypeID
      Else
        cardList(2, cardValues(temp1)) = .TypeID
      End If
      
      .Create_Panel 58, 110, , , , 29, 1 + cardValues(temp1), True, 1#
      
      .BasePosX = -336 + 66.6 * ((loop1 - 1) Mod 11)
      .BasePosY = 216 - 127 * ((loop1 - 1) \ 11)
      .BasePosZ = 130
      
      .PrimaryAxisRotation = Pi
      
      With .Get_SurfaceID(2)
        .Scale_Texture -1, -1
        .SurfaceOpacity = 0#
      End With
      
      temp2 = temp2 - 1
      
      For loop2 = temp1 To temp2 'shift remaining cards down
        cardValues(loop2) = cardValues(loop2 + 1)
      Next loop2
      
      .AddTo_Chain gameDeck
    End With
  Next loop1
  
  gameDeck.Calculate_and_Update_Chain
End Sub

Public Sub Process_GameCards()
  Dim tList As cls3D_Object, temp1 As Long
  
  If sneakPeak Then
    Do While sneakCard >= 100
      sneakCard = sneakCard + 1
      
      If sneakCard > 144 Then
        sneakCard = 0
        
        gameDeck.Get_ChainObjectID(lastSneakCard).User3 = 1
      ElseIf cardState(sneakCard - 100) = 1 Then
        lastSneakCard = sneakCard
        
        With gameDeck.Get_ChainObjectID(sneakCard) 'start next card flipping up
          .User1 = 100
          .User2 = 30
        End With
        
        flipCardSound
        
        Exit Do
      End If
    Loop
    
    Set tList = gameDeck
    
    Do While Not (tList Is Nothing)
      With tList
        Select Case .User1
          Case 100 'flipping up
            .User2 = .User2 - 3
            
            If .User2 = 0 Then
              .User1 = 101
              .User2 = 120
              
              .PrimaryAxisRotation = Pi
            Else
              .PrimaryAxisRotation = (Pi / 30) * (30 - .User2)
            End If
            
            .Update_ObjectPosition
          Case 101 'flipped up
            .User2 = .User2 - 1
            
            If .User2 = 0 Then
              .User1 = 102
              .User2 = 30
              
              flipCardSound
            End If
          Case 102 'flipping down
            .User2 = .User2 - 3
            
            If .User2 = 0 Then
              .User1 = 0
              .User2 = 0
              
              .PrimaryAxisRotation = 0
              
              If .User3 = 1 Then
                .User3 = 0
                
                sneakPeak = False
              End If
            Else
              .PrimaryAxisRotation = (Pi / 30) * (30 - .User2) + Pi
            End If
            
            .Update_ObjectPosition
        End Select
        
        Set tList = .chainNext
      End With
    Loop
  ElseIf verifyQuit = False Then
    If gameSubState < 0 Then
      gameSubState = gameSubState + 1
      
      Set tList = gameDeck
      
      Do While Not (tList Is Nothing)
        With tList
          temp1 = (159 - .TypeID) + gameSubState
          
          If temp1 >= 0 Then
            If temp1 < 15 Then
              If temp1 = 0 Then flipCardSound
              
              .PrimaryAxisRotation = (1 + temp1 / 15) * Pi
              
              .Update_ObjectPosition
              
              With .Get_SurfaceID(2)
                .SurfaceOpacity = temp1 / 15
                .Calculate_VerticesColour
              End With
            ElseIf temp1 = 15 Then
              .PrimaryAxisRotation = 0
              
              .Update_ObjectPosition
              
              With .Get_SurfaceID(2)
                .SurfaceOpacity = 1#
                .Calculate_VerticesColour
              End With
            End If
          End If
          
          Set tList = .chainNext
        End With
      Loop
    Else
      Select Case flippingMode(0)
        Case 1 'flipping up
          flippingState(0) = flippingState(0) + 3
          
          With gameDeck.Get_ChainObjectID(flippingCard(0) + 100)
            If flippingState(0) >= 30 Then
              .PrimaryAxisRotation = Pi
              
              flippingMode(0) = 2
            Else
              .PrimaryAxisRotation = (Pi / 30) * flippingState(0)
            End If
            
            .Update_ObjectPosition
          End With
        Case 2 'flipped up
          
        Case 3 'flipping down
          flippingState(0) = flippingState(0) + 3
          
          With gameDeck.Get_ChainObjectID(flippingCard(0) + 100)
            If flippingState(0) >= 30 Then
              .PrimaryAxisRotation = 0
              
              flippingMode(0) = 0
            Else
              .PrimaryAxisRotation = (Pi / 30) * flippingState(0) + Pi
            End If
            
            .Update_ObjectPosition
          End With
        Case 4 'fading out
          flippingState(0) = flippingState(0) + 3
          
          With gameDeck.Get_ChainObjectID(flippingCard(0) + 100)
            If flippingState(0) >= 30 Then
              With .Get_SurfaceID(2)
                .SurfaceOpacity = 0
                .Calculate_VerticesColour
              End With
              
              .Visible = False
              
              flippingMode(0) = 0
            Else
              With .Get_SurfaceID(2)
                .SurfaceOpacity = 1# / 30 * (30 - flippingState(0))
                .Calculate_VerticesColour
              End With
            End If
          End With
      End Select
      
      Select Case flippingMode(1)
        Case 1 'flipping up
          flippingState(1) = flippingState(1) + 3
          
          With gameDeck.Get_ChainObjectID(flippingCard(1) + 100)
            If flippingState(1) >= 30 Then
              .PrimaryAxisRotation = Pi
              
              flippingMode(1) = 2
              
              If .Get_SurfaceID(2).TextureSurfaceIndex = gameDeck.Get_ChainObjectID(flippingCard(0) + 100).Get_SurfaceID(2).TextureSurfaceIndex Then
                matchMode = 1
                cardsLeft = cardsLeft - 2
                
                DXSound.Play_Sound 8
              Else
                matchMode = 2
              End If
              
              movesMade = movesMade + 1
              matchWait = 30
            Else
              .PrimaryAxisRotation = (Pi / 30) * flippingState(1)
            End If
            
            .Update_ObjectPosition
          End With
        Case 2 'flipped up
          matchWait = matchWait - 1
          
          If matchWait <= 0 Then
            If matchMode = 1 Then
              flippingMode(0) = 4
              flippingMode(1) = 4
              
              cardState(flippingCard(0)) = 0
              cardState(flippingCard(1)) = 0
            Else
              flippingMode(0) = 3
              flippingMode(1) = 3
              
              cardState(flippingCard(0)) = 1
              cardState(flippingCard(1)) = 1
              
              flipCardSound
              flipCardSound
            End If
            
            flippingState(0) = 0
            flippingState(1) = 0
          End If
        Case 3 'flipping down
          flippingState(1) = flippingState(1) + 3
          
          With gameDeck.Get_ChainObjectID(flippingCard(1) + 100)
            If flippingState(1) >= 30 Then
              .PrimaryAxisRotation = 0
              
              flippingMode(1) = 0
              
              matchMode = 0
              gameSubState = 0
            Else
              .PrimaryAxisRotation = (Pi / 30) * flippingState(1) + Pi
            End If
            
            .Update_ObjectPosition
          End With
        Case 4 'fading out
          flippingState(1) = flippingState(1) + 3
          
          With gameDeck.Get_ChainObjectID(flippingCard(1) + 100)
            If flippingState(1) >= 30 Then
              With .Get_SurfaceID(2)
                .SurfaceOpacity = 0
                .Calculate_VerticesColour
              End With
              
              .Visible = False
              
              flippingMode(1) = 0
              
              matchMode = 0
              gameSubState = 0
              
              If cardsLeft = 0 Then
                gameState = 7
                
                stopTime = timeGetTime() \ 1000 - startTime
                
                DXSound.Play_Sound 9
              End If
            Else
              With .Get_SurfaceID(2)
                .SurfaceOpacity = 1# / 30 * (30 - flippingState(1))
                .Calculate_VerticesColour
              End With
            End If
          End With
      End Select
    End If
  End If
End Sub

Public Sub Process_GameCardsWin()
  Dim tObject As cls3D_Object
  
  Select Case flippingMode(0)
    Case 0 'pick card
      cardFlash = cardFlash + 1
      If cardFlash > 22 Then cardFlash = 1
      
      flippingCard(0) = cardList(1, cardFlash)
      flippingCard(1) = cardList(2, cardFlash)
      
      flippingState(0) = 0
      flippingMode(0) = 1
    Case 1 'fading in
      flippingState(0) = flippingState(0) + 1
      
      With gameDeck.Get_ChainObjectID(flippingCard(0))
        If flippingState(0) >= 30 Then
          With .Get_SurfaceID(2)
            .SurfaceOpacity = 1#
            .Calculate_VerticesColour
          End With
        Else
          With .Get_SurfaceID(2)
            .SurfaceOpacity = 1# / 30 * flippingState(0)
            .Calculate_VerticesColour
          End With
          
          .Visible = True
        End If
      End With
      
      With gameDeck.Get_ChainObjectID(flippingCard(1))
        If flippingState(0) >= 30 Then
          With .Get_SurfaceID(2)
            .SurfaceOpacity = 1#
            .Calculate_VerticesColour
          End With
          
          flippingMode(0) = 2
          flippingState(0) = 0
        Else
          With .Get_SurfaceID(2)
            .SurfaceOpacity = 1# / 30 * flippingState(0)
            .Calculate_VerticesColour
          End With
          
          .Visible = True
        End If
      End With
    Case 2 'faded in
      flippingState(0) = flippingState(0) + 1
      
      If flippingState(0) >= 30 Then
        flippingMode(0) = 3
        flippingState(0) = 0
      End If
    Case 3 'fading out
      flippingState(0) = flippingState(0) + 1
      
      With gameDeck.Get_ChainObjectID(flippingCard(0))
        If flippingState(0) >= 30 Then
          With .Get_SurfaceID(2)
            .SurfaceOpacity = 0
            .Calculate_VerticesColour
          End With
          
          .Visible = False
        Else
          With .Get_SurfaceID(2)
            .SurfaceOpacity = 1# / 30 * (30 - flippingState(0))
            .Calculate_VerticesColour
          End With
        End If
      End With
      
      With gameDeck.Get_ChainObjectID(flippingCard(1))
        If flippingState(0) >= 30 Then
          With .Get_SurfaceID(2)
            .SurfaceOpacity = 0
            .Calculate_VerticesColour
          End With
          
          .Visible = False
          
          flippingMode(0) = 0
        Else
          With .Get_SurfaceID(2)
            .SurfaceOpacity = 1# / 30 * (30 - flippingState(0))
            .Calculate_VerticesColour
          End With
        End If
      End With
  End Select
End Sub
