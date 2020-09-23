Attribute VB_Name = "modSubs"
' DirectX8 1st Person 3D Game VERSION 1.2 (19/03/2002)
' By Frederico Machado (indiofu@bol.com.br)
' Please vote for me if you like the game.
'
' ************************************************** '
' Just some Subs that I wrote all by myself          '
' I've created the Map loader and the map style.     '
' You can use these subs in your own games           '
' but don't forget to give me some credit.           '
' Just send me an e-mail, just to know, you know...  '
' Sorry my English, I'm Brazilian! :)                '
' ************************************************** '
'
' ************************************************** '
' I need help to fix the sky, and it needs to look   '
' like the sky of TrueVision, but I don't know how   '
' to do that. If anyone knows, please help me.       '
' I've set the MAGFILTER to D3DTEXF_POINT (see in    '
' the RenderSky() sub in modLandscape) to fix the    '
' "mix" problem of the sky textures but the textures '
' look horrible. Try to help me please!              '
' I've added support to load heightmaps to create    '
' custom landscapes, but I can't load a heightmap    '
' with more than 180x180 pixels, cause it let the    '
' framerate too down, even on my 800mhz 32MB VGA     '
' If you know how to fix it, please help me.         '
' We can't walk correctly through the terrain, it is '
' a problem. (The camera does litle jumps)           '
' ************************************************** '

Option Explicit

Global Path As String
Global TexturePath As String ' Our texture directory
Global MeshPath As String ' Our mesh directory
Global ObjectPath As String
Global DataPath As String

' Loads textures of walls and roof
Public Sub LoadTextures(strTextures As String)

  Dim strLine As String

  Open strTextures For Input As #1
    
    Do
    
      Input #1, strLine
      If Left$(strLine, 1) = ";" Or strLine = "" Or LCase(Right$(strLine, 4)) <> ".bmp" Then GoTo Jump  ' If the first char is a ; then it is a comment
      
      TexCount = TexCount + 1
      
      ReDim Preserve Textures(TexCount - 1) As Direct3DTexture8
      
      Set Textures(TexCount - 1) = D3DX.CreateTextureFromFile(D3DDevice, TexturePath & strLine)
      
Jump:
      
    Loop Until EOF(1)
    
  Close #1

End Sub

' Loads tree textures
Public Sub LoadTreeTextures(strTextures As String)

  Dim strLine As String

  Open strTextures For Input As #1
    
    Do
    
      Input #1, strLine
      If Left$(strLine, 1) = ";" Or strLine = "" Then GoTo Jump  ' If the first char is a ; then it is a comment
      
      TreeTexCount = TreeTexCount + 1
      
      ReDim Preserve TreeTextures(TreeTexCount - 1) As Direct3DTexture8
      
      Set TreeTextures(TreeTexCount - 1) = D3DX.CreateTextureFromFileEx(D3DDevice, TexturePath & strLine, 256, 256, D3DX_DEFAULT, 0, D3DFMT_A1R5G5B5, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, &HFF000000, ByVal 0, ByVal 0)
      
Jump:
      
    Loop Until EOF(1)
    
  Close #1

End Sub

' Loads sky textures
Public Sub LoadSkyTextures(strTextures As String)

  Dim numTextures As Integer
  Dim i As Integer
  Dim strLine As String

  Open strTextures For Input As #1
    
    Input #1, numTextures
    
    For i = 0 To numTextures - 1
      Input #1, strLine
      Set SkyTextures(i) = D3DX.CreateTextureFromFile(D3DDevice, TexturePath & strLine)
    Next
    
  Close #1

End Sub

' It adds a Floor or a Roof to the world
' change the Y value to add floors or roofs. Easy!
Public Sub AddFloorRoof(Width As Single, Height As Single, X As Single, Y As Single, Z As Single, Texture As Integer, FU As Single, FV As Single)

  Dim v(0 To 3) As CUSTOMVERTEX
  Dim VertexSizeInBytes As Long
  
  VBCount = VBCount + 1
  
  VertexSizeInBytes = Len(v(0))

  v(0).position = MakeVector(X + Width, Y, Z) ' Creates the floor or roof
  v(1).position = MakeVector(X, Y, Z)
  v(2).position = MakeVector(X + Width, Y, Z + Height)
  v(3).position = MakeVector(X, Y, Z + Height)
  v(0).tu = 0: v(0).tv = 0 ' Lets set the texture coordinates
  v(1).tu = FU: v(1).tv = 0
  v(2).tu = 0: v(2).tv = FV
  v(3).tu = FU: v(3).tv = FV
  
  CalculateNormals v
  
  ReDim Preserve VBuffers(VBCount - 1) As Direct3DVertexBuffer8
  
  ' Create the vertex buffer.
  Set VBuffers(VBCount - 1) = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
       0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)

  ' fill the vertex buffer from our array
  D3DVertexBuffer8SetData VBuffers(VBCount - 1), 0, VertexSizeInBytes * 4, 0, v(0)

  ReDim Preserve VBTex(VBCount - 1) As Integer
  VBTex(VBCount - 1) = Texture

End Sub

' Adds a wall to the world
' WType: 0 if it is a front or a back wall
'        1 if it is a left or a right wall
Public Sub Addwall(WType As Integer, Width As Single, Height As Single, X As Single, Y As Single, Z As Single, Texture As Integer, FU As Single, FV As Single)

  Dim v(0 To 3) As CUSTOMVERTEX
  Dim VertexSizeInBytes As Long
  
  VBCount = VBCount + 1
  
  VertexSizeInBytes = Len(v(0))

  If WType = 0 Then
    v(0).position = MakeVector(X, Y + Height, Z)
    v(1).position = MakeVector(X + Width, Y + Height, Z)
    v(2).position = MakeVector(X, Y, Z)
    v(3).position = MakeVector(X + Width, Y, Z)
  Else
    v(0).position = MakeVector(X, Y + Height, Z)
    v(1).position = MakeVector(X, Y + Height, Z + Width)
    v(2).position = MakeVector(X, Y, Z)
    v(3).position = MakeVector(X, Y, Z + Width)
  End If
  v(0).normal = GetPolygonNormal(v(0).position, v(1).position, v(2).position)
  v(1).normal = GetPolygonNormal(v(0).position, v(1).position, v(2).position)
  v(2).normal = GetPolygonNormal(v(0).position, v(1).position, v(2).position)
  v(3).normal = GetPolygonNormal(v(1).position, v(2).position, v(3).position)
  
  v(0).tu = 0: v(0).tv = 0
  v(1).tu = FU: v(1).tv = 0
  v(2).tu = 0: v(2).tv = FV
  v(3).tu = FU: v(3).tv = FV
  
  CalculateNormals v
  
  ReDim Preserve VBuffers(VBCount - 1) As Direct3DVertexBuffer8
  
  ' Create the vertex buffer.
  Set VBuffers(VBCount - 1) = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
       0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)

  ' fill the vertex buffer from our array
  D3DVertexBuffer8SetData VBuffers(VBCount - 1), 0, VertexSizeInBytes * 4, 0, v(0)

  ReDim Preserve VBTex(VBCount - 1) As Integer
  VBTex(VBCount - 1) = Texture

End Sub

' Adds a light to the world
' HELP HELP HELP HELP HELP
' I can't understand lights
' I want to add lights like real lamps
' in a room. just like Half Life lights.
' Help me to understand lights, please.
Public Sub AddLight(X As Single, Y As Single, Z As Single, LType As Integer, amba As Single, ambr As Single, ambg As Single, ambb As Single, _
        speca As Single, specr As Single, specg As Single, specb As Single, diffa As Single, diffr As Single, diffg As Single, diffb As Single, dirx As Single, diry As Single, dirz As Single, _
        atten0 As Single, atten1 As Single, atten2 As Single, falloff As Single, theta As Single, phi As Single, range As Single)

  Dim cnt As Integer
  LightCount = LightCount + 1
  cnt = LightCount - 1
  
  ReDim Preserve Lights(cnt) As D3DLIGHT8
  
  Lights(cnt).position = MakeVector(X, Y, Z)
  Lights(cnt).Type = LType
  Lights(cnt).Direction = MakeVector(dirx, diry, dirz)
  Lights(cnt).Ambient = MakeColorValue(amba, ambr, ambg, ambb)
  Lights(cnt).specular = MakeColorValue(speca, specr, specg, specb)
  Lights(cnt).diffuse = MakeColorValue(diffa, diffr, diffg, diffb)
  Lights(cnt).Attenuation0 = atten0
  Lights(cnt).Attenuation1 = atten1
  Lights(cnt).Attenuation2 = atten2
  Lights(cnt).falloff = falloff
  Lights(cnt).theta = theta
  Lights(cnt).phi = phi
  Lights(cnt).range = range
  
End Sub

' Adds an object file
Public Sub AddObject(ObjectFile As String, OX As Single, OY As Single, OZ As Single)
  
  Dim What As String
  Dim W As Integer
  Dim Width As Single
  Dim Height As Single
  Dim X As Single
  Dim Y As Single
  Dim Z As Single
  Dim FU As Single
  Dim FV As Single
  Dim Texture As Integer
  
  Open ObjectFile For Input As #2 ' Open the object file
    
    Do
      
      Input #2, What ' What kind of object
      If Left$(What, 1) = ";" Or What = "" Then GoTo Jump ' If the first char is a ; then it is a comment
      
      Select Case UCase(What)
        Case "FLOOR", "ROOF" ' Add a floor or a roof
          Input #2, Width, Height, X, Y, Z, Texture, FU, FV
          X = X + OX: Y = Y + OY: Z = Z + OZ
          AddFloorRoof Width, Height, X, Y, Z, Texture, FU, FV
        Case "WALL" ' Add a wall
          Input #2, W, Width, Height, X, Y, Z, Texture, FU, FV
          X = X + OX: Y = Y + OY: Z = Z + OZ
          Addwall W, Width, Height, X, Y, Z, Texture, FU, FV
          If W = 0 Then
            AddCollision True, Width, X, Z
          Else
            AddCollision False, Width, X, Z
          End If
      End Select
      
Jump:
      
    Loop Until EOF(2) ' Loop until end of file
  
  Close #2 ' Close the object file
  
  Exit Sub
  
End Sub

' Loads the map file
' The map contains walls, roofs, objects,
' lights, etc.
Public Sub LoadMap(MapFile As String)
  
  Dim What As String
  Dim W As Integer
  Dim Width As Single
  Dim Height As Single
  Dim X As Single
  Dim Y As Single
  Dim Z As Single
  Dim FU As Single
  Dim FV As Single
  Dim Texture As Integer
  Dim ObjFile As String
  
  Open MapFile For Input As #1 ' Open the map file
    
    Do
      
      Input #1, What ' What kind of object
      If Left$(What, 1) = ";" Or What = "" Then GoTo Jump ' If the first char is a ; then it is a comment
      
      Select Case UCase(What)
        Case "CAMERAPOS" ' Position
          Input #1, X, Y, Z
          camx = X: camy = Y: camz = Z
        Case "CAMERAROTXY" ' Angle and Pitch
          Input #1, X, Y
          Angle = X: pitch = Y
        Case "SETAMBLIGHT" ' Set the ambient light color
          Dim col(2) As Integer
          Input #1, col(0), col(1), col(2)
          D3DDevice.SetRenderState D3DRS_AMBIENT, RGB(col(0), col(1), col(2))
        Case "TURNLIGHTSOFF" ' Turn lights off
          D3DDevice.SetRenderState D3DRS_LIGHTING, 0
        Case "TURNLIGHTSON" ' Turn lights on
          D3DDevice.SetRenderState D3DRS_LIGHTING, 1
        Case "TERRAIN" ' Load a terrain
          Input #1, ObjFile, Height
          LoadTerrain TexturePath & ObjFile, Height
        Case "FLOOR", "ROOF" ' Add a floor or a roof
          Input #1, Width, Height, X, Y, Z, Texture, FU, FV
          AddFloorRoof Width, Height, X, Y, Z, Texture, FU, FV
        Case "WALL" ' Add a wall
          Input #1, W, Width, Height, X, Y, Z, Texture, FU, FV
          Addwall W, Width, Height, X, Y, Z, Texture, FU, FV
          If W = 0 Then
            AddCollision True, Width, X, Z
          Else
            AddCollision False, Width, X, Z
          End If
        Case "TREE" ' Add a tree
          Input #1, Width, Height, X, Y, Z, Texture
          AddTree Width, Height, X, Y, Z, Texture
        Case "SKY"
          Dim T(5) As Integer
          Input #1, T(0), T(1), T(2), T(3), T(4), T(5)
          AddSky T(0), T(1), T(2), T(3), T(4), T(5)
        Case "LIGHT" ' Add a light
          Dim LType As Integer, amba As Single, ambr As Single, ambg As Single, ambb As Single
          Dim speca As Single, specr As Single, specg As Single, specb As Single
          Dim diffa As Single, diffr As Single, diffg As Single, diffb As Single
          Dim dirx As Single, diry As Single, dirz As Single
          Dim atten0 As Single, atten1 As Single, atten2 As Single
          Dim falloff As Single, theta As Single, phi As Single, range As Single
          Input #1, X, Y, Z, LType, amba, ambr, ambg, ambb, speca, specr, specg, specb, diffa, diffr, diffg, diffb, dirx, diry, dirz, atten0, atten1, atten2, falloff, theta, phi, range
          AddLight X, Y, Z, LType, amba, ambr, ambg, ambb, speca, specr, specg, specb, diffa, diffr, diffg, diffb, dirx, diry, dirz, atten0, atten1, atten2, falloff, theta, phi, range
        Case "OBJECT"
          Input #1, ObjFile, X, Y, Z
          AddObject ObjectPath & ObjFile, X, Y, Z
      End Select
      
Jump:
      
    Loop Until EOF(1) ' Loop until end of file
  
  Close #1 ' Close the map file
  
  Exit Sub
  
End Sub
