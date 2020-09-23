Attribute VB_Name = "modLandscape"
' DirectX8 1st Person 3D Game VERSION 1.2 (19/03/2002)
' By Frederico Machado (indiofu@bol.com.br)
' Please vote for me if you like the game.
'
' I'd like to thanks to all DirectX programmers at PSC
' I've downloaded all I found about DirectX, and I learned
' too much from it.
' Special thanks to Richard Hayden, who introduced me in
' DirectX 8 with his 3D World and helped me a lot.
' Thank you very very much Richard!
'
' Sorry my English, I'm Brazilian! :)
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

Global LandHeight As Single
Global LandWidth As Integer
Global LandSize As Single

Global LandscapeVB() As Direct3DVertexBuffer8
Global LandVBCount As Long

Global TreeVB() As Direct3DVertexBuffer8 ' Our tree vertex buffer
Global TreeVBCount As Integer ' How many trees we have
Global TreeTextures() As Direct3DTexture8 ' Tree textures
Global TreeTexCount As Integer ' How many tree textures we have
Global TreeVBTex() As Integer ' saves what texture we will use to each tree

Global SkyVB(5) As Direct3DVertexBuffer8 ' 6 vertex buffers to the sky
Global SkyTextures(5) As Direct3DTexture8  ' sky textures
Global SkyTex(5) As Integer ' what texture we will use in each side of the sky
Dim SkyAdded As Boolean

Dim Map() As Long

' loads a terrain
Sub LoadTerrain(HeightMap As String, Optional Height As Single = 10)
  
  Dim GridSize As Single
  Dim X As Integer
  Dim Y As Integer
  Dim col As Long
  Dim xx As Single
  Dim yy As Single
  Dim LandVectors() As CUSTOMVERTEX
  
  frmMain.picHeight = LoadPicture(HeightMap)
  
  LandSize = frmMain.picHeight.Width

  LandHeight = Height
  GridSize = LandSize / LandSize
  
  ReDim Map(LandSize, LandSize)
  
  
  xx = 0
  For X = 1 To LandSize
    xx = xx + GridSize
    yy = 0
    For Y = 1 To LandSize
      yy = yy + GridSize
      ReDim LandVectors(0 To 3)
      
      col = frmMain.picHeight.Point(xx, yy + GridSize)
      If col = -1 Then col = 0
      LandVectors(0).position = MakeVector(X, Height * GreyScale(col), Y + 1)
      
      col = frmMain.picHeight.Point(xx + GridSize, yy + GridSize)
      If col = -1 Then col = 0
      LandVectors(1).position = MakeVector(X + 1, Height * GreyScale(col), Y + 1)
      
      col = frmMain.picHeight.Point(xx, yy)
      If col = -1 Then col = 0
      LandVectors(2).position = MakeVector(X, Height * GreyScale(col), Y)
      
      col = frmMain.picHeight.Point(xx + GridSize, yy)
      If col = -1 Then col = 0
      LandVectors(3).position = MakeVector(X + 1, Height * GreyScale(col), Y)
      
      CalculateNormals LandVectors
      LandVBCount = LandVBCount + 1
      
      Dim texXCoord As Byte, texYCoord As Byte
      Dim i As Integer
      Const n = 6
      texXCoord = Int(LandVectors(2).position.X / n)
      texYCoord = Int(LandVectors(2).position.Z / n)
      
      For i = 0 To 3
        LandVectors(i).tu = (LandVectors(i).position.X / n) - Int(LandVectors(i).position.X / n)
        LandVectors(i).tv = (LandVectors(i).position.Z / n) - Int(LandVectors(i).position.Z / n)
        
        If ((LandVectors(i).position.Z / n) - 1) = texYCoord Then
          LandVectors(i).tv = 1#
        End If
        
        If ((LandVectors(i).position.X / n) - 1) = texXCoord Then
          LandVectors(i).tu = 1#
        End If
      Next
      
      ReDim Preserve LandscapeVB(LandVBCount - 1) As Direct3DVertexBuffer8
      Set LandscapeVB(LandVBCount - 1) = D3DDevice.CreateVertexBuffer(Len(LandVectors(0)) * 4, _
          0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
      D3DVertexBuffer8SetData LandscapeVB(LandVBCount - 1), 0, Len(LandVectors(0)) * 4, 0, LandVectors(0)
    
      DoEvents
    
    Next
  Next
  
End Sub

' Creates the sky box
Public Sub AddSky(fronttex As Integer, backtex As Integer, lefttex As Integer, righttex As Integer, toptex As Integer, bottomtex As Integer)

  SkyTex(0) = fronttex: SkyTex(1) = backtex
  SkyTex(2) = lefttex: SkyTex(3) = righttex
  SkyTex(4) = toptex: SkyTex(5) = bottomtex

  Dim v(0 To 3) As CUSTOMVERTEX
  Dim VertexSizeInBytes As Long
  
  VertexSizeInBytes = Len(v(0))
  
  ' The front of the sky
  v(0).position = MakeVector(-10, 10, 10)
  v(1).position = MakeVector(10, 10, 10)
  v(2).position = MakeVector(-10, -10, 10)
  v(3).position = MakeVector(10, -10, 10)
  v(0).tu = 0: v(0).tv = 0
  v(1).tu = 1: v(1).tv = 0
  v(2).tu = 0: v(2).tv = 1
  v(3).tu = 1: v(3).tv = 1
  
  Set SkyVB(0) = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
       0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData SkyVB(0), 0, VertexSizeInBytes * 4, 0, v(0)

  ' The back of the sky
  v(0).position = MakeVector(-10, 10, -10)
  v(1).position = MakeVector(10, 10, -10)
  v(2).position = MakeVector(-10, -10, -10)
  v(3).position = MakeVector(10, -10, -10)
  v(0).tu = 1: v(0).tv = 0
  v(1).tu = 0: v(1).tv = 0
  v(2).tu = 1: v(2).tv = 1
  v(3).tu = 0: v(3).tv = 1
  
  Set SkyVB(1) = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
       0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData SkyVB(1), 0, VertexSizeInBytes * 4, 0, v(0)

  ' The left of the sky
  v(0).position = MakeVector(-10, 10, -10)
  v(1).position = MakeVector(-10, 10, 10)
  v(2).position = MakeVector(-10, -10, -10)
  v(3).position = MakeVector(-10, -10, 10)
  v(0).tu = 0: v(0).tv = 0
  v(1).tu = 1: v(1).tv = 0
  v(2).tu = 0: v(2).tv = 1
  v(3).tu = 1: v(3).tv = 1
  
  Set SkyVB(2) = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
       0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData SkyVB(2), 0, VertexSizeInBytes * 4, 0, v(0)

  ' The right of the sky
  v(0).position = MakeVector(10, 10, -10)
  v(1).position = MakeVector(10, 10, 10)
  v(2).position = MakeVector(10, -10, -10)
  v(3).position = MakeVector(10, -10, 10)
  v(0).tu = 1: v(0).tv = 0
  v(1).tu = 0: v(1).tv = 0
  v(2).tu = 1: v(2).tv = 1
  v(3).tu = 0: v(3).tv = 1
  
  Set SkyVB(3) = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
       0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData SkyVB(3), 0, VertexSizeInBytes * 4, 0, v(0)

  ' The top of the sky
  v(0).position = MakeVector(10, 10, -10)
  v(1).position = MakeVector(-10, 10, -10)
  v(2).position = MakeVector(10, 10, 10)
  v(3).position = MakeVector(-10, 10, 10)
  v(0).tu = 1: v(0).tv = 1
  v(1).tu = 1: v(1).tv = 0
  v(2).tu = 0: v(2).tv = 1
  v(3).tu = 0: v(3).tv = 0
  
  Set SkyVB(4) = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
       0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData SkyVB(4), 0, VertexSizeInBytes * 4, 0, v(0)

  ' The bottom of the sky
  v(0).position = MakeVector(10, -10, -10)
  v(1).position = MakeVector(-10, -10, -10)
  v(2).position = MakeVector(10, -10, 10)
  v(3).position = MakeVector(-10, -10, 10)
  v(0).tu = 0: v(0).tv = 0
  v(1).tu = 1: v(1).tv = 0
  v(2).tu = 0: v(2).tv = 1
  v(3).tu = 1: v(3).tv = 1
  
  Set SkyVB(5) = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
       0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData SkyVB(5), 0, VertexSizeInBytes * 4, 0, v(0)

  SkyAdded = True

End Sub

' Adds a tree to the world
Public Sub AddTree(Width As Single, Height As Single, X As Single, Y As Single, Z As Single, Texture As Integer)

  Dim cnt As Integer
  
  Dim v(0 To 3) As CUSTOMVERTEX
  Dim VertexSizeInBytes As Long
  
  VertexSizeInBytes = Len(v(0))

  TreeVBCount = TreeVBCount + 1
  cnt = TreeVBCount - 1

  ReDim Preserve Trees(cnt) As TREE
  
  If LandSize > 0 Then ' There is a terrain loaded
    Dim col As Long
    col = frmMain.picHeight.Point(CSng(CInt(X)), CSng(CInt(Z)))
    Y = (LandHeight * GreyScale(col))
  End If
  
  Trees(cnt).vPos.X = X
  Trees(cnt).vPos.Z = Z
  Trees(cnt).vPos.Y = Y
  
  Trees(cnt).v(0).position = MakeVector(-Width, 0 * Height, 0)
  Trees(cnt).v(1).position = MakeVector(-Width, 2 * Height, 0)
  Trees(cnt).v(2).position = MakeVector(Width, 0 * Height, 0)
  Trees(cnt).v(3).position = MakeVector(Width, 2 * Height, 0)
  Trees(cnt).v(0).tu = 0: Trees(cnt).v(0).tv = 1
  Trees(cnt).v(1).tu = 0: Trees(cnt).v(1).tv = 0
  Trees(cnt).v(2).tu = 1: Trees(cnt).v(2).tv = 1
  Trees(cnt).v(3).tu = 1: Trees(cnt).v(3).tv = 0

  Trees(cnt).iTreeTexture = Texture
  
  ReDim Preserve TreeVB(cnt) As Direct3DVertexBuffer8
  ' Create the vertex buffer
  Set TreeVB(cnt) = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_MANAGED)
  ' fill the vertex buffer from our array
  D3DVertexBuffer8SetData TreeVB(cnt), 0, VertexSizeInBytes * 4, 0, Trees(cnt).v(0)

End Sub

Sub RenderLandscape(SizeofVertex As Long)

  Dim i As Integer
  
  If SkyAdded Then RenderSky SizeofVertex
  
  If LandSize > 0 Then ' There is a terrain loaded
  
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    For i = 0 To LandVBCount - 1
  
      D3DDevice.SetTexture 0, Textures(0)
      D3DDevice.SetStreamSource 0, LandscapeVB(i), SizeofVertex
      D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
  
    Next
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

  End If

  If TreeTexCount >= 0 Then RenderTrees SizeofVertex

End Sub

Sub RenderSky(SizeofVertex As Long)

  Dim matView As D3DMATRIX, matViewSave As D3DMATRIX, hr As Long
  Dim matProj As D3DMATRIX
  Dim i As Integer

  ' Disable the Zbuffer and render the sky
  D3DDevice.GetTransform D3DTS_VIEW, matViewSave
  matView = matViewSave
  matView.m41 = 0: matView.m42 = 0: matView.m43 = 0
  D3DDevice.SetTransform D3DTS_VIEW, matView
  D3DDevice.SetRenderState D3DRS_ZENABLE, 0
    
  ' It makes the sky textures fill correctly
  ' but the sky looks horrible
  D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
  
  D3DXMatrixPerspectiveFovLH matProj, g_pi / 3.5, 1, 1, 10000
  D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    
  ' Setup the sky
  For i = 0 To 5
    D3DDevice.SetTexture 0, SkyTextures(SkyTex(i))
    D3DDevice.SetStreamSource 0, SkyVB(i), SizeofVertex
    D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
  Next
  
  D3DXMatrixPerspectiveFovLH matProj, g_pi / 4, 1, 1, 10000
  D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    
  D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    
  D3DDevice.SetTransform D3DTS_VIEW, matViewSave
  D3DDevice.SetRenderState D3DRS_ZENABLE, 1
  ' Enable Zbuffer again

End Sub

Sub RenderTrees(SizeofVertex As Long)

  Dim i As Integer
  
  ' We will use trasparency in trees
  D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1  'TRUE
  D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
  D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
  D3DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1 'TRUE
  D3DDevice.SetRenderState D3DRS_ALPHAREF, &H8&
  D3DDevice.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
  
  D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
  ' Render our trees and set their position in the matrix
  For i = 0 To TreeVBCount - 1
    D3DDevice.SetTexture 0, TreeTextures(Trees(i).iTreeTexture)
    matBillboardMatrix.m41 = Trees(i).vPos.X
    matBillboardMatrix.m42 = Trees(i).vPos.Y
    matBillboardMatrix.m43 = Trees(i).vPos.Z
    D3DDevice.SetTransform D3DTS_WORLD, matBillboardMatrix
  
    D3DDevice.SetStreamSource 0, TreeVB(i), SizeofVertex
    D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
  Next
  D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
  
  D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0  'TRUE
  D3DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 0 'TRUE
  
  Dim matWorld As D3DMATRIX
  D3DXMatrixIdentity matWorld
  D3DDevice.SetTransform D3DTS_WORLD, matWorld

End Sub

Function GreyScale(LongCol As Long) As Single
  
  Dim r As Single
  Dim g As Single
  Dim b As Single
  
  Long2RGB LongCol, r, g, b
  GreyScale = (r + b + g) / 765

End Function

Sub Long2RGB(LongCol As Long, r As Single, g As Single, b As Single)

  r = LongCol And 255
  g = (LongCol And 65280) \ 256&
  b = (LongCol And 16711680) \ 65535

End Sub
