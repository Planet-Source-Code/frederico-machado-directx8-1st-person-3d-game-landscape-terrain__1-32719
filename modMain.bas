Attribute VB_Name = "modMain"
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

Global DX As New DirectX8 ' The main DirectX 8 object
Global D3DX As New D3DX8 ' Used to help
Global D3D As Direct3D8 ' Used to create the D3DDevice
Global D3DDevice As Direct3DDevice8 ' Rendering device

Global VBuffers() As Direct3DVertexBuffer8 ' It is a buffer of our walls and floors
Global VBCount As Integer ' How many Vertex Buffers we have
Global Textures() As Direct3DTexture8 ' Contains te textures of walls and floors
Global TexCount As Integer ' How many textures we have
Global VBTex() As Integer ' saves what texture we will use to each wall and floor

Global Lights() As D3DLIGHT8 ' Saves our lights
Global LightCount As Integer ' How many lights we have

Global DI As DirectInput8 'this is DirectInput, used to monitor the keys on the keyboard in my case
Global DIDev As DirectInputDevice8 'this device will be the keyboard
Global DIState As DIKEYBOARDSTATE 'to check the state of keys

Global DIMouse As DirectInputDevice8 ' Mouse device
Global DIMState As DIMOUSESTATE ' to check mouse movements and clicks

Global camx, camy, camz As Single ' holds the camera pos

Global Angle, AngleConv As Single 'holds the angle, at which the camera is pointing
Global pitch As Single 'holds the pitch of the camera (this is where the camera is pointing in terms of the y axis, ie. up and down etc.)

' TEXT
Dim MainFont As D3DXFont
Dim MainFontDesc As IFont
Dim TextRect As RECT
Dim fnt As New StdFont

' Frame Rate Calculations
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim FPS_LastCheck As Long
Dim FPS_Count As Long
Dim FPS_Current As Integer

' a structure for custom vertex type
Public Type CUSTOMVERTEX
    position As D3DVECTOR   '3d position for vertex
    normal As D3DVECTOR     'normal used in lighting calculations
    color As Long           'color of the vertex
    tu As Single            'texture map coordinate
    tv As Single            'texture map coordinate
End Type

' custom FVF, which describes our custom vertex structure
Public Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

' Holds tree info
Public Type TREE
    v(3) As CUSTOMVERTEX
    vPos As D3DVECTOR
    iTreeTexture As Integer
End Type

Global matBillboardMatrix As D3DMATRIX   ' Used for billboard orientation
Global Trees() As TREE                ' Array of tree info

Public Const g_pi As Single = 3.141592653 'pi
Public Const g_90d As Single = g_pi / 2 '90 degrees in radians
Public Const g_180d As Single = g_pi '180 degrees in radians
Public Const g_270d As Single = (g_pi / 2) * 3 '270 degrees in radians
Public Const g_360d As Single = g_pi * 2 '360 degrees in radians

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Sub InitD3D()

  On Local Error Resume Next
  
  Set D3D = DX.Direct3DCreate() ' Create D3D
  If D3D Is Nothing Then GoTo D3DError
  
  Dim Mode As D3DDISPLAYMODE
  D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode ' Get the current display mode
  
  ' Set up the structure used to create the D3DDevice. Since we are now
  ' using more complex geometry, we will create a device with a zbuffer.
  ' the D3DFMT_D16 indicates we want a 16 bit z buffer.
  Dim D3Dpp As D3DPRESENT_PARAMETERS
  D3Dpp.SwapEffect = D3DSWAPEFFECT_DISCARD
  D3Dpp.BackBufferWidth = 800
  D3Dpp.BackBufferHeight = 600
  D3Dpp.BackBufferFormat = Mode.Format
  D3Dpp.BackBufferCount = 1
  D3Dpp.EnableAutoDepthStencil = 1
  D3Dpp.AutoDepthStencilFormat = D3DFMT_D16
  D3Dpp.Windowed = 1
  
  ' Create the D3DDevice
  ' If you do not have hardware 3d acceleration. Enable the reference rasterizer
  ' using the DirectX control panel and change D3DDEVTYPE_HAL to D3DDEVTYPE_REF
  Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3Dpp)
  If D3DDevice Is Nothing Then GoTo D3DError
  
  ' Turn off culling, so we see the front and back of walls and floors
  D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
  D3DDevice.SetRenderState D3DRS_DITHERENABLE, 1
  D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
  ' Turn on the zbuffer
  D3DDevice.SetRenderState D3DRS_ZENABLE, 1
  
  D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
  D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
  
  D3DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
  D3DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
    
  ' It makes the textures look better
  D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
  D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
  D3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_ANISOTROPIC
  D3DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, 16
  
  fnt.Name = "Tahoma"
  fnt.Size = 12
  Set MainFontDesc = fnt
  Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
  
  Exit Sub
  
D3DError:
  
  MsgBox "Unable to CreateDevice (see InitD3D() source for comments)", vbCritical, "Error"
  Unload frmMain

End Sub

Sub InitDI()

  ' Create Direct Input
  Set DI = DX.DirectInputCreate()

  ' Create keyboard device
  Set DIDev = DI.CreateDevice("GUID_SysKeyboard")
  ' Set common data format to keyboard
  DIDev.SetCommonDataFormat DIFORMAT_KEYBOARD
  DIDev.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
  DIDev.Acquire
  
  ' Create Mouse device
  Set DIMouse = DI.CreateDevice("GUID_SysMouse")
  ' Set common data format to mouse
  DIMouse.SetCommonDataFormat DIFORMAT_MOUSE
  DIMouse.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
  DIMouse.Acquire

End Sub

Sub DXKeyboard()

  DIDev.GetDeviceStateKeyboard DIState
  
  Dim OldPos As D3DVECTOR
  OldPos = MakeVector(camx, camy, camz)
  
  ' If esc is pressed then ends the game
  If DIState.Key(DIK_ESCAPE) <> 0 Then Unload frmMain
  
  ' Lets move forward
  If DIState.Key(DIK_UP) <> 0 Or DIState.Key(DIK_W) <> 0 Then
    If DIState.Key(DIK_LSHIFT) <> 0 Or DIState.Key(DIK_RSHIFT) <> 0 Then
      ' Running forward
      camx = camx + (Sin(AngleConv) * 0.8)
      camz = camz + (Cos(AngleConv) * 0.8)
    Else
      ' Walking forward
      camx = camx + (Sin(AngleConv) * 0.5)
      camz = camz + (Cos(AngleConv) * 0.5)
    End If
  End If
  
  ' Lets move back
  If DIState.Key(DIK_DOWN) <> 0 Or DIState.Key(DIK_S) <> 0 Then
    camx = camx - (Sin(AngleConv) * 0.5)
    camz = camz - (Cos(AngleConv) * 0.5)
  End If
  
  ' Lets rotate left
  If DIState.Key(DIK_LEFT) <> 0 Then
    Angle = Angle + (g_90d / 18)
    If Angle < 0 Then Angle = g_360d - (-Angle)
  End If
  
  ' Lets rotate right
  If DIState.Key(DIK_RIGHT) <> 0 Then
    Angle = Angle - (g_90d / 18)
    If Angle > g_360d Then Angle = 0 + (Angle - g_360d)
  End If
  
  ' Strafe left
  If DIState.Key(DIK_A) <> 0 Then
    camx = camx + (Sin(AngleConv - g_90d) * 0.5)
    camz = camz + (Cos(AngleConv - g_90d) * 0.5)
  End If
  
  ' Strafe right
  If DIState.Key(DIK_D) <> 0 Then
    camx = camx + (Sin(AngleConv + g_90d) * 0.5)
    camz = camz + (Cos(AngleConv + g_90d) * 0.5)
  End If
  
  ' Convert to correct angle system
  AngleConv = g_360d - Angle
  
  ' Get camera position and check collision
  Dim CVector As D3DVECTOR
  CVector = MakeVector(camx, camy, camz)
  
  ' Let check the collision detection
  CVector = CheckCollision(OldPos, CVector)
  
  ' Set the camera to the CheckCollision() results
  camx = CVector.X
  camz = CVector.Z
  camy = 7
  
  If LandSize > 0 Then ' There is a terrain loaded
    Dim col As Long
    col = frmMain.picHeight.Point(CSng(CInt(camx)), CSng(CInt(camz)))
    camy = (LandHeight * GreyScale(col)) + (10 * 0.3)
  End If

End Sub

Sub DXMouse()

  ' Lets get the mouse state
  DIMouse.GetDeviceStateMouse DIMState
  
  Angle = Angle - (DIMState.lX * 0.005)
  pitch = pitch - (DIMState.lY * 0.005)
  
  If pitch < -1.5 Then pitch = -1.5
  If pitch > 1.5 Then pitch = 1.5
  
End Sub

Sub SetupMatrices()

  Dim matView As D3DMATRIX
  Dim matRotation As D3DMATRIX
  Dim matPitch As D3DMATRIX
  Dim matLook As D3DMATRIX
  Dim matPos As D3DMATRIX
  Dim matWorld As D3DMATRIX
  Dim matProj As D3DMATRIX
    
  'setup world matrix
  D3DXMatrixIdentity matWorld
  D3DDevice.SetTransform D3DTS_WORLD, matWorld
  
  'make them identity matrices
  D3DXMatrixIdentity matView
  D3DXMatrixIdentity matPos
  D3DXMatrixIdentity matRotation
  D3DXMatrixIdentity matLook
  'rotate around x and y, for angle and pitch
  D3DXMatrixRotationY matRotation, Angle
  D3DXMatrixRotationX matPitch, pitch
  'multiply angle and pitch matrices together to create one 'look' matrix
  D3DXMatrixMultiply matLook, matRotation, matPitch
  'put the position of the camera into the translation matrix, matPos
  D3DXMatrixTranslation matPos, -camx, -camy, -camz
  'multiply that with the look matrix to make the complete view matrix
  D3DXMatrixMultiply matView, matPos, matLook
  'which we can then set as the view matrix:
  D3DDevice.SetTransform D3DTS_VIEW, matView
  'update details form
  
  ' Rotate the trees to the camera angle
  D3DXMatrixRotationY matBillboardMatrix, -Angle

  'setup the projection matrix
  D3DXMatrixPerspectiveFovLH matProj, g_pi / 4, 1, 1, 10000
  D3DDevice.SetTransform D3DTS_PROJECTION, matProj

End Sub

Sub SetupLights()

  Dim i As Integer
  Dim col As D3DCOLORVALUE
  
  ' Set up a material
  Dim mtrl As D3DMATERIAL8
  With col: .r = 1: .g = 1: .b = 1: .a = 0:   End With
  mtrl.diffuse = col
  mtrl.Ambient = col
  D3DDevice.SetMaterial mtrl
  
  ' Lets render our lights
  For i = 0 To LightCount - 1
    D3DDevice.SetLight i, Lights(i) 'let d3d know about the light
    D3DDevice.LightEnable i, 1      'turn it on
  Next
  
End Sub

Sub Render()

  Dim v As CUSTOMVERTEX
  Dim SizeofVertex As Long
  Dim i As Integer
  
  Do
  
    DoEvents
  
    ' Get mouse and keyboard state
    DXMouse
    DXKeyboard
    
    ' Clear the backbuffer to a black color
    ' Clear the zbuffer to 1
    D3DDevice.Clear ByVal 0, ByVal 0, D3DCLEAR_ZBUFFER, &H0, 1, 0
    
    ' Begin the scene
    D3DDevice.BeginScene
    
    ' Setup the world, view, and projection matrices
    SetupMatrices
    ' Setup the lights
    SetupLights
    
    D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
    ' Draw the triangles in the vertex buffer
    ' Note we are now using a triangle strip of vertices
    ' instead of a triangle list
    SizeofVertex = Len(v)

    RenderLandscape SizeofVertex

    'draw the contents of our vertex buffers, remembering to change to the correct textures.
    For i = 0 To VBCount - 1
      D3DDevice.SetTexture 0, Textures(VBTex(i))
      D3DDevice.SetStreamSource 0, VBuffers(i), SizeofVertex
      D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    Next
    
    TextRect.Top = 0
    TextRect.bottom = 20
    TextRect.Right = 75
    D3DX.DrawText MainFont, &HFFFFCC00, CStr(FPS_Current) & " FPS", TextRect, 0
    
    ' End the scene
    D3DDevice.EndScene
    
    ' Present the backbuffer contents to the front buffer (screen)
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
  
    If GetTickCount() - FPS_LastCheck >= 1000 Then
        FPS_Current = FPS_Count
        FPS_Count = 0 'reset the counter
        FPS_LastCheck = GetTickCount()
    End If
    FPS_Count = FPS_Count + 1
  
  Loop

End Sub

Sub ExitD3D()
  
  ' Lets 'unload' everything
  Set DX = Nothing
  Set D3DX = Nothing
  Set D3D = Nothing
  Set D3DDevice = Nothing
  Set DI = Nothing
  
  DIDev.Unacquire
  DIMouse.Unacquire
  
  Set DIDev = Nothing
  Set DIMouse = Nothing
    
End Sub

' It makes a 3D vector
Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
  MakeVector.X = X
  MakeVector.Y = Y
  MakeVector.Z = Z
End Function

' It makes a D3D Color Value
Function MakeColorValue(ByVal a As Single, ByVal r As Single, ByVal g As Single, ByVal b As Single) As D3DCOLORVALUE
  MakeColorValue.a = a
  MakeColorValue.r = r
  MakeColorValue.g = g
  MakeColorValue.b = b
End Function

Sub CalculateNormals(CVector() As CUSTOMVERTEX)

  On Error Resume Next
  
  Dim i As Byte
  Dim i2 As Byte
  Dim vec As D3DVECTOR
  Dim vec2 As D3DVECTOR
  Dim vec3 As D3DVECTOR
  Dim vec4 As D3DVECTOR
            
  For i = 0 To UBound(CVector) Step 2
    vec4 = GetPolygonNormal(CVector(i).position, CVector(i + 1).position, CVector(i + 2).position)
    If i > 0 Then
      vec2 = vec4
      vec3 = CVector(i).normal
      vec = AverageOf2Vectors(vec2, vec3)
    Else
      vec = vec4
    End If
    CVector(i).normal = vec
    If i > 2 Then
      vec2 = vec4
      vec3 = CVector(i + 1).normal
      vec = AverageOf2Vectors(vec2, vec3)
    Else
      vec = vec4
    End If
    CVector(i + 1).normal = vec
    If i > 2 Then
      vec2 = vec4
      vec3 = CVector(i + 1).normal
      vec = AverageOf2Vectors(vec2, vec3)
    Else
      vec = vec4
    End If
    CVector(i + 2).normal = vec
  Next

End Sub

Function AverageOf2Vectors(vec1 As D3DVECTOR, vec2 As D3DVECTOR) As D3DVECTOR

  D3DXVec3Add AverageOf2Vectors, vec1, vec2
  AverageOf2Vectors.X = AverageOf2Vectors.X / 2
  AverageOf2Vectors.Y = AverageOf2Vectors.Y / 2
  AverageOf2Vectors.Z = AverageOf2Vectors.Z / 2

End Function

' ***********************************
' Richard Hayden's Subs
' ***********************************

'get a polygon's normal
Public Function GetPolygonNormal(vecPolygon1 As D3DVECTOR, vecPolygon2 As D3DVECTOR, vecPolygon3 As D3DVECTOR) As D3DVECTOR
    Dim vec1 As D3DVECTOR
    Dim vec2 As D3DVECTOR
    Dim vtemp As D3DVECTOR

    vec1.X = (vecPolygon2.X - vecPolygon1.X)
    vec1.Y = (vecPolygon2.Y - vecPolygon1.Y)
    vec1.Z = (vecPolygon2.Z - vecPolygon1.Z)

    vec2.X = (vecPolygon3.X - vecPolygon1.X)
    vec2.Y = (vecPolygon3.Y - vecPolygon1.Y)
    vec2.Z = (vecPolygon3.Z - vecPolygon1.Z)

    vtemp = CrossProduct(vec1, vec2)

    GetPolygonNormal = vtemp
End Function

'get crossproduct
Public Function CrossProduct(vecPolygon1 As D3DVECTOR, vecPolygon2 As D3DVECTOR) As D3DVECTOR
    Dim vtemp As D3DVECTOR

    vtemp.X = (vecPolygon1.Y * vecPolygon2.Z) - (vecPolygon1.Z * vecPolygon2.Y)
    vtemp.Y = (vecPolygon1.Z * vecPolygon2.X) - (vecPolygon1.X * vecPolygon2.Z)
    vtemp.Z = (vecPolygon1.X * vecPolygon2.Y) - (vecPolygon1.Y * vecPolygon2.X)

    vtemp = NormalizeVector(vtemp)

    CrossProduct = vtemp
End Function

'normalize a vector
Public Function NormalizeVector(vec As D3DVECTOR) As D3DVECTOR
    Dim vLength As Single
    Dim vector As D3DVECTOR

    vLength = GetVectorLength(vec)

    vector.X = vec.X / vLength
    vector.Y = vec.Y / vLength
    vector.Z = vec.Z / vLength

    NormalizeVector = vector
End Function

'get a vector's length
Public Function GetVectorLength(vec As D3DVECTOR) As Single
    GetVectorLength = Sqr((vec.X * vec.X) + (vec.Y * vec.Y) + (vec.Z * vec.Z))
End Function
