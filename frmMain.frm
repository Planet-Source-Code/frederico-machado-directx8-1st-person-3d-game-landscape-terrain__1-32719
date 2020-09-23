VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "DirectX 8"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHeight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Load()

  Path = App.Path  ' Set the variable to the app path
  If Right$(Path, 1) <> "\" Then Path = Path & "\"  ' Just to be sure
  TexturePath = Path & "Textures\" ' Set the texture path
  MeshPath = Path & "Meshes\" ' Set the mesh path
  ObjectPath = Path & "Objects\" ' Set the object path
  DataPath = Path & "Data\" ' Set the data path

  Show ' Show the form, someone said that dumb things happens if we don't show the form
  
  'ShowCursor 0 ' Hide the mouse cursor
  
  DoEvents ' Let the PC do what it has to do
  
  InitD3D ' Initialize Direct3D
  InitDI ' Initialize Direct Input
  
  ' Load textures
  LoadTextures DataPath & "Textures.ini"
  LoadSkyTextures DataPath & "SkyTextures.ini"
  LoadTreeTextures DataPath & "TreeTextures.ini"
  
  ' Load the map file
  LoadMap Path & "Maps\sample.txt"
  
  ' Render the world
  Render

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  ExitD3D ' Unload DirectX
  
  ShowCursor 1 ' Show the mouse cursor
  
  DoEvents ' Let the PC do what it has to do
  
  End ' Guess...
  
End Sub
