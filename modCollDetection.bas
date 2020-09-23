Attribute VB_Name = "modCollDetection"
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

Public Type OBJ_COORDS
  xx As Boolean
  Width As Single
  X As Single
  Z As Single
End Type

Global CollObjects() As OBJ_COORDS ' It will contains our walls
Public ObjCount As Integer ' The number of objects that we have to test

' This Sub add a wall to the collision detection
' XX = it is true if the wall is a front or a back wall
' other case it is false, easy!
' Width is the size of the wall
' X = position in the X coord where it will be placed
' Z = position in the Z coord where it will be placed
Public Sub AddCollision(xx As Boolean, Width As Single, X As Single, Z As Single)
  
  ObjCount = ObjCount + 1 ' Add one to the number of objects
  ReDim Preserve CollObjects(ObjCount - 1) As OBJ_COORDS ' Add one to the objects
  
  ' Set the parameters
  CollObjects(ObjCount - 1).xx = xx
  CollObjects(ObjCount - 1).Width = Width
  CollObjects(ObjCount - 1).X = X
  CollObjects(ObjCount - 1).Z = Z
  
End Sub

Public Function CheckCollision(OldPos As D3DVECTOR, NewPos As D3DVECTOR) As D3DVECTOR
  
  For i = 0 To ObjCount - 1
  
    If CollObjects(i).xx = True Then ' It is a front or a back wall
      If NewPos.Z < (CollObjects(i).Z + 1.5) And NewPos.Z > (CollObjects(i).Z - 1.5) Then ' Test if we are hitting the wall
        If NewPos.X >= CollObjects(i).X - 1 And NewPos.X <= (CollObjects(i).X + CollObjects(i).Width) + 1 Then ' Verify if we are between the start and the end of the wall
          If OldPos.Z > CollObjects(i).Z Then ' Verify what side of the wall we are
            NewPos.Z = (CollObjects(i).Z + 1.5)
          ElseIf OldPos.Z < CollObjects(i).Z Then ' Verify what side of the wall we are
            NewPos.Z = (CollObjects(i).Z - 1.5)
          End If
          GoTo Jump ' Lets verify the next wall
        End If
      End If
    Else ' It is a left or a right wall
      If NewPos.X < (CollObjects(i).X + 1.5) And NewPos.X > (CollObjects(i).X - 1.5) Then ' Test if we are hitting the wall
        If NewPos.Z >= CollObjects(i).Z - 1 And NewPos.Z <= (CollObjects(i).Z + CollObjects(i).Width) + 1 Then ' Verify if we are between the start and the end of the wall
          If OldPos.X > CollObjects(i).X Then ' Verify what side of the wall we are
            NewPos.X = (CollObjects(i).X + 1.5)
          ElseIf OldPos.X < CollObjects(i).X Then ' Verify what side of the wall we are
            NewPos.X = (CollObjects(i).X - 1.5)
          End If
          GoTo Jump ' Lets verify the next wall
        End If
      End If
    End If
    
Jump:
    
  Next ' Test the next wall
  
  ' Set our new position
  CheckCollision.X = NewPos.X: CheckCollision.Z = NewPos.Z
  
End Function

