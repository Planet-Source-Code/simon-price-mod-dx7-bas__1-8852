Attribute VB_Name = "ModDX7"
'ModDX7 - by Simon Price
'a module of simple funtions to make DirectX 7 easier to program

Public DirectX As New DirectX7
Public DX_Draw As DirectDraw7

Dim InExMode As Boolean

Sub SetDisplayMode(Width As Integer, Height As Integer, Colors As Byte)
'set's the display mode to the required size and colors
 DX_Draw.SetDisplayMode Width, Height, Colors, 0, DDSDM_DEFAULT
End Sub

Sub RestoreDisplayMode()
'puts the screen back to normal
 DX_Draw.RestoreDisplayMode
End Sub

Sub CreateSurfaceFromFile(Surface As DirectDrawSurface7, Surfdesc As DDSURFACEDESC2, Filename As String, Width As Integer, Height As Integer)
'loads a bitmap from a file and makes a pic from it
     Surfdesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
     Surfdesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
     Surfdesc.lWidth = Width
     Surfdesc.lHeight = Height
     Set Surface = DX_Draw.CreateSurfaceFromFile(Filename, Surfdesc)
End Sub

Sub Init(hwnd As Long)
If InExMode Then Exit Sub

'starts up everyfink
On Error GoTo CrapThingDontWork
'creates direct draw. whopee
Set DX_Draw = DirectX.DirectDrawCreate("")
'allow us to do cool stuff
DX_Draw.SetCooperativeLevel hwnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
InExMode = True

CrapThingDontWork:

End Sub

Sub CreateSurface(Surface As DirectDrawSurface7, Surfdesc As DDSCAPS2, Width As Integer, Height As Integer)
'creates a plain pic
     Surfdesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
     Surfdesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
     Surfdesc.lWidth = Width
     Surfdesc.lHeight = Height
End Sub

Sub StretchBlt(Pic As DirectDrawSurface7, x As Integer, y As Integer, Width As Integer, Height As Integer, DestPic As DirectDrawSurface7, DestX As Integer, DestY As Integer, DestWidth As Integer, DestHeight As Integer)
WaitTillOK
Dim Box As RECT
Box.Left = x
Box.Top = y
Box.Right = x + Width
Box.Bottom = y + Height

Dim DestBox As RECT
DestBox.Left = DestX
DestBox.Top = DestY
DestBox.Right = DestX + DestWidth
DestBox.Bottom = DestY + DestHeight

Pic.Blt DestBox, DestPic, Box, DDBLT_WAIT
End Sub

Sub BitBlt(Pic As DirectDrawSurface7, x As Integer, y As Integer, Width As Integer, Height As Integer, DestPic As DirectDrawSurface7, DestX As Integer, DestY As Integer)
WaitTillOK
Dim DestBox As RECT
DestBox.Left = DestX
DestBox.Top = DestY
DestBox.Right = DestX + DestWidth
DestBox.Bottom = DestY + DestHeight

Pic.BltFast x, y, DestPic, DestBox, DDBLTFAST_WAIT
End Sub

Sub WaitTillOK()
Dim bRestore As Boolean

bRestore = False
Do Until ExModeActive 'short way of saying "do until it returns true"
    DoEvents 'Lets windows do other things
    bRestore = True
Loop

' if we lost and got back the surfaces, then restore them
DoEvents 'Lets windows do it's things
If bRestore Then
    bRestore = False
    ddraw.RestoreAllSurfaces
    TestForm.InitSurfaces ' must init the surfaces again if they we're lost. When this happens the first line of initsurfaces is important
End If
End Sub

Function ExModeActive() As Boolean
     Dim TestCoopRes As Long 'holds the return value of the test.

     TestCoopRes = DX_Draw.TestCooperativeLevel 'Tells DDraw to do the test

     If (TestCoopRes = DD_OK) Then
         ExModeActive = True 'everything is fine
     Else
         ExModeActive = False 'this computer doesn't support this mode
     End If
 End Function
 
Sub EndIt(hwnd As Long)
DX_Draw.SetCooperativeLevel hwnd, DDSCL_NORMAL
InExMode = False
End Sub

Sub AddColorKey(Surface As DirectDrawSurface7, ColorKey As DDCOLORKEY, low As Long, high As Long)
ColorKey.low = low
ColorKey.high = high
Surface.SetColorKey DDCKEY_SRCBLT, ColorKey
End Sub
