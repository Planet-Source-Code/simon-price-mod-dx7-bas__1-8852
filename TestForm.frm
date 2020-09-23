VERSION 5.00
Begin VB.Form TestForm 
   Caption         =   "Test ModDX7"
   ClientHeight    =   3828
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   4968
   LinkTopic       =   "Form1"
   ScaleHeight     =   3828
   ScaleWidth      =   4968
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer PrintFramerate 
      Interval        =   1000
      Left            =   480
      Top             =   600
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start - change the screen res and bounce a crappy ball around the screen please."
      Height          =   972
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   1932
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim View As DirectDrawSurface7
Dim ViewDesc As DDSURFACEDESC2
Dim ViewCaps As DDSCAPS2
Dim BackBuffer As DirectDrawSurface7
Dim BackBufferDesc As DDSURFACEDESC2
Dim BackBufferCaps As DDSCAPS2
Dim Background As DirectDrawSurface7
Dim BackgroundDesc As DDSURFACEDESC2
Dim Ball As DirectDrawSurface7
Dim BallDesc As DDSURFACEDESC2
Dim BallColorKey As DDCOLORKEY

Dim Key As Byte
Dim Frames As Integer
Dim BallX As Integer
Dim BallY As Integer
Dim BallXM As Integer
Dim BallYM As Integer


Public Sub InitSurfaces()
Set View = Nothing
Set BackBuffer = Nothing

ViewDesc.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
ViewDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
ViewDesc.lBackBufferCount = 1
Set View = DX_Draw.CreateSurface(ViewDesc)

BackBufferCaps.lCaps = DDSCAPS_BACKBUFFER
Set BackBuffer = View.GetAttachedSurface(BackBufferCaps)
BackBuffer.GetSurfaceDesc ViewDesc

Set Background = Nothing
Set Ball = Nothing
ModDX7.CreateSurfaceFromFile Background, BackgroundDesc, App.Path & "\ACheapBackground.bmp", 640, 480
ModDX7.CreateSurfaceFromFile Ball, BallDesc, App.Path & "\Ball.bmp", 100, 100
End Sub

Private Sub CmdStart_Click()
'On Error GoTo WotaLoadaCack
ModDX7.Init Me.hwnd
ModDX7.SetDisplayMode 640, 480, 16

InitSurfaces

MainLoop

WotaLoadaCack:
ModDX7.EndIt Me.hwnd
End
End Sub

Sub MainLoop()
BallX = 320
BallY = 240
BallXM = 5
BallYM = 5
Dim Box As RECT
Dim BallBox As RECT
On Error GoTo CrappyErrorAlert
Box.Bottom = 480
Box.Right = 640
BallBox.Right = 100
BallBox.Bottom = 100
ModDX7.AddColorKey Ball, BallColorKey, vbBlack, vbBlack

Do
DoEvents
MoveBall
BackBuffer.BltFast 0, 0, Background, Box, DDBLTFAST_WAIT
BackBuffer.BltFast BallX, BallY, Ball, BallBox, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
View.Flip Nothing, DDFLIP_WAIT
Frames = Frames + 1
Loop Until GetKeyState(vbKeyEscape)

CrappyErrorAlert:
ClearUp
End Sub

Sub MoveBall()
Dim x As Integer
Dim y As Integer

x = BallX + BallXM
y = BallY + BallYM

Select Case x
  Case 0 To 540
    BallX = x
  Case Else
    BallXM = -BallXM
End Select

Select Case y
  Case 0 To 380
    BallY = y
  Case Else
    BallYM = -BallYM
End Select
End Sub

Sub ClearUp()
ModDX7.RestoreDisplayMode
ModDX7.EndIt (hwnd)
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Key = KeyCode
End Sub

Private Sub PrintFramerate_Timer()
Debug.Print Frames
Frames = 0
End Sub
