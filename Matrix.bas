Attribute VB_Name = "Module1"
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Option Explicit

Public dxDirectX As New DirectX7
Public dxDirectDraw7 As DirectDraw7

Public sdFront As DDSURFACEDESC2
Public srfFront As DirectDrawSurface7
Public sdBackBuffer As DDSURFACEDESC2
Public srfBackBuffer As DirectDrawSurface7

Public Type typColumnItem
  x As Integer
  y As Integer
  intGreenColour As Integer
  intCharacter As Integer
End Type

Public Type typColumn
  Item() As typColumnItem
  intX As Integer
  intY As Integer
  intLength As Integer
  booActive As Boolean
  booOffScreen As Boolean
  intCounter As Integer
End Type

Public intColumn(79) As typColumn
Public fntCustom As New StdFont
Public intTotalColumns As Integer
Public booAlternate As Boolean





Sub Main()

SetupProg

Randomize

srfFront.SetFontBackColor vbBlack
srfFront.SetFontTransparency False


With fntCustom
    .Name = "Matrix"
    .Size = 8
    .Bold = False
End With

srfFront.SetFont fntCustom

Dim i As Integer, h As Integer


For i = 0 To 79
    intColumn(i).intLength = Int(Rnd * 10) + 30
    ReDim intColumn(i).Item(intColumn(i).intLength)
    intColumn(i).booActive = False
    intColumn(i).intCounter = 0
    intColumn(i).booOffScreen = False
    intColumn(i).intX = i
    intColumn(i).intY = 0
    For h = 0 To intColumn(i).intLength
        intColumn(i).Item(h).intCharacter = Int(Rnd * 43) + 65
        intColumn(i).Item(h).intGreenColour = Int(Rnd * 255)
    Next h
Next i

intTotalColumns = 0
booAlternate = True
Form1.tmrMain.Enabled = True


End Sub

Sub SetupProg()

ShowCursor (0)

Set srfBackBuffer = Nothing
Set srfFront = Nothing

Set dxDirectDraw7 = dxDirectX.DirectDrawCreate("")
dxDirectDraw7.SetCooperativeLevel Form1.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWMODEX
Call dxDirectDraw7.SetDisplayMode(640, 480, 32, 0, DDSDM_DEFAULT)
sdFront.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
sdFront.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
sdFront.lBackBufferCount = 1
Set srfFront = dxDirectDraw7.CreateSurface(sdFront)

Dim ddCaps As DDSCAPS2
ddCaps.lCaps = DDSCAPS_BACKBUFFER
Set srfBackBuffer = srfFront.GetAttachedSurface(ddCaps)
srfBackBuffer.GetSurfaceDesc sdBackBuffer


End Sub
