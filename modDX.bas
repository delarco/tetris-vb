Attribute VB_Name = "modDX"
Public DX As DirectX7
Public DDraw As DirectDraw7

Public PrimarySurf As DirectDrawSurface7
Public BackBuffer As DirectDrawSurface7
Public BackBufferStatus As DirectDrawSurface7

Dim DDSDesc As DDSURFACEDESC2
Dim Caps As DDSCAPS2
Public DDClipper As DirectDrawClipper

Public surfBlocos As DirectDrawSurface7
Public surfPeca As DirectDrawSurface7

Public rBackBuffer As RECT
Public rBackBufferStatus As RECT
Public rScr As RECT
Public rPeca As RECT

Function Init_DX() As Boolean
On Local Error GoTo Init_DX_Erro
Set DX = New DirectX7
If DX Is Nothing Then MsgBox "Error: couldn't create dx object.": Exit Function
Set DDraw = DX.DirectDrawCreate("")
If DDraw Is Nothing Then MsgBox "Error: couldn't create ddraw object.": Exit Function
'Cooperative level
DDraw.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL
'Configura DDSDesc para criar uma Primary Surface
DDSDesc.lFlags = DDSD_CAPS
DDSDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
DDSDesc.lBackBufferCount = 1
Set PrimarySurf = DDraw.CreateSurface(DDSDesc)
'Configura o Clipper
Set DDClipper = DDraw.CreateClipper(0)
DDClipper.SetHWnd frmMain.picGame.hWnd
PrimarySurf.SetClipper DDClipper
'Configura o backbuffer
DDSDesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
DDSDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
DDSDesc.lWidth = frmMain.picGame.Width
DDSDesc.lHeight = frmMain.picGame.Height
Set BackBuffer = DDraw.CreateSurface(DDSDesc)
Init_DX = True
Exit Function
Init_DX_Erro:
MsgBox "Can't initialize DirectX.", vbCritical, "Error"
End
End Function

Sub End_DX()
Set DX = Nothing
Set DDraw = Nothing
Set PrimarySurf = Nothing
Set BackBuffer = Nothing
End Sub

Function Init_Surfaces() As Boolean
On Local Error GoTo Init_Surfaces_Erro
DDSDesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
DDSDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY

DDSDesc.lWidth = 120
DDSDesc.lHeight = 20
Set surfBlocos = DDraw.CreateSurface(DDSDesc)

Dim tmpDC As Long
tmpDC = surfBlocos.GetDC
BitBlt tmpDC, 0, 0, 120, 20, frmMain.PicBlocos(1).hDC, 0, 0, &H8800C6
BitBlt tmpDC, 0, 0, 120, 20, frmMain.PicBlocos(0).hDC, 0, 0, &H660046
surfBlocos.ReleaseDC tmpDC

DDSDesc.lWidth = 80
DDSDesc.lHeight = 80
Set surfPeca = DDraw.CreateSurface(DDSDesc)

DDSDesc.lWidth = 80
DDSDesc.lHeight = 300
Set BackBufferStatus = DDraw.CreateSurface(DDSDesc)

rBackBuffer.Right = 200: rBackBuffer.Bottom = 300

rPeca.Bottom = 80: rPeca.Right = 80

rBackBufferStatus.Right = 80: rBackBufferStatus.Bottom = 300

Init_Surfaces = True
Exit Function
Init_Surfaces_Erro:
MsgBox "Error: Init_Surfaces"
End Function
